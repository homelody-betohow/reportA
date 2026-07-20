"""
K4_合并_分摊_订单统计 — 将海外仓仓租并入订单统计（月报步骤 15→16）

【流水线位置】
  上游：K3 产出「(处理完成)所有-海外仓-仓租明细.xlsx」
       + 订单统计「(已完成-15)订单统计-{日期}.xlsx」
  本脚本：按「站点商品ID识别码」合并仓租 → LM-BC 按销量分摊 → 特殊站点仓租归并
  下游：写出「(已完成-16)订单统计-…」；月报后续由 C2_ManoRent合并.py
       再把 MANO 仓租写入同一份「已完成-16」的「FBA仓租费」

【核心处理】
  1. left merge：订单统计 ← 仓租「海外仓仓租费」
  2. 仓租有、订单统计无的 SKU 行追加进结果（避免仓租丢失）
  3. 透传「所有仓库-无平台-需要分摊的费用」（仅写在第 1 行，供后续分摊脚本读）
  4. LM-BC 总仓租对半分给 -BC-ls / -BC-xj，再按各自销量占比摊到行；删除站点=LM-BC 行（避免双计）
  5. 若干「壳站点」仓租归并到同平台销量更大的真实站点（见 replace_site_with_rent）

【平台字段】
  「平台」「平台商品ID识别码」已在 K3 补全；本脚本追加仓租缺失行时一并带入，不再二次映射。
"""

from __future__ import annotations

import importlib.util
from pathlib import Path

import pandas as pd

# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT
from A_报表.A0_设置_时间段.A0_set_date import folder_name, shared_date

# 壳站点归并清单：(平台, 壳站点)
# 仓租常记在这些「壳」上；利润核算要并到同平台销量更大的真实站点。
# 顺序有意义：LM-BC-ls 必须在 LM-BC 分摊之后执行（main 里先分摊再归并）。
SHELL_SITE_MERGES: list[tuple[str, str]] = [
    ("MANO-EU", "MANO-FR"),
    ("REAL", "REAL-FB"),
    ("LM", "LM-BTH"),
    ("LM", "LM-BC-ls"),
    ("AMAZON-EU", "AMAZON-DE"),
]


# ---------------------------------------------------------------------------
# 路径 / 读入
# ---------------------------------------------------------------------------

def _order_stats_path() -> Path:
    """(已完成-15) 订单统计路径。"""
    return Path(
        DESKTOP_ROOT,
        f"{folder_name}{shared_date}",
        "订单统计",
        f"(已完成-15)订单统计-{shared_date}.xlsx",
    )


def _cang_zu_path() -> Path:
    """K3 产出的合并仓租明细路径。"""
    return Path(
        DESKTOP_ROOT,
        f"{folder_name}{shared_date}",
        "仓租",
        "(处理完成)所有-海外仓-仓租明细.xlsx",
    )


def load_inputs() -> tuple[pd.DataFrame, pd.DataFrame, Path]:
    """读取订单统计与仓租明细；返回 (订单统计, 仓租, 订单统计路径)。"""
    order_path = _order_stats_path()
    cang_zu_path = _cang_zu_path()
    order_df = pd.read_excel(order_path)
    cang_zu_df = pd.read_excel(cang_zu_path)
    return order_df, cang_zu_df, order_path


# ---------------------------------------------------------------------------
# 合并仓租 + 补缺失行 + 透传无平台分摊费用
# ---------------------------------------------------------------------------

def merge_rent_onto_orders(
    order_df: pd.DataFrame,
    cang_zu_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    按「站点商品ID识别码」把仓租挂到订单统计上。

    步骤
    ----
    1. left merge：只取仓租侧的「海外仓仓租费」
    2. 追加：仓租有、订单统计没有的识别码行（避免仓租丢失）
    3. 列对齐：结果列 = 订单统计全部列 + 「海外仓仓租费」；缺失列补 None
    4. 「海外仓仓租费」空值填 0，避免后续加减出现 NaN
    """
    result = pd.merge(
        order_df,
        cang_zu_df[["站点商品ID识别码", "海外仓仓租费"]],
        on="站点商品ID识别码",
        how="left",
    )

    # 仓租侧独有识别码：整行追加（含 K3 已补全的平台 / 平台商品ID识别码 等）
    missing = cang_zu_df[
        ~cang_zu_df["站点商品ID识别码"].isin(order_df["站点商品ID识别码"])
    ]
    result = pd.concat([result, missing], ignore_index=True)

    expected_columns = list(order_df.columns) + ["海外仓仓租费"]
    for col in expected_columns:
        if col not in result.columns:
            result[col] = None

    result["海外仓仓租费"] = result["海外仓仓租费"].fillna(0)
    return result[expected_columns]


def attach_unallocated_rent_total(
    result_df: pd.DataFrame,
    cang_zu_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    透传「所有仓库-无平台-需要分摊的费用」。

    K3 把汇总值写在仓租表第 1 行；这里原样写到结果第 1 行，
    供后续脚本（如 K6）读取做无平台费用分摊，本脚本不按行拆开。
    """
    total = cang_zu_df["所有仓库-无平台-需要分摊的费用"].iloc[0]
    out = result_df.copy()
    out["所有仓库-无平台-需要分摊的费用"] = None
    out.at[0, "所有仓库-无平台-需要分摊的费用"] = total
    return out


# ---------------------------------------------------------------------------
# LM-BC 仓租分摊
# ---------------------------------------------------------------------------

def allocate_lm_bc_rent(
    result_df: pd.DataFrame,
    cang_zu_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    将 K3 汇总的「LM-BC的仓租」摊到订单统计里的 -BC-ls / -BC-xj 行。

    背景
    ----
    仓租侧常记在「LM-BC」汇总维度，订单统计实际站点是 LM-*-BC-ls / LM-*-BC-xj。

    规则
    ----
    1. 取仓租表第 1 行的「LM-BC的仓租」总值
    2. 先对半分给两个 suffix 组：BC-ls、BC-xj（全局合计，不按 base 站点再拆）
    3. 组内按「销量」占比摊到每一行「海外仓仓租费」
       行仓租 = 半额 × (该行销量 / 该 suffix 总销量)；销量为 0 时填 0

    注意：会覆盖这些行此前 merge 得到的仓租（若有），以分摊结果为准。
    分摊成功后删除「站点=LM-BC」行（汇总池，非订单站点），避免与明细 merge/追加重复计入。
    """
    lm_bc_sum = cang_zu_df["LM-BC的仓租"].iloc[0]
    if pd.isna(lm_bc_sum):
        lm_bc_sum = 0
    out = result_df.copy()

    mask = out["站点"].str.contains(r"-BC-ls$|-BC-xj$", regex=True, na=False)
    if not mask.any():
        if lm_bc_sum:
            print(
                "[警告] 有 LM-BC的仓租 但订单统计无 -BC-ls/-BC-xj 行，"
                "未分摊，保留 LM-BC 站点行"
            )
        return out

    filtered = out.loc[mask].copy()
    filtered["suffix"] = filtered["站点"].str.extract(r"-(BC-ls|BC-xj)$", expand=False)
    filtered["total_sales_by_suffix"] = filtered.groupby("suffix")["销量"].transform("sum")

    half_amount = lm_bc_sum / 2
    filtered["海外仓仓租费"] = (
        half_amount * filtered["销量"] / filtered["total_sales_by_suffix"]
    ).fillna(0)

    out.loc[filtered.index, "海外仓仓租费"] = filtered["海外仓仓租费"]

    pool_mask = out["站点"] == "LM-BC"
    n_drop = int(pool_mask.sum())
    if n_drop:
        out = out.loc[~pool_mask].reset_index(drop=True)
        print(f"已删除 LM-BC 汇总站点行 {n_drop} 行（仓租已摊至 -BC-ls/-BC-xj）")
    return out


# ---------------------------------------------------------------------------
# 壳站点仓租归并
# ---------------------------------------------------------------------------

def _candidate_sites(df: pd.DataFrame, ping_tai: str, site: str) -> list[str]:
    """
    同平台内、可作为归并目标的候选站点（按销量合计降序）。

    匹配关键字
    ----------
    - MANO-EU：用完整 site 字符串（如 MANO-FR）
    - 其它平台：用「-」分段后的最后一段（如 REAL-FB → FB，LM-BTH → BTH）
    """
    site_end = site if ping_tai == "MANO-EU" else site.split("-")[-1]
    plat_df = df[df["平台"] == ping_tai]
    filtered = plat_df[
        plat_df["站点"].str.contains(site_end, na=False) & (plat_df["站点"] != site)
    ]
    site_sales = (
        filtered.groupby("站点", as_index=False)["销量"]
        .sum()
        .sort_values(by="销量", ascending=False)
    )
    return site_sales["站点"].tolist()


def replace_site_with_rent(
    df: pd.DataFrame,
    ping_tai: str,
    site: str,
) -> pd.DataFrame:
    """
    将「壳站点」上的仓租归并到同平台、销量更大的真实站点。

    业务场景
    --------
    某些站点（如 MANO-FR、REAL-FB、LM-BTH）在仓租/订单里存在，
    但利润核算希望把费用并到同平台下其它实际出单站点上。

    匹配键
    ------
    新识别码 = 候选站点 + 原行「商品ID」
    - 若新识别码已存在 → 把该行「海外仓仓租费」累加到已有行，并删除壳站点原行
    - 若所有候选都不存在 → 用销量最大的那个候选站点改写该行（改站点与识别码后追加），
      仍删除原行，避免仓租丢失
    - 若没有任何候选站点 → 保留原行不动（打印警告）
    """
    replacements = _candidate_sites(df, ping_tai, site)
    print(f"M{site}的站点替换顺序，按站点销量大到小排序：\n{replacements}")

    if not replacements:
        print(f"[警告] 平台 {ping_tai} 下找不到 {site} 的候选替换站点，跳过归并")
        return df

    target_rows = df[df["站点"] == site].copy()
    if target_rows.empty:
        return df

    rows_to_drop: list = []
    rows_to_append: list = []

    for idx, row in target_rows.iterrows():
        matched = False
        for new_site in replacements:
            new_id = f"{new_site}{row['商品ID']}"
            hit = df.index[df["站点商品ID识别码"] == new_id]
            if len(hit):
                # 目标识别码已存在：仓租费累加过去
                df.loc[hit[0], "海外仓仓租费"] += row["海外仓仓租费"]
                matched = True
                break

        if not matched:
            # 无一命中：落到销量最大的候选站点，改写站点/识别码后追加
            last_site = replacements[0]
            new_row = row.copy()
            new_row["站点"] = last_site
            new_row["站点商品ID识别码"] = f"{last_site}{row['商品ID']}"
            rows_to_append.append(new_row)

        # 无论是否匹配成功，壳站点原行都要删掉
        rows_to_drop.append(idx)

    df = df.drop(index=rows_to_drop)
    if rows_to_append:
        df = pd.concat([df, pd.DataFrame(rows_to_append)], ignore_index=True)
    return df


def merge_shell_sites(result_df: pd.DataFrame) -> pd.DataFrame:
    """按 SHELL_SITE_MERGES 依次做壳站点仓租归并。"""
    out = result_df
    for ping_tai, site in SHELL_SITE_MERGES:
        out = replace_site_with_rent(out, ping_tai=ping_tai, site=site)
    return out


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def main() -> None:
    # 1. 读入
    order_df, cang_zu_df, order_path = load_inputs()

    # 2～3. 合并仓租 + 补缺失行
    result_df = merge_rent_onto_orders(order_df, cang_zu_df)

    # 4. 透传无平台分摊费用（写在第 1 行）
    result_df = attach_unallocated_rent_total(result_df, cang_zu_df)

    # 5. LM-BC → -BC-ls / -BC-xj 按销量分摊
    result_df = allocate_lm_bc_rent(result_df, cang_zu_df)

    # 6. 壳站点仓租归并（含 LM-BC-ls，须在步骤 5 之后）
    result_df = merge_shell_sites(result_df)

    # 7. 写出：15 → 16；月报后续 C2_ManoRent合并 会写入「FBA仓租费」
    output_path = Path(str(order_path).replace("已完成-15", "已完成-16"))
    result_df.to_excel(output_path, index=False)
    print(f"处理完成，结果已保存到{output_path}")


if __name__ == "__main__":
    main()
