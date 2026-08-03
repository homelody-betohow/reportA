"""
K1_4_合并订单统计 — 将海外仓仓租并入订单统计（15→16）

【流水线位置】
  上游：K1_3 产出「(平台分摊)所有-海外仓-仓租明细.xlsx」
       + 订单统计「(已完成-15)订单统计-{日期}.xlsx」
  本脚本：按「站点商品ID识别码」合并仓租
  下游：写出「(已完成-16)订单统计-…」；月报后续由 C2_ManoRent合并.py
       再把 MANO 仓租写入同一份「已完成-16」的「FBA仓租费」

【核心处理】
  1. 读 K1_3 Sheet「平台分摊」；按「站点商品ID识别码」汇总「海外仓仓租费」
     （同一识别码可能对应多运营负责人行，识别码不含负责人，须先 sum）
  2. left merge：订单统计 ← 汇总后的海外仓仓租费
  3. 仓租有、订单统计无的识别码行追加进结果（避免仓租丢失）
  4. 透传无平台分摊费用（K1_3「无平台-仓租费用」→ 结果列
     「所有仓库-无平台-需要分摊的费用」，仅写在第 1 行，供后续分摊脚本读）

【说明】
  - 站点 / 运营负责人边界已在 K1_3 处理；本脚本只做识别码挂接 + 无平台透传。
  - 不做 LM-BC 对半分、不做壳站点替换。

用法：
  python modules/K_storage_fee_product_info/K1_3_合并仓租+站点分摊.py
  python modules/K_storage_fee_product_info/K1_4_合并订单统计(月报-需执行ManoRent).py
"""

from __future__ import annotations

import importlib.util
from pathlib import Path

import pandas as pd

# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.cang_zu_decimal import round_rent, round_rent_series  # noqa: E402
from config.A0_paths import DESKTOP_ROOT  # noqa: E402
from config.A0_set_date import folder_name, shared_date  # noqa: E402

# K1_3 Sheet1 列名；写出时仍用旧列名，兼容 K6 等下游
_SHEET_PLATFORM = "平台分摊"
_COL_ID = "站点商品ID识别码"
_COL_FEE = "海外仓仓租费"
_COL_NO_PLATFORM_SRC = "无平台-仓租费用"
_COL_NO_PLATFORM_OUT = "所有仓库-无平台-需要分摊的费用"
_COL_OWNER = "运营负责人"
_BLANK_IDS = frozenset({"", "nan", "None", "NaN"})


def _first_nonempty(series: pd.Series) -> str:
    for v in series:
        if pd.isna(v):
            continue
        s = str(v).strip()
        if s and s.lower() not in ("nan", "none"):
            return s
    return ""


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
    """K1_3 产出的合并仓租明细路径。"""
    return Path(
        DESKTOP_ROOT,
        f"{folder_name}{shared_date}",
        "仓租",
        "(平台分摊)所有-海外仓-仓租明细.xlsx",
    )


def load_inputs() -> tuple[pd.DataFrame, pd.DataFrame, Path]:
    """读取订单统计与仓租明细；返回 (订单统计, 仓租, 订单统计路径)。"""
    order_path = _order_stats_path()
    cang_zu_path = _cang_zu_path()
    if not order_path.is_file():
        raise FileNotFoundError(f"未找到订单统计：{order_path}")
    if not cang_zu_path.is_file():
        raise FileNotFoundError(
            f"未找到仓租明细：{cang_zu_path}\n"
            f"请先运行 K1_3_合并仓租+站点分摊.py"
        )
    order_df = pd.read_excel(order_path)
    try:
        cang_zu_df = pd.read_excel(cang_zu_path, sheet_name=_SHEET_PLATFORM)
    except ValueError:
        # 兼容旧单 Sheet 文件
        cang_zu_df = pd.read_excel(cang_zu_path, sheet_name=0)
    return order_df, cang_zu_df, order_path


def _aggregate_rent_by_site_uid(cang_zu_df: pd.DataFrame) -> pd.DataFrame:
    """
    按「站点商品ID识别码」汇总海外仓仓租费。

    K1_3 明细含运营负责人，识别码 = 站点+商品ID（不含负责人）；
    同识别码多行须先 sum，再 merge，避免订单行被笛卡尔放大。
    """
    need = [_COL_ID, _COL_FEE]
    missing_cols = [c for c in need if c not in cang_zu_df.columns]
    if missing_cols:
        raise KeyError(f"仓租表缺少列 {missing_cols}")

    work = cang_zu_df.copy()
    work[_COL_ID] = work[_COL_ID].map(
        lambda v: "" if pd.isna(v) else str(v).strip()
    )
    work[_COL_FEE] = round_rent_series(work[_COL_FEE]).fillna(0)
    # 空识别码不参与挂接（无平台汇总写在第 1 行费用列，另途透传）
    work = work.loc[~work[_COL_ID].isin(_BLANK_IDS)].copy()

    before_n = len(work)
    agg: dict[str, object] = {_COL_FEE: (_COL_FEE, "sum")}
    # 追加行时尽量保留维度列（取首个非空）
    for col in ("SKU", "商品ID", "平台", "站点", "平台商品ID识别码", _COL_OWNER):
        if col in work.columns:
            agg[col] = (col, _first_nonempty)

    out = (
        work.groupby(_COL_ID, as_index=False, dropna=False)
        .agg(**{k: v for k, v in agg.items()})
    )
    out[_COL_FEE] = round_rent_series(out[_COL_FEE])
    after_n = len(out)
    if before_n != after_n:
        print(
            f"[汇总] 仓租按「{_COL_ID}」合并：{before_n} → {after_n} 行"
            f"（同识别码多负责人已 sum）"
        )
    return out


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
    1. 仓租按识别码汇总海外仓仓租费
    2. left merge：只取汇总后的「海外仓仓租费」
    3. 追加：仓租有、订单统计没有的识别码行（避免仓租丢失）
    4. 列对齐：结果列 = 订单统计全部列 + 「海外仓仓租费」
    5. 「海外仓仓租费」空值填 0，避免后续加减出现 NaN
    """
    if _COL_ID not in order_df.columns:
        raise KeyError(f"订单统计缺少列「{_COL_ID}」")

    rent = _aggregate_rent_by_site_uid(cang_zu_df)
    rent_fee = rent[[_COL_ID, _COL_FEE]].copy()
    rent_total = float(round_rent(rent_fee[_COL_FEE].sum()))

    result = pd.merge(
        order_df,
        rent_fee,
        on=_COL_ID,
        how="left",
    )

    order_ids = set(
        order_df[_COL_ID]
        .map(lambda v: "" if pd.isna(v) else str(v).strip())
        .tolist()
    )
    missing = rent.loc[~rent[_COL_ID].isin(order_ids)].copy()
    n_append = 0
    if not missing.empty:
        n_append = len(missing)
        result = pd.concat([result, missing], ignore_index=True)
        print(f"[追加] 订单统计无匹配的仓租行 {n_append} 条（已并入结果）")

    # 结果列 = 订单全部列 + 海外仓仓租费（追加行上的平台/站点等与订单同名列会保留）
    expected_columns = list(order_df.columns)
    if _COL_FEE not in expected_columns:
        expected_columns.append(_COL_FEE)

    for col in expected_columns:
        if col not in result.columns:
            result[col] = None

    result[_COL_FEE] = round_rent_series(result[_COL_FEE]).fillna(0)

    matched_fee = float(
        round_rent(
            result.loc[
                result[_COL_ID]
                .map(lambda v: "" if pd.isna(v) else str(v).strip())
                .isin(order_ids),
                _COL_FEE,
            ].sum()
        )
    )
    append_fee = float(
        round_rent(missing[_COL_FEE].sum()) if n_append else 0.0
    )
    print(
        f"[对账] 仓租汇总合计={rent_total:.4f}；"
        f"挂到订单行={matched_fee:.4f}；追加行={append_fee:.4f}；"
        f"核对={matched_fee + append_fee:.4f}"
    )
    return result[expected_columns]


def _read_unallocated_total(cang_zu_df: pd.DataFrame) -> float:
    """从 K1_3「无平台-仓租费用」第 1 行取值；兼容旧列名。"""
    if _COL_NO_PLATFORM_SRC in cang_zu_df.columns:
        col = _COL_NO_PLATFORM_SRC
    elif _COL_NO_PLATFORM_OUT in cang_zu_df.columns:
        col = _COL_NO_PLATFORM_OUT
    else:
        print(
            f"[警告] 仓租表无「{_COL_NO_PLATFORM_SRC}」/"
            f"「{_COL_NO_PLATFORM_OUT}」，无平台费用按 0"
        )
        return 0.0
    val = cang_zu_df[col].iloc[0]
    if pd.isna(val):
        return 0.0
    return float(round_rent(val))


def attach_unallocated_rent_total(
    result_df: pd.DataFrame,
    cang_zu_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    透传无平台分摊费用到结果第 1 行。

    K1_3 把汇总值写在「无平台-仓租费用」第 1 行；这里写入下游兼容列名
    「所有仓库-无平台-需要分摊的费用」，供 K6 等读取，本脚本不按行拆开。
    """
    total = _read_unallocated_total(cang_zu_df)
    out = result_df.copy()
    out[_COL_NO_PLATFORM_OUT] = None
    if len(out) > 0:
        out.at[0, _COL_NO_PLATFORM_OUT] = total
    print(f"[无平台] {_COL_NO_PLATFORM_OUT}={total:.4f}（写在第 1 行）")
    return out


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def main() -> int:
    order_df, cang_zu_df, order_path = load_inputs()
    print(f"[读取] 订单统计 {len(order_df)} 行 ← {order_path}")
    print(
        f"[读取] 仓租 Sheet「{_SHEET_PLATFORM}」{len(cang_zu_df)} 行"
        f" ← {_cang_zu_path()}"
    )

    result_df = merge_rent_onto_orders(order_df, cang_zu_df)
    result_df = attach_unallocated_rent_total(result_df, cang_zu_df)

    # 写出：15 → 16；月报后续 C2_ManoRent合并 会写入「FBA仓租费」
    output_path = Path(str(order_path).replace("已完成-15", "已完成-16"))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    result_df.to_excel(output_path, index=False)

    fee_sum = float(
        round_rent(
            pd.to_numeric(result_df[_COL_FEE], errors="coerce").fillna(0).sum()
        )
    )
    matched = int(
        (pd.to_numeric(result_df[_COL_FEE], errors="coerce").fillna(0) != 0).sum()
    )
    print(
        f"处理完成：结果 {len(result_df)} 行，"
        f"有仓租行={matched}，海外仓仓租费合计={fee_sum:.4f}"
    )
    print(f"结果已保存到{output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
