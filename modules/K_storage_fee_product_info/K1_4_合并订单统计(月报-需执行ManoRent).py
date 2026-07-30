"""
K1_4_合并订单统计 — 将海外仓仓租并入订单统计（月报步骤 15→16）

【流水线位置】
  上游：K1_3 产出「(平台分摊)所有-海外仓-仓租明细.xlsx」
       + 订单统计「(已完成-15)订单统计-{日期}.xlsx」
  本脚本：按「站点商品ID识别码」合并仓租
  下游：写出「(已完成-16)订单统计-…」；月报后续由 C2_ManoRent合并.py
       再把 MANO 仓租写入同一份「已完成-16」的「FBA仓租费」

【核心处理】
  1. left merge：订单统计 ← 仓租「海外仓仓租费」
  2. 仓租有、订单统计无的识别码行追加进结果（避免仓租丢失）
  3. 透传无平台分摊费用（K1_3「无平台-仓租费用」→ 结果列
     「所有仓库-无平台-需要分摊的费用」，仅写在第 1 行，供后续分摊脚本读）

【相对旧 K4 不再做的事】
  - 不做 LM-BC 总仓租对半分（K1_3 已按站点销量阶梯分摊）
  - 不做壳站点替换（MANO-FR / AMAZON-DE / LM-BTH …）
    原因：旧链路仓租先落到 PLATFORM_TO_SITE「壳站点」，再靠 K4 归并到真实店铺站点；
    K1_3 已直接按订单统计「站点」分摊，识别码应与 (已完成-15) 对齐，
    再替换会误删壳站点订单行并二次挪费。

【平台字段】
  「平台」「平台商品ID识别码」已在 K1_3 写出；本脚本追加仓租缺失行时一并带入，不再二次映射。

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

from config.A0_paths import DESKTOP_ROOT  # noqa: E402
from config.A0_set_date import folder_name, shared_date  # noqa: E402

# K1_3 Sheet1 列名；写出时仍用旧列名，兼容 K6 等下游
_COL_NO_PLATFORM_SRC = "无平台-仓租费用"
_COL_NO_PLATFORM_OUT = "所有仓库-无平台-需要分摊的费用"


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
    # K1_3 多 Sheet：默认读第一个「平台分摊」
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
    need = ["站点商品ID识别码", "海外仓仓租费"]
    missing_cols = [c for c in need if c not in cang_zu_df.columns]
    if missing_cols:
        raise KeyError(f"仓租表缺少列 {missing_cols}")
    if "站点商品ID识别码" not in order_df.columns:
        raise KeyError("订单统计缺少列「站点商品ID识别码」")

    result = pd.merge(
        order_df,
        cang_zu_df[need],
        on="站点商品ID识别码",
        how="left",
    )

    # 仓租侧独有识别码：整行追加（含 K1_3 已写出的平台 / 平台商品ID识别码 等）
    missing = cang_zu_df[
        ~cang_zu_df["站点商品ID识别码"].isin(order_df["站点商品ID识别码"])
    ]
    # 空识别码不追加（无平台汇总行等）
    if not missing.empty:
        blank_id = (
            missing["站点商品ID识别码"].isna()
            | missing["站点商品ID识别码"].astype(str).str.strip().isin(
                ["", "nan", "None"]
            )
        )
        missing = missing.loc[~blank_id]
    if not missing.empty:
        result = pd.concat([result, missing], ignore_index=True)
        print(f"[追加] 订单统计无匹配的仓租行 {len(missing)} 条（已并入结果）")

    expected_columns = list(order_df.columns) + ["海外仓仓租费"]
    for col in expected_columns:
        if col not in result.columns:
            result[col] = None

    result["海外仓仓租费"] = result["海外仓仓租费"].fillna(0)
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
    return float(val)


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
    print(f"[读取] 仓租明细 {len(cang_zu_df)} 行 ← {_cang_zu_path()}")

    result_df = merge_rent_onto_orders(order_df, cang_zu_df)
    result_df = attach_unallocated_rent_total(result_df, cang_zu_df)

    # 写出：15 → 16；月报后续 C2_ManoRent合并 会写入「FBA仓租费」
    output_path = Path(str(order_path).replace("已完成-15", "已完成-16"))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    result_df.to_excel(output_path, index=False)

    fee_sum = float(
        pd.to_numeric(result_df["海外仓仓租费"], errors="coerce").fillna(0).sum()
    )
    matched = int(
        (pd.to_numeric(result_df["海外仓仓租费"], errors="coerce").fillna(0) != 0).sum()
    )
    print(
        f"处理完成：结果 {len(result_df)} 行，"
        f"有仓租行={matched}，海外仓仓租费合计={fee_sum:.4f}"
    )
    print(f"结果已保存到{output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
