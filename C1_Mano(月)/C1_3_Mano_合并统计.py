"""
C1_3_Mano_合并统计.py — 将 MANO 实际 VAT/佣金写入订单统计（仅月报）

功能：
  1. 读取 (处理完成)MANO-VAT和佣金-*.xlsx
  2. 读取 C5 生成的 (已完成-8)订单统计-{shared_date}.xlsx
  3. 先将 平台=MANO-EU 的 平台费(非AMZ)、销售税(非AMZ) 置 0
  4. 按「SKU-站点识别码」匹配：有则更新，无则新增行
     commissionVatIncl → 平台费(非AMZ)
     vatOnProduct      → 销售税(非AMZ)
     新增行销量统一为 0（不以 VAT quantity 回填）
"""

import importlib.util
import warnings
from pathlib import Path

import numpy as np
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

from A_报表.Z_method.style import Color
from A_报表.Z_method.platform_shop import map_region_to_platform
from A_报表.A0_设置_时间段.A0_set_date import folder_name, shared_date
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

SON_SITE_ID_COL = "SKU-站点识别码"
SON_PLATFORM_ID_COL = "SKU-平台识别码"
SITE_COL = "站点"
SKU_COL = "SKU"
PLATFORM_COL = "平台"
MANO_PLATFORM = "MANO-EU"
FEE_COL = "平台费(非AMZ)"
TAX_COL = "销售税(非AMZ)"
FEE_TOTAL_COL = "平台费合计"
TAX_TOTAL_COL = "销售税合计"
FEE_AMZ_COL = "平台费(AMZ)"
TAX_AMZ_COL = "销售税(AMZ)"
QTY_COL = "销量"

MANO_VAT_GLOB = "(处理完成)MANO-VAT和佣金-*.xlsx"
MANO_VAT_DIR = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\mano-vat"
ORDER_STATS_DIR = Path(fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计")

originalBakFile = '(已完成-8)订单统计-{shared_date}-original.xlsx'
ORDER_STATS_PATH = ORDER_STATS_DIR / f"(已完成-8)订单统计-{shared_date}.xlsx"
originalBakPath = ORDER_STATS_DIR / f"(已完成-8)订单统计-{shared_date}-original.xlsx"

# 若备份不存在，则创建备份
if not Path(originalBakPath).exists():
    pd.read_excel(ORDER_STATS_PATH).to_excel(originalBakPath, index=False)


VAT_COMMISSION_SRC = "commissionVatIncl"
VAT_TAX_SRC = "vatOnProduct"
VAT_WH_SKU_COL = "仓库SKU"


def _strip_df_strings(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
    return df


def _round_money(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    for col in cols:
        if col in df.columns:
            df[col] = np.round(pd.to_numeric(df[col], errors="coerce"), 2)
    return df


def _load_mano_vat_files(mano_dir: Path) -> pd.DataFrame:
    files = sorted(mano_dir.glob(MANO_VAT_GLOB))
    if not files:
        raise FileNotFoundError(f"未找到 MANO VAT 文件：{mano_dir}\\{MANO_VAT_GLOB}")

    parts = []
    for fp in files:
        df = _strip_df_strings(pd.read_excel(fp))
        parts.append(df)
        print(f"{Color.CYAN}[读取]{Color.RESET} {fp.name}（{len(df)} 行）")

    vat_df = pd.concat(parts, ignore_index=True)
    vat_df = vat_df[
        vat_df[SON_SITE_ID_COL].notna()
        & (vat_df[SON_SITE_ID_COL].astype(str).str.strip() != "")
    ].copy()

    dup_cnt = int(vat_df[SON_SITE_ID_COL].duplicated().sum())
    if dup_cnt:
        raise ValueError(
            f"MANO VAT 中「SKU-站点识别码」存在重复 {dup_cnt} 条，请先重跑 C1_2 汇总"
        )
    return vat_df


def _recalc_totals(df: pd.DataFrame, mask: pd.Series | None = None) -> pd.DataFrame:
    if mask is None:
        mask = pd.Series(True, index=df.index)
    fee_amz = pd.to_numeric(df.loc[mask, FEE_AMZ_COL], errors="coerce").fillna(0)
    tax_amz = pd.to_numeric(df.loc[mask, TAX_AMZ_COL], errors="coerce").fillna(0)
    fee = pd.to_numeric(df.loc[mask, FEE_COL], errors="coerce").fillna(0)
    tax = pd.to_numeric(df.loc[mask, TAX_COL], errors="coerce").fillna(0)
    df.loc[mask, FEE_TOTAL_COL] = np.round(fee_amz + fee, 2)
    df.loc[mask, TAX_TOTAL_COL] = np.round(tax_amz + tax, 2)
    return df


def _reset_mano_eu_fees(order_df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """C5 估算的 MANO 平台费/VAT 先清零，再由 MANO VAT 报表回填实际值。"""
    if PLATFORM_COL not in order_df.columns:
        raise KeyError(f"订单统计缺少列 {PLATFORM_COL!r}，请先运行 C5")

    mano_mask = order_df[PLATFORM_COL].astype(str).str.strip() == MANO_PLATFORM
    reset_cnt = int(mano_mask.sum())
    if reset_cnt:
        order_df.loc[mano_mask, FEE_COL] = 0
        order_df.loc[mano_mask, TAX_COL] = 0
        order_df = _recalc_totals(order_df, mano_mask)
        print(
            f"{Color.YELLOW}[清零]{Color.RESET} {MANO_PLATFORM}："
            f"{reset_cnt} 行 {FEE_COL}、{TAX_COL} 已置 0"
        )
    return order_df, reset_cnt


def _apply_vat_updates(order_df: pd.DataFrame, vat_df: pd.DataFrame) -> tuple[pd.DataFrame, int, int]:
    """按SKU-站点识别码更新已有行，返回 (order_df, 更新行数, 待新增 vat 行)。"""
    lookup = vat_df.set_index(SON_SITE_ID_COL)
    match_mask = order_df[SON_SITE_ID_COL].isin(lookup.index)
    update_cnt = int(match_mask.sum())

    if update_cnt:
        keys = order_df.loc[match_mask, SON_SITE_ID_COL]
        order_df.loc[match_mask, FEE_COL] = keys.map(lookup[VAT_COMMISSION_SRC])
        order_df.loc[match_mask, TAX_COL] = keys.map(lookup[VAT_TAX_SRC])
        order_df = _recalc_totals(order_df, match_mask)

    new_vat = vat_df[~vat_df[SON_SITE_ID_COL].isin(order_df[SON_SITE_ID_COL])].copy()
    return order_df, update_cnt, len(new_vat)


def _build_new_rows(vat_new: pd.DataFrame, order_columns: list[str]) -> pd.DataFrame:
    """将 MANO VAT 中订单统计不存在的识别码，构造成新行。"""
    if vat_new.empty:
        return pd.DataFrame(columns=order_columns)

    base = pd.DataFrame({col: np.nan for col in order_columns}, index=range(len(vat_new)))
    base[SON_SITE_ID_COL] = vat_new[SON_SITE_ID_COL].values
    base[SITE_COL] = vat_new[SITE_COL].values if SITE_COL in vat_new.columns else np.nan
    base[SKU_COL] = vat_new[VAT_WH_SKU_COL].values if VAT_WH_SKU_COL in vat_new.columns else np.nan
    base[FEE_COL] = pd.to_numeric(vat_new[VAT_COMMISSION_SRC], errors="coerce").values
    base[TAX_COL] = pd.to_numeric(vat_new[VAT_TAX_SRC], errors="coerce").values
    base[QTY_COL] = 0

    # 数值列默认 0，避免新增行出现空值
    numeric_defaults = [
        "平台销售额", "头程", "关税", "派送费", "重发数量", "订单采购成本", "重发采购成本",
        "退款额", "退款数量", "销售额", FEE_AMZ_COL, TAX_AMZ_COL, "提现费", "映射佣金比",
    ]
    for col in numeric_defaults:
        if col in base.columns:
            base[col] = pd.to_numeric(base[col], errors="coerce").fillna(0)

    # 站点 → 平台（数据源：platform_shop）
    mapped = map_region_to_platform(base, site_col=SITE_COL)
    base[PLATFORM_COL] = mapped["映射平台"]
    base[SON_PLATFORM_ID_COL] = base[PLATFORM_COL].astype(str) + base[SKU_COL].astype(str)
    base = _recalc_totals(base)
    return base[order_columns]


def _save_excel(df: pd.DataFrame, output_path: str) -> str:
    try:
        df.to_excel(output_path, index=False)
        return output_path
    except PermissionError:
        alt = output_path.replace(".xlsx", "-另存.xlsx")
        df.to_excel(alt, index=False)
        return alt


def main() -> None:
    if folder_name != "月报":
        print(f"{Color.YELLOW}[跳过] C1_3_Mano_合并统计 仅月报执行，当前 folder_name={folder_name!r}{Color.RESET}")
        return

    order_path = Path(ORDER_STATS_PATH)
    if not order_path.is_file():
        raise FileNotFoundError(f"未找到订单统计文件：{order_path}")

    mano_dir = Path(MANO_VAT_DIR)
    vat_df = _load_mano_vat_files(mano_dir)
    order_df = _strip_df_strings(pd.read_excel(order_path))
    order_columns = list(order_df.columns)

    for col in (FEE_COL, TAX_COL, FEE_TOTAL_COL, TAX_TOTAL_COL):
        if col not in order_df.columns:
            raise KeyError(f"订单统计缺少列 {col!r}，请先运行 C5")

    order_df, reset_cnt = _reset_mano_eu_fees(order_df)
    order_df, update_cnt, new_cnt = _apply_vat_updates(order_df, vat_df)

    if new_cnt:
        new_rows = _build_new_rows(
            vat_df[~vat_df[SON_SITE_ID_COL].isin(order_df[SON_SITE_ID_COL])],
            order_columns,
        )
        order_df = pd.concat([order_df, new_rows], ignore_index=True)
        print(f"{Color.GREEN}[新增]{Color.RESET} {new_cnt} 行（MANO VAT 有、订单统计无）")

    order_df = _round_money(
        order_df,
        [FEE_COL, TAX_COL, FEE_TOTAL_COL, TAX_TOTAL_COL],
    )

    saved = _save_excel(order_df, str(order_path))
    print(
        f"{Color.GREEN}[更新]{Color.RESET} 已写回订单统计："
        f"清零 {reset_cnt} 行，更新 {update_cnt} 行，新增 {new_cnt} 行，合计 {len(order_df)} 行"
    )
    print(f"处理完成，output_path：{saved}")
    print(f"{Color.GREEN}一切正常，请进行下一步操作（如 D6 合并广告费）{Color.RESET}")


if __name__ == "__main__":
    main()
