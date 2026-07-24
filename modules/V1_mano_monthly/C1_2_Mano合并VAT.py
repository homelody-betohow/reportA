"""
C1_2_Mano合并VAT.py — 按SKU-站点识别码汇总 MANO VAT 和佣金（仅月报）

功能：
  1. 读取 仓租\\mano 目录下 (已完成-1)MANO-VAT和佣金-*.xlsx
  2. 按「SKU-站点识别码」分组（等同 VLOOKUP 汇总）
  3. quantity 与费用列求和，维度列保留每组首行
  4. 输出 (处理完成)MANO-VAT和佣金-*.xlsx
"""

import importlib.util
import warnings
from pathlib import Path

import numpy as np
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

from common.style import Color
from config.A0_set_date import folder_name, shared_date
from config.A0_paths import DESKTOP_ROOT

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

SON_SITE_ID_COL = "SKU-站点识别码"
SITE_COL = "站点"
SELLER_SKU_COL = "sellerSku"
WH_SKU_COL = "仓库SKU"
PRODUCT_UID_COL = "商品ID"

MANO_DIR = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\mano-vat"
INPUT_GLOB = "(已完成-1)MANO-VAT和佣金-*.xlsx"

# 与 C1_1 一致：分组时求和的数量/费用列
_SUM_COLUMNS = (
    "quantity",
    "amountVatIncl",
    "commissionVatIncl",
    "sellerCouponVatIncl",
    "netAmount",
    "productPriceVatExcl",
    "vatOnProduct",
    "shippingPriceVatExcl",
    "vatOnShipping",
    "amountVatExcl",
    "commissionVatExcl",
    "vatOnCommission",
    "sellerCouponVatExcl",
    "vatOnSellerCoupon",
)

# 分组时保留每组第一行的维度列
_FIRST_COLUMNS = (
    SITE_COL,
    SELLER_SKU_COL,
    WH_SKU_COL,
    PRODUCT_UID_COL,
    "type",
    "currency",
)


def _strip_df_strings(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
    return df


def _sum_columns_in_df(df: pd.DataFrame) -> list[str]:
    return [c for c in _SUM_COLUMNS if c in df.columns]


def _first_columns_in_df(df: pd.DataFrame) -> list[str]:
    return [c for c in _FIRST_COLUMNS if c in df.columns]


def _output_name(input_name: str) -> str:
    if input_name.startswith("(已完成-1)"):
        return "(处理完成)" + input_name[len("(已完成-1)") :]
    return f"(处理完成){input_name}"


def _group_by_son_site_id(df: pd.DataFrame) -> pd.DataFrame:
    if SON_SITE_ID_COL not in df.columns:
        raise KeyError(f"未找到列 {SON_SITE_ID_COL!r}，当前列名: {df.columns.tolist()}")

    df = df[df[SON_SITE_ID_COL].notna() & (df[SON_SITE_ID_COL].astype(str).str.strip() != "")].copy()
    if df.empty:
        raise ValueError("无有效「SKU-站点识别码」数据，无法汇总")

    sum_cols = _sum_columns_in_df(df)
    first_cols = _first_columns_in_df(df)
    if not sum_cols:
        raise ValueError(f"未找到可汇总的数量/费用列，当前列名: {df.columns.tolist()}")

    agg: dict[str, str] = {c: "sum" for c in sum_cols}
    for c in first_cols:
        agg[c] = "first"

    grouped = df.groupby(SON_SITE_ID_COL, as_index=False).agg(agg)

    for col in sum_cols:
        grouped[col] = pd.to_numeric(grouped[col], errors="coerce").fillna(0)
        grouped[col] = np.round(grouped[col], 2)

    # 输出列：维度列靠前，数量/费用列随后（保持与源表相近的可读顺序）
    ordered_cols = [SON_SITE_ID_COL]
    for c in first_cols:
        if c not in ordered_cols:
            ordered_cols.append(c)
    for c in sum_cols:
        if c not in ordered_cols:
            ordered_cols.append(c)
    return grouped[ordered_cols]


def _save_excel(df: pd.DataFrame, output_path: str) -> str:
    try:
        df.to_excel(output_path, index=False)
        return output_path
    except PermissionError:
        alt = output_path.replace(".xlsx", "-另存.xlsx")
        df.to_excel(alt, index=False)
        return alt


def _list_input_files(mano_dir: Path) -> list[Path]:
    files = sorted(mano_dir.glob(INPUT_GLOB))
    return [p for p in files if p.is_file() and not p.name.startswith("(处理完成)")]


def process_one_file(file_path: Path) -> str:
    df = pd.read_excel(file_path)
    df = _strip_df_strings(df)
    before_rows = len(df)

    grouped = _group_by_son_site_id(df)
    after_rows = len(grouped)

    output_path = file_path.parent / _output_name(file_path.name)
    saved = _save_excel(grouped, str(output_path))

    print(
        f"{Color.GREEN}[汇总]{Color.RESET} {file_path.name}："
        f"{before_rows} 行 → {after_rows} 行（按SKU-站点识别码）"
    )
    return saved


def main() -> None:
    if folder_name != "月报":
        print(f"{Color.YELLOW}[跳过] C1_2_Mano合并VAT 仅月报执行，当前 folder_name={folder_name!r}{Color.RESET}")
        return

    mano_dir = Path(MANO_DIR)
    if not mano_dir.is_dir():
        raise FileNotFoundError(f"未找到 MANO 目录：{mano_dir}")

    input_files = _list_input_files(mano_dir)
    if not input_files:
        raise FileNotFoundError(
            f"目录下未找到待处理文件：{mano_dir}\\{INPUT_GLOB}"
        )

    print(f"{Color.CYAN}[C1_2_Mano合并VAT] 月报 {shared_date}，共 {len(input_files)} 个文件{Color.RESET}")
    for file_path in input_files:
        saved = process_one_file(file_path)
        print(f"处理完成，文件另存为：{saved}")

    print(f"{Color.GREEN}一切正常，请进行下一步操作{Color.RESET}")


if __name__ == "__main__":
    main()
