"""
G2_1_Temu罚款.py — 合并 TEMU 罚款 Excel 中「支出-*」工作表

功能：
  1. 读取 TEMU-罚款 目录下所有 *罚款*.xlsx 文件
  2. 合并各文件中工作表名以「支出-」开头的数据
  3. 从文件名「XXX-罚款」提取店铺名，写入「店铺」列
  4. 原工作表名写入「支出类型」列
  5. 合并时统一订单编号（违规单号并入后删除）
  6. 按 A0_set_date 汇率换算欧元，新增「结算金额」「结算币种」
  7. 输出汇总 Excel：TEMU-罚款汇总-{shared_date}.xlsx
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
from config.A0_set_date import (
    CAD_to_EUR,
    Ft_to_EUR,
    Lei_to_EUR,
    RMB_di_EUR,
    USD_to_EUR,
    folder_name,
    kc_to_EUR,
    kr_to_EUR,
    shared_date,
    zl_to_EUR,
)
from config.A0_paths import DESKTOP_ROOT

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

ORDER_NO_COL = "订单编号"
VIOLATION_NO_COL = "违规单号"
AMOUNT_COL = "支出金额"
CURRENCY_COL = "币种"
SETTLE_AMOUNT_COL = "结算金额"
SETTLE_CURRENCY_COL = "结算币种"
SETTLE_CURRENCY = "EUR"
SHOP_COL = "店铺"
EXPENSE_TYPE_COL = "支出类型"
SHOP_NAME_MARKER = "-罚款"
SHEET_PREFIX = "支出-"
FILE_GLOB = "*罚款*.xlsx"
OUTPUT_NAME = f"TEMU-罚款汇总-{shared_date}.xlsx"

# 各币种 → 欧元 乘数（CNY/RMB 为除以 RMB_di_EUR）
_CURRENCY_TO_EUR_RATE: dict[str, float] = {
    "EUR": 1.0,
    "USD": USD_to_EUR,
    "CNY": 1 / RMB_di_EUR,
    "RMB": 1 / RMB_di_EUR,
    "CAD": CAD_to_EUR,
    "CZK": kc_to_EUR,
    "PLN": zl_to_EUR,
    "HUF": Ft_to_EUR,
    "RON": Lei_to_EUR,
    "SEK": kr_to_EUR,
}

TEMU_FINE_DIR = Path(fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\TEMU-罚款")


def _strip_df_strings(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
    return df


def _is_expense_sheet(sheet_name: str) -> bool:
    return str(sheet_name).startswith(SHEET_PREFIX)


def _extract_shop_name(xlsx_path: Path) -> str:
    """从文件名 XXX-罚款-*.xlsx 中提取店铺名 XXX。"""
    stem = xlsx_path.stem
    if SHOP_NAME_MARKER not in stem:
        raise ValueError(
            f"文件名 {xlsx_path.name!r} 不符合「XXX{SHOP_NAME_MARKER}」格式，无法提取店铺名"
        )
    return stem.split(SHOP_NAME_MARKER, 1)[0]


def _normalize_order_no(df: pd.DataFrame) -> pd.DataFrame:
    """不同工作表列名不一致：违规单号并入订单编号后删除违规单号列。"""
    has_order = ORDER_NO_COL in df.columns
    has_violation = VIOLATION_NO_COL in df.columns

    if has_order and has_violation:
        df[ORDER_NO_COL] = df[ORDER_NO_COL].fillna(df[VIOLATION_NO_COL])
    elif has_violation and not has_order:
        df[ORDER_NO_COL] = df[VIOLATION_NO_COL]

    if VIOLATION_NO_COL in df.columns:
        df = df.drop(columns=[VIOLATION_NO_COL])
    return df


def _to_eur(amount, currency) -> float | None:
    if pd.isna(amount):
        return None
    if pd.isna(currency):
        raise ValueError(f"支出金额 {amount!r} 缺少币种，无法换算欧元")

    rate = _CURRENCY_TO_EUR_RATE.get(str(currency).strip().upper())
    if rate is None:
        supported = ", ".join(sorted(_CURRENCY_TO_EUR_RATE))
        raise ValueError(f"未知币种 {currency!r}，请在 A0_set_date 中补充汇率。当前支持：{supported}")
    return float(np.round(float(amount) * rate, 2))


def _add_settlement_columns(df: pd.DataFrame) -> pd.DataFrame:
    if AMOUNT_COL not in df.columns:
        raise KeyError(f"缺少列 {AMOUNT_COL!r}")
    if CURRENCY_COL not in df.columns:
        raise KeyError(f"缺少列 {CURRENCY_COL!r}")

    df[SETTLE_AMOUNT_COL] = df.apply(
        lambda row: _to_eur(row[AMOUNT_COL], row[CURRENCY_COL]), axis=1
    )
    df[SETTLE_CURRENCY_COL] = SETTLE_CURRENCY
    return df


def _order_output_columns(df: pd.DataFrame) -> pd.DataFrame:
    front_cols = [SHOP_COL, EXPENSE_TYPE_COL]
    tail_cols = [SETTLE_AMOUNT_COL, SETTLE_CURRENCY_COL]
    middle_cols = [
        c for c in df.columns if c not in front_cols + tail_cols
    ]
    return df[front_cols + middle_cols + tail_cols]


def _collect_fine_files(fine_dir: Path, output_path: Path) -> list[Path]:
    files = sorted(
        p
        for p in fine_dir.glob(FILE_GLOB)
        if p.is_file()
        and not p.name.startswith("~$")
        and p.resolve() != output_path.resolve()
        and not p.name.startswith("(处理完成)")
    )
    return files


def _read_expense_sheets(xlsx_path: Path) -> list[pd.DataFrame]:
    frames: list[pd.DataFrame] = []
    try:
        xls = pd.ExcelFile(xlsx_path, engine="openpyxl")
    except PermissionError:
        print(
            f"{Color.YELLOW}[跳过]{Color.RESET} {xlsx_path.name}："
            f"文件被占用，请先关闭 Excel 后重试"
        )
        return frames

    shop_name = _extract_shop_name(xlsx_path)

    with xls:
        expense_sheets = [name for name in xls.sheet_names if _is_expense_sheet(name)]
        if not expense_sheets:
            print(
                f"{Color.YELLOW}[跳过]{Color.RESET} {xlsx_path.name}："
                f"未找到以「{SHEET_PREFIX}」开头的工作表"
            )
            return frames

        for sheet_name in expense_sheets:
            df = pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl")
            df = df.dropna(how="all")
            if df.empty:
                print(
                    f"{Color.YELLOW}[跳过]{Color.RESET} {xlsx_path.name} / {sheet_name}：空表"
                )
                continue
            df = _strip_df_strings(df)
            df = _normalize_order_no(df)
            df.insert(0, EXPENSE_TYPE_COL, sheet_name)
            df.insert(0, SHOP_COL, shop_name)
            frames.append(df)
            print(
                f"{Color.CYAN}[读取]{Color.RESET} {xlsx_path.name} / {sheet_name}：{len(df)} 行"
            )
    return frames


def _save_excel(df: pd.DataFrame, output_path: Path) -> Path:
    try:
        df.to_excel(output_path, index=False, engine="openpyxl")
        return output_path
    except PermissionError:
        alt = output_path.with_name(output_path.stem + "-另存.xlsx")
        df.to_excel(alt, index=False, engine="openpyxl")
        print(
            f"{Color.YELLOW}[提示]{Color.RESET} 目标文件被占用，已另存为：{alt}"
        )
        return alt


def main() -> None:
    if not TEMU_FINE_DIR.is_dir():
        raise FileNotFoundError(f"未找到目录：{TEMU_FINE_DIR}")

    output_path = TEMU_FINE_DIR / OUTPUT_NAME
    fine_files = _collect_fine_files(TEMU_FINE_DIR, output_path)
    if not fine_files:
        raise FileNotFoundError(
            f"目录 {TEMU_FINE_DIR} 下未找到匹配 {FILE_GLOB!r} 的文件"
        )

    print(
        f"{Color.CYAN}[扫描]{Color.RESET} 目录：{TEMU_FINE_DIR}，"
        f"共 {len(fine_files)} 个罚款文件"
    )

    all_frames: list[pd.DataFrame] = []
    for file_path in fine_files:
        all_frames.extend(_read_expense_sheets(file_path))

    if not all_frames:
        raise ValueError("未读取到任何「支出-*」工作表数据，请检查源文件")

    merged_df = pd.concat(all_frames, ignore_index=True, sort=False)
    merged_df = _normalize_order_no(merged_df)
    merged_df = _add_settlement_columns(merged_df)
    merged_df = _order_output_columns(merged_df)

    saved = _save_excel(merged_df, output_path)
    print(
        f"{Color.GREEN}[完成]{Color.RESET} 汇总 {len(merged_df)} 行，"
        f"来自 {len(all_frames)} 个工作表"
    )
    print(f"处理完成，output_path：{saved}")


if __name__ == "__main__":
    main()
