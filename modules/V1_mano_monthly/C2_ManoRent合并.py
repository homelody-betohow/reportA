"""
C2_ManoRent合并.py — MANO MMF 仓租按商品ID识别码汇总，并回填订单统计（仅月报）

功能：
  1. 读取 C2_ManoRent 生成的 ALL-WarehouseRent.xlsx
  2. 按「商品ID识别码」分组，对 GROSS_AMOUNT_VAT_EXC 求和
  3. 将汇总结果写入同一文件的新工作表「按商品ID识别码汇总」
  4. 将汇总金额写入 K4 生成的 (已完成-16)订单统计 的「FBA仓租费」列
     （匹配键：仓租表「站点」+「商品ID」= 订单统计「站点」+「商品ID」；匹配行直接覆盖，不与原值相加）
""" 

import importlib.util
import warnings
from pathlib import Path
import shutil
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
from common.platform_shop import map_region_to_platform
from config.A0_set_date import folder_name, shared_date
from config.A0_paths import DESKTOP_ROOT

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

PRODUCT_SITE_ID_COL = "商品ID识别码"
SITE_PRODUCT_ID_COL = "站点商品ID识别码"
GROSS_AMOUNT_COL = "GROSS_AMOUNT_VAT_EXC"
RENT_COL = "FBA仓租费"
SITE_COL = "站点"
SKU_COL = "SKU"
PRODUCT_UID_COL = "商品ID"
PLATFORM_COL = "平台"
PLATFORM_PRODUCT_ID_COL = "平台商品ID识别码"
SON_PLATFORM_ID_COL = "SKU-平台识别码"


MANO_DIR = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\mano"
ORDER_STATS_DIR = Path(fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计")
RENT_FILE_NAME = "ALL-WarehouseRent.xlsx"
SUMMARY_SHEET = "按商品ID识别码汇总"
ORDER_STATS_PATH = ORDER_STATS_DIR / f"(已完成-16)订单统计-{shared_date}.xlsx"
ORIGINAL_ORDER_STATS_PATH = ORDER_STATS_DIR / f"(已完成-16)订单统计-{shared_date}-original.xlsx"


def _norm_str(val) -> str:
    return str(val).strip() if pd.notna(val) else ""


def _strip_df_strings(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
    return df


def _find_gross_amount_col(df: pd.DataFrame) -> str:
    if GROSS_AMOUNT_COL in df.columns:
        return GROSS_AMOUNT_COL
    for col in df.columns:
        if str(col).strip().upper().startswith(GROSS_AMOUNT_COL):
            return col
    raise KeyError(
        f"未找到 {GROSS_AMOUNT_COL!r} 列，当前列名: {df.columns.tolist()}"
    )


def _group_rent_by_site_product(df: pd.DataFrame) -> pd.DataFrame:
    """按「站点」+「商品ID」分组汇总仓租，并生成商品ID识别码。"""
    for col in (SITE_COL, PRODUCT_UID_COL):
        if col not in df.columns:
            raise KeyError(f"仓租明细缺少列 {col!r}")

    gross_col = _find_gross_amount_col(df)
    work = df.copy()
    work[SITE_COL] = work[SITE_COL].map(_norm_str)
    work[PRODUCT_UID_COL] = work[PRODUCT_UID_COL].map(_norm_str)
    work = work[(work[SITE_COL] != "") & (work[PRODUCT_UID_COL] != "")].copy()
    work[gross_col] = pd.to_numeric(work[gross_col], errors="coerce").fillna(0)
    work[PRODUCT_SITE_ID_COL] = work[SITE_COL] + work[PRODUCT_UID_COL]

    agg_map: dict[str, str] = {gross_col: "sum"}
    for col in (SKU_COL,):
        if col in work.columns:
            agg_map[col] = "first"

    grouped = work.groupby([SITE_COL, PRODUCT_UID_COL], as_index=False).agg(agg_map)
    grouped = grouped.rename(columns={gross_col: RENT_COL})
    grouped[RENT_COL] = np.round(grouped[RENT_COL], 2)
    grouped[PRODUCT_SITE_ID_COL] = grouped[SITE_COL] + grouped[PRODUCT_UID_COL]
    return grouped


def _append_summary_sheet(rent_path: Path, summary_df: pd.DataFrame) -> None:
    with pd.ExcelWriter(
        rent_path,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace",
    ) as writer:
        summary_df.to_excel(writer, sheet_name=SUMMARY_SHEET, index=False)


def _ensure_original_order_stats_backup(order_path: Path) -> Path:
    original_path = Path(ORIGINAL_ORDER_STATS_PATH)
    if not original_path.exists():
        pd.read_excel(order_path).to_excel(original_path, index=False)
        print(f"{Color.GREEN}[备份]{Color.RESET} 已保存：{original_path}")
    else:
        # 将备份的copy一份到订单统计目录下
        shutil.copy(original_path, order_path)
        print(f"{Color.GREEN}[备份]{Color.RESET} 已复制到订单统计目录：{order_path}")
    return original_path


def _fill_platform_from_site(df: pd.DataFrame, mask: pd.Series) -> pd.DataFrame:
    """站点有值但平台为空时，按站点映射平台，并补全平台商品ID识别码 / SKU-平台识别码。"""
    if PLATFORM_COL not in df.columns or not mask.any():
        return df

    empty_platform = df[PLATFORM_COL].isna() | (df[PLATFORM_COL].astype(str).str.strip() == "")
    target = mask & empty_platform
    if not target.any():
        return df

    mapped = map_region_to_platform(df.loc[target], site_col=SITE_COL)
    df.loc[target, PLATFORM_COL] = mapped["映射平台"]
    if PLATFORM_PRODUCT_ID_COL in df.columns and PRODUCT_UID_COL in df.columns:
        df.loc[target, PLATFORM_PRODUCT_ID_COL] = (
            df.loc[target, PLATFORM_COL].astype(str)
            + df.loc[target, PRODUCT_UID_COL].astype(str)
        )
    if SON_PLATFORM_ID_COL in df.columns and SKU_COL in df.columns:
        df.loc[target, SON_PLATFORM_ID_COL] = (
            df.loc[target, PLATFORM_COL].astype(str) + df.loc[target, SKU_COL].astype(str)
        )
    return df


def _build_appended_rows(missing_rent: pd.DataFrame, order_df: pd.DataFrame) -> pd.DataFrame:
    """将仓租汇总中订单统计不存在的「站点+商品ID」构造成新行。"""
    order_columns = list(order_df.columns)
    new_rows = pd.DataFrame({col: np.nan for col in order_columns}, index=range(len(missing_rent)))
    new_rows[SITE_COL] = missing_rent[SITE_COL].values
    new_rows[PRODUCT_UID_COL] = missing_rent[PRODUCT_UID_COL].values
    new_rows[SITE_PRODUCT_ID_COL] = missing_rent[PRODUCT_SITE_ID_COL].values
    new_rows[RENT_COL] = missing_rent[RENT_COL].values

    if SKU_COL in missing_rent.columns and SKU_COL in new_rows.columns:
        new_rows[SKU_COL] = missing_rent[SKU_COL].values

    filled_cols = {
        SITE_COL,
        PRODUCT_UID_COL,
        SITE_PRODUCT_ID_COL,
        RENT_COL,
        SKU_COL,
        PLATFORM_COL,
        PLATFORM_PRODUCT_ID_COL,
        SON_PLATFORM_ID_COL,
    }
    for col in order_columns:
        if col not in filled_cols and order_df[col].dtype.kind in "biufc":
            new_rows[col] = 0

    new_rows = _fill_platform_from_site(new_rows, pd.Series(True, index=new_rows.index))

    # 平台已有值时补全「平台商品ID识别码 = 平台 + 商品ID」
    if PLATFORM_PRODUCT_ID_COL in new_rows.columns:
        has_plat = (
            new_rows[PLATFORM_COL].notna()
            & (new_rows[PLATFORM_COL].astype(str).str.strip() != "")
            & (new_rows[PLATFORM_COL].astype(str).str.strip() != "nan")
        )
        new_rows.loc[has_plat, PLATFORM_PRODUCT_ID_COL] = (
            new_rows.loc[has_plat, PLATFORM_COL].astype(str).str.strip()
            + new_rows.loc[has_plat, PRODUCT_UID_COL].astype(str).str.strip()
        )
    return new_rows


def _order_match_keys(order_df: pd.DataFrame) -> pd.MultiIndex:
    if SITE_COL not in order_df.columns or PRODUCT_UID_COL not in order_df.columns:
        raise KeyError(f"订单统计缺少列 {SITE_COL!r} 或 {PRODUCT_UID_COL!r}，请先运行前置步骤")
    return pd.MultiIndex.from_arrays(
        [
            order_df[SITE_COL].map(_norm_str),
            order_df[PRODUCT_UID_COL].map(_norm_str),
        ],
        names=[SITE_COL, PRODUCT_UID_COL],
    )


def _update_order_stats(order_df: pd.DataFrame, rent_df: pd.DataFrame) -> tuple[pd.DataFrame, int, int]:
    """
    将 MANO 仓租汇总写入订单统计「FBA仓租费」（匹配行直接覆盖，不与原值相加）。
    按「站点」+「商品ID」匹配；平台为空时按站点补全。
    返回 (result_df, 更新行数, 新增行数)
    """
    if RENT_COL not in order_df.columns:
        order_df[RENT_COL] = 0.0
    else:
        order_df[RENT_COL] = pd.to_numeric(order_df[RENT_COL], errors="coerce").fillna(0)

    rent_lookup = rent_df.set_index([SITE_COL, PRODUCT_UID_COL])[RENT_COL]
    order_keys = _order_match_keys(order_df)
    match_mask = order_keys.isin(rent_lookup.index)
    update_cnt = int(match_mask.sum())

    if update_cnt:
        rent_write = pd.Series(
            [rent_lookup.get(key, 0.0) for key in order_keys],
            index=order_df.index,
        )
        order_df.loc[match_mask, RENT_COL] = np.round(rent_write[match_mask].values, 2)
        order_df = _fill_platform_from_site(order_df, match_mask)

    existing_keys = set(zip(order_keys.get_level_values(0), order_keys.get_level_values(1)))
    rent_keys = zip(
        rent_df[SITE_COL].map(_norm_str),
        rent_df[PRODUCT_UID_COL].map(_norm_str),
    )
    missing_rent = rent_df[
        [key not in existing_keys for key in rent_keys]
    ].copy()
    append_cnt = len(missing_rent)
    if append_cnt:
        order_df = pd.concat(
            [order_df, _build_appended_rows(missing_rent, order_df)],
            ignore_index=True,
        )

    return order_df, update_cnt, append_cnt


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
        print(
            f"{Color.YELLOW}[跳过] C2_ManoRent合并 仅月报执行，当前 folder_name={folder_name!r}{Color.RESET}"
        )
        return

    rent_path = Path(MANO_DIR) / RENT_FILE_NAME
    if not rent_path.is_file():
        raise FileNotFoundError(f"未找到 MANO 仓租文件：{rent_path}（请先运行 C2_ManoRent.py）")

    order_path = Path(ORDER_STATS_PATH)
    if not order_path.is_file():
        raise FileNotFoundError(f"未找到订单统计文件：{order_path}（请先运行 K5）")

    detail_df = _strip_df_strings(pd.read_excel(rent_path))
    summary_df = _group_rent_by_site_product(detail_df)
    print(
        f"{Color.CYAN}[汇总]{Color.RESET} {len(detail_df)} 行明细 → "
        f"{len(summary_df)} 个站点+商品ID，仓租合计 {summary_df[RENT_COL].sum():.2f}"
    )

    _append_summary_sheet(rent_path, summary_df)
    print(f"{Color.GREEN}[保存]{Color.RESET} 汇总表已写入：{rent_path} → 工作表「{SUMMARY_SHEET}」")

    original_order_path = _ensure_original_order_stats_backup(order_path)
    order_df = _strip_df_strings(pd.read_excel(original_order_path))
    result_df, update_cnt, append_cnt = _update_order_stats(order_df, summary_df)

    saved = _save_excel(result_df, str(order_path))
    print(
        f"{Color.GREEN}[更新]{Color.RESET} 已写回订单统计："
        f"更新 {update_cnt} 行，新增 {append_cnt} 行，合计 {len(result_df)} 行"
    )
    if append_cnt:
        print(
            f"{Color.YELLOW}[检查] {append_cnt} 条仓租汇总在订单统计中无匹配行，已追加为新行{Color.RESET}"
        )
        missing_keys = {
            (row[SITE_COL], row[PRODUCT_UID_COL])
            for _, row in summary_df.iterrows()
        } - set(zip(
            order_df[SITE_COL].map(_norm_str),
            order_df[PRODUCT_UID_COL].map(_norm_str),
        ))
        preview = summary_df[
            summary_df.apply(
                lambda row: (row[SITE_COL], row[PRODUCT_UID_COL]) in missing_keys,
                axis=1,
            )
        ].head(10)
        if not preview.empty:
            print(preview[[SITE_COL, PRODUCT_UID_COL, RENT_COL]].to_string(index=False))

    print(f"处理完成，output_path：{saved}")
    print(f"{Color.GREEN}一切正常，请检查 FBA仓租费 是否已正确回填{Color.RESET}")


if __name__ == "__main__":
    main()
