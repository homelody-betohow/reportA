"""
C1_Mano_vat.py — MANO VAT 和佣金报表：回填仓库 SKU、商品 ID（仅月报）

功能：
  1. 读取桌面月报目录下 仓租\\mano 中的 MANO-VAT和佣金-*.xlsx
  2. sellerSku 含「+」的组合产品拆成多行，费用列均摊
  3. 在 sellerSku 列后插入「仓库SKU」「商品ID」「SKU-站点识别码」
  4. 通过 sales_order_shipped.warehouse_sku → warehouse_sku 回填仓库 SKU
  5. 通过 product_sku.product_sku → product_uid 回填商品 ID
  6. 生成「SKU-站点识别码」= 站点 + 仓库SKU

输出：同目录下 (已完成-1)原文件名.xlsx
"""

import importlib.util
import sys
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd
import pymysql.cursors

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
from common.split_rows_data_SKU import split_one_rows_data
from config.A0_set_date import folder_name, shared_date
from config.A0_paths import DESKTOP_ROOT

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

_REPORT_PRA_ROOT = next(
    (p / "reportPRA" for p in Path(__file__).resolve().parents if (p / "reportPRA").is_dir()),
    None,
)
if _REPORT_PRA_ROOT and str(_REPORT_PRA_ROOT) not in sys.path:
    sys.path.append(str(_REPORT_PRA_ROOT))

from database.db_connection import get_db_manager  # noqa: E402  # pyright: ignore[reportMissingImports]

SHIPPED_TABLE = "sales_order_shipped"
PRODUCT_SKU_TABLE = "product_sku"
_KEY_CHUNK = 200
SELLER_SKU_COL = "sellerSku"
SITE_COL = "站点"
WH_SKU_COL = "仓库SKU"
PRODUCT_UID_COL = "商品ID"
SON_SITE_ID_COL = "SKU-站点识别码"

MANO_DIR = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\mano-vat"
FILE_GLOB = "MANO-VAT和佣金-*.xlsx"

# MANO VAT 报表中需要随组合 SKU 拆行均摊的金额列（quantity 等非费用列不拆）
_FEE_COLUMNS = (
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


def _chunked_in_query(cur, sql_template: str, keys: list[str], extra_params: tuple = ()) -> list[dict]:
    """按批次执行 IN 查询，返回合并后的 dict 行列表。"""
    if not keys:
        return []
    results: list[dict] = []
    for i in range(0, len(keys), _KEY_CHUNK):
        chunk = keys[i : i + _KEY_CHUNK]
        placeholders = ",".join(["%s"] * len(chunk))
        cur.execute(sql_template.format(placeholders=placeholders), extra_params + tuple(chunk))
        results.extend(cur.fetchall())
    return results


def _ts(row: dict) -> float:
    """ship_time 转时间戳；无法解析则视为 0（更旧）。"""
    st = row.get("ship_time")
    if isinstance(st, datetime):
        return st.timestamp()
    if st is not None and hasattr(st, "timestamp"):
        try:
            return float(st.timestamp())
        except Exception:
            pass
    if isinstance(st, str):
        s = st.strip()
        if not s:
            return 0.0
        try:
            return datetime.fromisoformat(s.replace(" ", "T")).timestamp()
        except Exception:
            return 0.0
    return 0.0


def _rid(row: dict) -> int:
    try:
        return int(row.get("id") or 0)
    except (TypeError, ValueError):
        return 0


def _is_newer(a: dict, b: dict) -> bool:
    ta, tb = _ts(a), _ts(b)
    if ta != tb:
        return ta > tb
    return _rid(a) > _rid(b)


def _normalize_seller_sku(val) -> str:
    if pd.isna(val):
        return ""
    return str(val).strip()


def fetch_warehouse_sku_by_platform_sku(platform_skus: list[str]) -> dict[str, str]:
    """
    从 sales_order_shipped 按 platform_sku 查 warehouse_sku。
    同一 platform_sku 多条时，保留 ship_time / id 最新的一条。
    """
    platform_skus = sorted({_normalize_seller_sku(x) for x in platform_skus if _normalize_seller_sku(x)})
    if not platform_skus:
        return {}

    sql = f"""
        SELECT id, ship_time, platform_sku, warehouse_sku
        FROM `{SHIPPED_TABLE}`
        WHERE platform_sku IN ({{placeholders}})
          AND warehouse_sku IS NOT NULL
          AND TRIM(warehouse_sku) <> ''
        ORDER BY id ASC
    """
    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        rows = _chunked_in_query(cur, sql, platform_skus)
    finally:
        cur.close()
        conn.close()

    chosen: dict[str, dict] = {}
    dup_keys: set[str] = set()
    for row in rows:
        platform_sku = _normalize_seller_sku(row.get("platform_sku"))
        warehouse_sku = str(row.get("warehouse_sku") or "").strip()
        if not platform_sku or not warehouse_sku:
            continue
        prev = chosen.get(platform_sku)
        if prev is None:
            chosen[platform_sku] = row
            continue
        prev_wh = str(prev.get("warehouse_sku") or "").strip()
        if prev_wh != warehouse_sku:
            dup_keys.add(platform_sku)
            if _is_newer(row, prev):
                chosen[platform_sku] = row

    if dup_keys:
        print(
            f"{Color.CYAN}[DB][警告] sales_order_shipped 中同一 platform_sku 对应多个 warehouse_sku，"
            f"已按 ship_time/id 选择较新记录（共 {len(dup_keys)} 个 SKU）{Color.RESET}"
        )

    return {
        k: str(v.get("warehouse_sku") or "").strip()
        for k, v in chosen.items()
        if str(v.get("warehouse_sku") or "").strip()
    }


def fetch_product_uid_by_warehouse_sku(warehouse_skus: list[str]) -> dict[str, str]:
    """从 product_sku 按 product_sku 查 product_uid。"""
    warehouse_skus = sorted({str(x).strip() for x in warehouse_skus if x and str(x).strip()})
    if not warehouse_skus:
        return {}

    sql = f"""
        SELECT product_sku, product_uid
        FROM `{PRODUCT_SKU_TABLE}`
        WHERE product_sku IN ({{placeholders}})
          AND product_uid IS NOT NULL
          AND TRIM(product_uid) <> ''
    """
    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        rows = _chunked_in_query(cur, sql, warehouse_skus)
    finally:
        cur.close()
        conn.close()

    mapping: dict[str, str] = {}
    for row in rows:
        sku = str(row.get("product_sku") or "").strip()
        uid = str(row.get("product_uid") or "").strip()
        if sku and uid:
            mapping[sku] = uid
    return mapping


def _strip_df_strings(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
    return df


def _fee_columns_in_df(df: pd.DataFrame) -> list[str]:
    return [c for c in _FEE_COLUMNS if c in df.columns]


def _split_combo_seller_sku(df: pd.DataFrame) -> pd.DataFrame:
    """sellerSku 含 '+' 的组合产品拆成多行，费用列按子 SKU 数量均摊。"""
    combo_mask = df[SELLER_SKU_COL].astype(str).str.contains(r"[+,]", na=False, regex=True)
    combo_cnt = int(combo_mask.sum())
    if combo_cnt == 0:
        return df

    fee_cols = _fee_columns_in_df(df)
    if not fee_cols:
        raise ValueError(f"未找到可均摊的费用列，当前列名: {df.columns.tolist()}")

    df = split_one_rows_data(
        input_df=df,
        data_column=SELLER_SKU_COL,
        value_column=fee_cols,
    )
    print(
        f"{Color.YELLOW}[组合SKU]{Color.RESET} 拆分 {combo_cnt} 行（sellerSku 含 +/，），"
        f"费用列均摊：{', '.join(fee_cols)}"
    )
    return df


def _ensure_cols_after_seller_sku(df: pd.DataFrame) -> pd.DataFrame:
    """在 sellerSku 后确保存在「仓库SKU」「商品ID」列（若已有则先移除再按位置插入）。"""
    if SELLER_SKU_COL not in df.columns:
        raise KeyError(f"未找到列 {SELLER_SKU_COL!r}，当前列名: {df.columns.tolist()}")

    for col in (WH_SKU_COL, PRODUCT_UID_COL, SON_SITE_ID_COL):
        if col in df.columns:
            df = df.drop(columns=[col])

    insert_pos = df.columns.get_loc(SELLER_SKU_COL) + 1
    df.insert(insert_pos, WH_SKU_COL, pd.NA)
    df.insert(insert_pos + 1, PRODUCT_UID_COL, pd.NA)
    df.insert(insert_pos + 2, SON_SITE_ID_COL, pd.NA)
    return df


def _fill_sku_and_product_uid(df: pd.DataFrame) -> pd.DataFrame:
    seller_skus = df[SELLER_SKU_COL].map(_normalize_seller_sku).tolist()
    wh_map = fetch_warehouse_sku_by_platform_sku(seller_skus)
    if wh_map:
        print(
            f"{Color.GREEN}[DB] 从 sales_order_shipped 查到 {len(wh_map)} 条 "
            f"platform_sku → warehouse_sku 映射{Color.RESET}"
        )

    df[WH_SKU_COL] = df[SELLER_SKU_COL].map(
        lambda x: wh_map.get(_normalize_seller_sku(x), pd.NA)
    )

    mapped_wh = df[WH_SKU_COL].dropna().astype(str).str.strip().tolist()
    uid_map = fetch_product_uid_by_warehouse_sku(mapped_wh)
    if uid_map:
        print(
            f"{Color.GREEN}[DB] 从 product_sku 查到 {len(uid_map)} 条 "
            f"product_sku → product_uid 映射{Color.RESET}"
        )

    df[PRODUCT_UID_COL] = df[WH_SKU_COL].map(
        lambda x: uid_map.get(str(x).strip(), pd.NA) if pd.notna(x) and str(x).strip() else pd.NA
    )
    return df


def _fill_son_site_id(df: pd.DataFrame) -> pd.DataFrame:
    """SKU-站点识别码 = 站点 + 仓库SKU（与 D3_MANO、订单统计等脚本一致）。"""
    if SITE_COL not in df.columns:
        raise KeyError(f"未找到列 {SITE_COL!r}，当前列名: {df.columns.tolist()}")

    site = df[SITE_COL].map(lambda x: str(x).strip() if pd.notna(x) else "")
    wh = df[WH_SKU_COL].map(lambda x: str(x).strip() if pd.notna(x) else "")
    df[SON_SITE_ID_COL] = site + wh
    df.loc[site.eq("") | wh.eq(""), SON_SITE_ID_COL] = pd.NA
    return df


def _save_excel(df: pd.DataFrame, output_path: str) -> str:
    try:
        df.to_excel(output_path, index=False)
        return output_path
    except PermissionError:
        alt = output_path.replace(".xlsx", "-另存.xlsx")
        df.to_excel(alt, index=False)
        return alt


def _list_input_files(mano_dir: Path) -> list[Path]:
    files = sorted(mano_dir.glob(FILE_GLOB))
    return [p for p in files if p.is_file() and not p.name.startswith("(已完成")]


def process_one_file(file_path: Path) -> str:
    df = pd.read_excel(file_path)
    df = _strip_df_strings(df)
    df = _split_combo_seller_sku(df)
    df = _ensure_cols_after_seller_sku(df)
    df = _fill_sku_and_product_uid(df)
    df = _fill_son_site_id(df)

    output_path = file_path.parent / f"(已完成-1){file_path.name}"
    saved = _save_excel(df, str(output_path))

    missing_wh = df[SELLER_SKU_COL].notna() & df[WH_SKU_COL].isna()
    missing_uid = df[WH_SKU_COL].notna() & df[PRODUCT_UID_COL].isna()
    missing_son_id = df[WH_SKU_COL].notna() & df[SON_SITE_ID_COL].isna()
    if missing_wh.any():
        print(
            f"{Color.YELLOW}[检查] {file_path.name}：{missing_wh.sum()} 行未映射到仓库SKU"
            f"（sellerSku 在 sales_order_shipped 中无匹配）{Color.RESET}"
        )
        preview = df.loc[missing_wh, [SELLER_SKU_COL]].drop_duplicates().head(10)
        print(preview.to_string(index=False))
    if missing_uid.any():
        print(
            f"{Color.YELLOW}[检查] {file_path.name}：{missing_uid.sum()} 行未映射到商品ID"
            f"（仓库SKU 在 product_sku 中无 product_uid）{Color.RESET}"
        )
        preview = df.loc[missing_uid, [SELLER_SKU_COL, WH_SKU_COL]].drop_duplicates().head(10)
        print(preview.to_string(index=False))
    if missing_son_id.any():
        print(
            f"{Color.YELLOW}[检查] {file_path.name}：{missing_son_id.sum()} 行未生成SKU-站点识别码"
            f"（站点或仓库SKU为空）{Color.RESET}"
        )
        preview = df.loc[missing_son_id, [SITE_COL, SELLER_SKU_COL, WH_SKU_COL]].drop_duplicates().head(10)
        print(preview.to_string(index=False))

    return saved


def main() -> None:
    if folder_name != "月报":
        print(f"{Color.YELLOW}[跳过] C1_Mano_vat 仅月报执行，当前 folder_name={folder_name!r}{Color.RESET}")
        return

    mano_dir = Path(MANO_DIR)
    if not mano_dir.is_dir():
        raise FileNotFoundError(f"未找到 MANO 目录：{mano_dir}")

    input_files = _list_input_files(mano_dir)
    if not input_files:
        raise FileNotFoundError(
            f"目录下未找到待处理文件：{mano_dir}\\{FILE_GLOB}（已排除 (已完成*) 文件）"
        )

    print(f"{Color.CYAN}[C1_Mano_vat] 月报 {shared_date}，共 {len(input_files)} 个文件{Color.RESET}")
    for file_path in input_files:
        saved = process_one_file(file_path)
        print(f"处理完成，文件另存为：{saved}")

    print(f"{Color.GREEN}一切正常，请检查未映射行并手动补充（如有）{Color.RESET}")


if __name__ == "__main__":
    main()
