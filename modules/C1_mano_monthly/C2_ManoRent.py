"""
C2_ManoRent.py — 合并 MANO MMF 仓租明细并回填站点、SKU、商品ID（仅月报）

功能：
  1. 读取桌面月报目录下 仓租\\mano 中文件名含 @ 的 xlsx（各站点仓租导出）
  2. 从文件名 @ 前提取「站点」（如 MANO-FR-OHPAMF@2025-06.xlsx → MANO-FR-OHPAMF）
  3. 合并为 ALL-WarehouseRent.xlsx（仅保留 GROSS_AMOUNT_VAT_EXC > 0 的行）
  4. 新增列：站点、SKU、商品ID、商品ID识别码（SELLER_SKU 保留原值）
  5. 标准化 seller_sku（去 EXM 前缀、尾缀等）后，若在 product_sku 命中则直接回填 SKU、商品ID
  6. 未命中则：product_sku_mapping（manomano / platform）兜底 → product_sku.product_uid
  7. 商品ID识别码 = 站点 + 商品ID
"""

import importlib.util
import sys
import warnings
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
from config.A0_set_date import folder_name, shared_date
from config.A0_paths import DESKTOP_ROOT

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

from database.db_connection import get_db_manager  # noqa: E402

PSM_TABLE = "product_sku_mapping"
PARTNER_CODE_MANO = "manomano"
PRODUCT_SKU_TABLE = "product_sku"
_KEY_CHUNK = 200

SITE_COL = "站点"
SKU_COL = "SKU"
PRODUCT_UID_COL = "商品ID"
PRODUCT_SITE_ID_COL = "商品ID识别码"
SELLER_SKU_CANDIDATES = ("SELLER_SKU", "sellerSku", "seller_sku")
GROSS_AMOUNT_COL = "GROSS_AMOUNT_VAT_EXC"

MANO_DIR = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\mano"
FILE_GLOB = "*@*.xlsx"
OUTPUT_NAME = "ALL-WarehouseRent.xlsx"


def _strip_df_strings(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
    return df


def _normalize_seller_sku(val) -> str:
    """
    标准化 seller_sku：
    - 去除首尾空格
    - 去掉开头的 EXM
    - 去掉尾缀 -1/-2/-3/-4/-5、_NEVER_USED、-AT
    """
    if pd.isna(val):
        return ""
    sku = str(val).strip()
    
    # 去掉开头的 EXM
    if sku.upper().startswith("EXM"):
        sku = sku[3:]
    
    # 去掉尾缀（按顺序检查，去掉第一个匹配的）
    suffixes = ["_NEVER_USED", "-AT", "-1", "-2", "-3", "-4", "-5"]
    for suffix in suffixes:
        if sku.upper().endswith(suffix.upper()):
            sku = sku[: -len(suffix)]
            break  # 只去掉一次

    return sku.strip()


def _chunked_in_query(cur, sql_template: str, keys: list[str], extra_params: tuple = ()) -> list[dict]:
    if not keys:
        return []
    results: list[dict] = []
    for i in range(0, len(keys), _KEY_CHUNK):
        chunk = keys[i : i + _KEY_CHUNK]
        placeholders = ",".join(["%s"] * len(chunk))
        cur.execute(sql_template.format(placeholders=placeholders), extra_params + tuple(chunk))
        results.extend(cur.fetchall())
    return results


def _site_from_filename(filename: str) -> str:
    """文件名 @ 前为站点，如 MANO-FR-OHPAMF@2025-06.xlsx。"""
    return filename.split("@", 1)[0].strip()


def _find_seller_sku_col(df: pd.DataFrame) -> str:
    for col in SELLER_SKU_CANDIDATES:
        if col in df.columns:
            return col
    raise KeyError(
        f"未找到 SELLER_SKU 列（尝试过 {SELLER_SKU_CANDIDATES}），当前列名: {df.columns.tolist()}"
    )


def _find_gross_amount_col(df: pd.DataFrame) -> str:
    if GROSS_AMOUNT_COL in df.columns:
        return GROSS_AMOUNT_COL
    for col in df.columns:
        if str(col).strip().upper().startswith(GROSS_AMOUNT_COL):
            return col
    raise KeyError(
        f"未找到 {GROSS_AMOUNT_COL!r} 列，当前列名: {df.columns.tolist()}"
    )


def _filter_positive_gross_amount(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """仅保留 GROSS_AMOUNT_VAT_EXC > 0 的行，返回 (筛选后 df, 剔除行数)。"""
    gross_col = _find_gross_amount_col(df)
    before = len(df)
    amount = pd.to_numeric(df[gross_col], errors="coerce").fillna(0)
    filtered = df[amount > 0].copy()
    return filtered, before - len(filtered)


def fetch_product_sku_from_mapping(seller_skus: list[str]) -> dict[str, str]:
    """
    从 product_sku_mapping 按 seller_sku 查 product_sku（MANO 平台维度）。
    返回 seller_sku → product_sku，同一 seller_sku 取 updated_at 最新一条。
    """
    seller_skus = sorted({str(x).strip() for x in seller_skus if x and str(x).strip()})
    if not seller_skus:
        return {}

    sql = f"""
        SELECT seller_sku, product_sku, updated_at
        FROM `{PSM_TABLE}`
        WHERE partner_code = %s
          AND partner_type = 'platform'
          AND mapping_type = 'single'
          AND is_active = 1
          AND seller_sku IN ({{placeholders}})
          AND product_sku IS NOT NULL
          AND TRIM(product_sku) <> ''
        ORDER BY updated_at DESC
    """
    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        rows: list[dict] = []
        for i in range(0, len(seller_skus), _KEY_CHUNK):
            chunk = seller_skus[i : i + _KEY_CHUNK]
            placeholders = ",".join(["%s"] * len(chunk))
            cur.execute(
                sql.format(placeholders=placeholders),
                (PARTNER_CODE_MANO,) + tuple(chunk),
            )
            rows.extend(cur.fetchall())
    finally:
        cur.close()
        conn.close()

    mapping: dict[str, str] = {}
    for row in rows:
        seller_sku = str(row.get("seller_sku") or "").strip()
        product_sku = str(row.get("product_sku") or "").strip()
        if seller_sku and product_sku and seller_sku not in mapping:
            mapping[seller_sku] = product_sku
    return mapping


def fetch_product_sku_direct(sku_keys: list[str]) -> dict[str, tuple[str, str]]:
    """
    按 product_sku 精确查询，返回 key → (product_sku, product_uid)。
    key 与入参 product_sku 一致（用于标准化后的 seller_sku 直接命中）。
    """
    sku_keys = sorted({str(x).strip() for x in sku_keys if x and str(x).strip()})
    if not sku_keys:
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
        rows = _chunked_in_query(cur, sql, sku_keys)
    finally:
        cur.close()
        conn.close()

    mapping: dict[str, tuple[str, str]] = {}
    for row in rows:
        sku = str(row.get("product_sku") or "").strip()
        uid = str(row.get("product_uid") or "").strip()
        if sku and uid:
            mapping[sku] = (sku, uid)
    return mapping


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


def _lookup_product_sku_from_mapping(
    norm_sku: str,
    raw_sku: str,
    mapping: dict[str, str],
) -> str | None:
    """优先标准化 seller_sku，再尝试原值。"""
    for key in (norm_sku, raw_sku):
        if key and key in mapping:
            return mapping[key]
    return None


def _list_input_files(mano_dir: Path) -> list[Path]:
    files = sorted(mano_dir.glob(FILE_GLOB))
    skip_prefixes = ("(已完成", "(处理完成)")
    return [
        p
        for p in files
        if p.is_file()
        and p.name != OUTPUT_NAME
        and not any(p.name.startswith(prefix) for prefix in skip_prefixes)
    ]


def _reorder_columns(df: pd.DataFrame, seller_col: str) -> pd.DataFrame:
    """站点放首列；SKU、商品ID、商品ID识别码紧跟 SELLER_SKU 之后。"""
    extra_cols = [SKU_COL, PRODUCT_UID_COL, PRODUCT_SITE_ID_COL]
    for col in extra_cols:
        if col not in df.columns:
            raise KeyError(f"缺少列 {col!r}")

    front = [SITE_COL] if SITE_COL in df.columns else []
    seller_idx = df.columns.get_loc(seller_col)
    before_seller = [c for c in df.columns if c not in front + extra_cols and df.columns.get_loc(c) < seller_idx]
    after_seller = [
        c
        for c in df.columns
        if c not in front + extra_cols + [seller_col] and df.columns.get_loc(c) > seller_idx
    ]
    ordered = front + before_seller + [seller_col] + extra_cols + after_seller
    seen: set[str] = set()
    final_cols: list[str] = []
    for c in ordered:
        if c in df.columns and c not in seen:
            final_cols.append(c)
            seen.add(c)
    for c in df.columns:
        if c not in seen:
            final_cols.append(c)
    return df[final_cols]


def _fill_sku_and_product_uid(df: pd.DataFrame) -> pd.DataFrame:
    """
    回填 SKU、商品ID：
      1. 标准化 seller_sku 后查 product_sku → 直接回填
      2. 未命中 → product_sku_mapping（manomano）兜底 → product_uid
    SELLER_SKU 列保留原值。
    """
    seller_col = _find_seller_sku_col(df)
    for col in (SKU_COL, PRODUCT_UID_COL):
        if col in df.columns:
            df = df.drop(columns=[col])

    raw_skus = df[seller_col].tolist()
    norm_skus = df[seller_col].map(_normalize_seller_sku).tolist()
    row_count = len(df)

    wh_skus: list[str | None] = [None] * row_count
    uids: list[str | None] = [None] * row_count

    # ① 标准化 seller_sku 直接命中 product_sku
    unique_norm = sorted({s for s in norm_skus if s})
    direct_map = fetch_product_sku_direct(unique_norm)
    if direct_map:
        print(f"{Color.GREEN}[DB] product_sku 直接命中 {len(direct_map)} 个标准化 seller_sku{Color.RESET}")

    direct_hit = 0
    for i, norm in enumerate(norm_skus):
        if norm and norm in direct_map:
            sku, uid = direct_map[norm]
            wh_skus[i] = sku
            uids[i] = uid
            direct_hit += 1
    if direct_hit:
        print(f"{Color.GREEN}[DB] product_sku 直接回填 {direct_hit} 行 SKU + 商品ID{Color.RESET}")

    # ② 未命中行：product_sku_mapping 兜底
    pending_indices = [i for i in range(row_count) if wh_skus[i] is None and norm_skus[i]]
    if pending_indices:
        mapping_keys = sorted(
            {
                key
                for i in pending_indices
                for key in (norm_skus[i], str(raw_skus[i]).strip() if pd.notna(raw_skus[i]) else "")
                if key
            }
        )
        psm_map = fetch_product_sku_from_mapping(mapping_keys)
        if psm_map:
            print(f"{Color.CYAN}[DB] product_sku_mapping 查到 {len(psm_map)} 条 seller_sku 映射{Color.RESET}")

        psm_hit = 0
        for i in pending_indices:
            raw = str(raw_skus[i]).strip() if pd.notna(raw_skus[i]) else ""
            product_sku = _lookup_product_sku_from_mapping(norm_skus[i], raw, psm_map)
            if product_sku:
                wh_skus[i] = product_sku
                psm_hit += 1
        if psm_hit:
            print(f"{Color.CYAN}[DB] product_sku_mapping 补全 {psm_hit} 行 SKU{Color.RESET}")

    # ③ 补商品ID（SKU 已有、商品ID 仍空的行）
    need_uid_skus = sorted({wh_skus[i] for i in range(row_count) if wh_skus[i] and not uids[i]})
    if need_uid_skus:
        uid_map = fetch_product_uid_by_warehouse_sku(need_uid_skus)
        if uid_map:
            print(
                f"{Color.GREEN}[DB] product_sku 查到 {len(uid_map)} 条 "
                f"product_sku → product_uid 映射{Color.RESET}"
            )
        for i in range(row_count):
            if wh_skus[i] and not uids[i]:
                uids[i] = uid_map.get(str(wh_skus[i]).strip())

    insert_pos = df.columns.get_loc(seller_col) + 1
    df.insert(insert_pos, SKU_COL, wh_skus)
    df.insert(insert_pos + 1, PRODUCT_UID_COL, uids)
    return df


def _fill_product_site_id(df: pd.DataFrame) -> pd.DataFrame:
    if PRODUCT_SITE_ID_COL in df.columns:
        df = df.drop(columns=[PRODUCT_SITE_ID_COL])

    site = df[SITE_COL].map(lambda x: str(x).strip() if pd.notna(x) else "")
    uid = df[PRODUCT_UID_COL].map(lambda x: str(x).strip() if pd.notna(x) else "")
    df[PRODUCT_SITE_ID_COL] = site + uid
    invalid = site.eq("") | uid.eq("") | uid.eq("nan")
    df.loc[invalid, PRODUCT_SITE_ID_COL] = pd.NA
    return df


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
        print(f"{Color.YELLOW}[跳过] C2_ManoRent 仅月报执行，当前 folder_name={folder_name!r}{Color.RESET}")
        return

    mano_dir = Path(MANO_DIR)
    if not mano_dir.is_dir():
        raise FileNotFoundError(f"未找到 MANO 仓租目录：{mano_dir}")

    input_files = _list_input_files(mano_dir)
    if not input_files:
        raise FileNotFoundError(f"目录下未找到待合并文件：{mano_dir}\\{FILE_GLOB}")

    parts: list[pd.DataFrame] = []
    total_dropped = 0
    for file_path in input_files:
        site = _site_from_filename(file_path.stem)
        df = _strip_df_strings(pd.read_excel(file_path))
        df.insert(0, SITE_COL, site)
        before_rows = len(df)
        df, dropped = _filter_positive_gross_amount(df)
        total_dropped += dropped
        parts.append(df)
        drop_msg = f"，剔除 {dropped} 行（GROSS_AMOUNT_VAT_EXC ≤ 0）" if dropped else ""
        print(
            f"{Color.CYAN}[读取]{Color.RESET} {file_path.name}（站点={site}，"
            f"{before_rows} 行 → {len(df)} 行{drop_msg}）"
        )

    merged = pd.concat(parts, ignore_index=True)
    seller_col = _find_seller_sku_col(merged)
    merged = _fill_sku_and_product_uid(merged)
    merged = _fill_product_site_id(merged)
    merged = _reorder_columns(merged, seller_col)

    output_path = mano_dir / OUTPUT_NAME
    saved = _save_excel(merged, str(output_path))

    missing_wh = merged[seller_col].notna() & merged[SKU_COL].isna()
    missing_uid = merged[SKU_COL].notna() & merged[PRODUCT_UID_COL].isna()
    if missing_wh.any():
        print(
            f"{Color.YELLOW}[检查] {missing_wh.sum()} 行未映射到 SKU"
            f"（标准化 seller_sku 在 product_sku / product_sku_mapping 均无匹配）{Color.RESET}"
        )
        preview = merged.loc[missing_wh, [SITE_COL, seller_col]].drop_duplicates().head(10)
        print(preview.to_string(index=False))
    if missing_uid.any():
        print(
            f"{Color.YELLOW}[检查] {missing_uid.sum()} 行未映射到商品ID"
            f"（SKU 在 product_sku 中无 product_uid）{Color.RESET}"
        )
        preview = merged.loc[missing_uid, [SITE_COL, seller_col, SKU_COL]].drop_duplicates().head(10)
        print(preview.to_string(index=False))

    print(
        f"{Color.GREEN}[合并]{Color.RESET} 共 {len(input_files)} 个文件 → {len(merged)} 行"
        f"（剔除 GROSS_AMOUNT_VAT_EXC ≤ 0 共 {total_dropped} 行），已保存：{saved}"
    )
    print(f"{Color.GREEN}一切正常，请检查未映射行并手动补充（如有）{Color.RESET}")


if __name__ == "__main__":
    main()
