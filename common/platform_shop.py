"""
platform_shop 统一映射：

1) 店铺 → 站点 / 平台（替代原桌面「站点-匹配表.xlsx」）
2) 站点 → 平台（market_region → market_code）
3) 站点 → VAT / 平台佣金（DB 优先，桌面「VAT、平台费-映射.xlsx」兜底）

字段对应（Excel → DB）：
  站点 → market_region
  平台 → market_code
  区域 → platform_site
  币种 → currency
  平台费（佣金） → commission_rate
  VAT税 → vat_rate
"""
from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
import pymysql.cursors

from database.db_connection import get_db_manager

# LM_BC_FR / LM_RP_FR：映射站点按平台 sku 前缀追加 -ls / -xj（DB 存的是无后缀站点）
LM_FR_SHOPS: frozenset[str] = frozenset({"LM_BC_FR", "LM_RP_FR"})
_LM_REGION_SUFFIX_RE = re.compile(r"-(?:ls|xj)$")

_EXCEL_SHEET = "VAT税、佣金"
_EXCEL_FILE = "VAT、平台费-映射.xlsx"

_PLATFORM_SHOP_FEE_SQL = """
    SELECT
        TRIM(market_region) AS market_region,
        TRIM(market_code) AS market_code,
        TRIM(platform) AS platform,
        TRIM(platform_site) AS platform_site,
        TRIM(currency) AS currency,
        commission_rate,
        vat_rate
    FROM platform_shop
    WHERE shop_status = 1
      AND TRIM(market_region) <> ''
"""


# platform_shop 表字段对应关系：
#   shop_name_en  → 平台账号（店铺英文名）
#   market_region → 站点
#   market_code   → 平台
#   platform_site → 用于“店铺英文名-站点”级别的精确区分（对应原 Excel 的“特殊-平台账号”）
_PLATFORM_SHOP_SQL = """
    SELECT
        TRIM(shop_name_en)  AS shop_name_en,
        TRIM(platform_site) AS platform_site,
        TRIM(market_region) AS market_region,
        TRIM(market_code)   AS market_code
    FROM platform_shop
    WHERE TRIM(shop_name_en) <> ''
    ORDER BY id ASC
"""


def lm_suffix_from_platform_sku(platform_sku: str) -> str:
    """平台 sku 以 ls- 开头 → -ls，否则 → -xj。"""
    s = str(platform_sku or "").strip()
    return "-ls" if s.startswith("ls-") else "-xj"


def strip_lm_region_suffix(region: str) -> str:
    """查询 platform_shop 前去掉 LM 映射站点后缀 -ls / -xj。"""
    s = str(region or "").strip()
    return _LM_REGION_SUFFIX_RE.sub("", s) if s else s


def apply_lm_fr_region_suffix(
    df: pd.DataFrame,
    *,
    shop_col: str = "店铺英文名",
    platform_sku_col: str = "平台sku",
    region_col: str = "映射站点",
    shops: frozenset[str] | tuple[str, ...] = LM_FR_SHOPS,
    inplace: bool = False,
) -> pd.DataFrame:
    """
    LM_BC_FR / LM_RP_FR：按平台 sku 前缀给映射站点加 -ls / -xj 后缀。

    需在 map_shop_platform_region 之后调用；后缀仅用于业务识别码，查库时应 strip_lm_region_suffix。
    """
    result = df if inplace else df.copy()
    for shop in shops:
        if shop_col not in result.columns:
            continue
        mask = result[shop_col] == shop
        if not mask.any():
            continue
        suffix = result.loc[mask, platform_sku_col].astype(str).map(lm_suffix_from_platform_sku)
        result.loc[mask, region_col] = result.loc[mask, region_col].astype(str) + suffix
    return result


def _normalize_site_keys_for_lookup(site_keys: pd.Series) -> pd.Series:
    """站点列查 platform_shop 时去掉 LM 的 -ls / -xj 后缀。"""
    return site_keys.map(strip_lm_region_suffix)


def _default_excel_path() -> Path:
    from config.A0_paths import DESKTOP_ROOT

    return Path(DESKTOP_ROOT) / _EXCEL_FILE


def fetch_platform_shop_fee_df() -> pd.DataFrame:
    """读取启用店铺的平台费/VAT 映射（按 market_region 去重）。"""
    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        cur.execute(_PLATFORM_SHOP_FEE_SQL)
        rows = cur.fetchall()
    finally:
        cur.close()
        conn.close()

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    return df.drop_duplicates(subset=["market_region"], keep="first").reset_index(drop=True)


def fetch_excel_fee_df(excel_path: str | Path | None = None) -> pd.DataFrame:
    """读取 Excel 兜底映射表。"""
    path = Path(excel_path) if excel_path else _default_excel_path()
    if not path.is_file():
        return pd.DataFrame(columns=["market_region", "commission_rate", "vat_rate"])

    df = pd.read_excel(path, sheet_name=_EXCEL_SHEET)
    df.columns = df.columns.str.strip()
    df = df[["站点", "平台费（佣金）", "VAT税"]].rename(
        columns={
            "站点": "market_region",
            "平台费（佣金）": "commission_rate",
            "VAT税": "vat_rate",
        }
    )
    df["market_region"] = df["market_region"].astype(str).str.strip()
    df = df[df["market_region"].notna() & (df["market_region"] != "nan")]
    return df.drop_duplicates(subset=["market_region"], keep="last").reset_index(drop=True)


def fetch_platform_shop_rows() -> list[dict]:
    """从数据库 platform_shop 表读取店铺→站点/平台的映射数据。"""
    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        cur.execute(_PLATFORM_SHOP_SQL)
        rows = cur.fetchall()
    finally:
        cur.close()
        conn.close()
    return rows


def build_shop_maps(rows: list[dict]) -> tuple[dict, dict, dict]:
    """
    构建三张映射字典（同一 key 多行时，后出现的覆盖先出现的，与原 Excel 逻辑一致）：
      - platform_by_shop:      店铺英文名                → 平台(market_code)
      - region_by_shop:        店铺英文名                → 站点(market_region)   （兜底）
      - region_by_shop_site:   店铺英文名-站点(platform_site) → 站点(market_region)（精确）
    """
    platform_by_shop: dict[str, str] = {}
    region_by_shop: dict[str, str] = {}
    region_by_shop_site: dict[str, str] = {}
    for r in rows:
        shop = str(r.get("shop_name_en") or "").strip()
        if not shop:
            continue
        site = str(r.get("platform_site") or "").strip()
        region = str(r.get("market_region") or "").strip()
        code = str(r.get("market_code") or "").strip()
        if code:
            platform_by_shop[shop] = code
        if region:
            region_by_shop[shop] = region
            if site:
                region_by_shop_site[f"{shop}-{site}"] = region
    return platform_by_shop, region_by_shop, region_by_shop_site


def map_shop_platform_region(
    main_df: pd.DataFrame,
    shop_col: str,
    site_col: str | None = "站点",
    *,
    shop_site_col: str = "店铺英文名-站点",
    platform_col: str = "映射平台",
    region_col: str = "映射站点",
) -> pd.DataFrame:
    """
    按店铺从 platform_shop 写入「映射平台」「映射站点」。

    映射站点优先级（site_col 有值时）：
      1) 店铺英文名-站点（platform_site）精确匹配
      2) 店铺英文名兜底
    site_col=None 时仅按店铺英文名兜底，不生成「店铺英文名-站点」。
    未命中为空（与原「站点-匹配表」对 站点/平台 列行为一致）。
    """
    if shop_col not in main_df.columns:
        raise KeyError(f"主表缺少店铺列：{shop_col}")
    if site_col is not None and site_col not in main_df.columns:
        raise KeyError(f"主表缺少站点列：{site_col}")

    platform_by_shop, region_by_shop, region_by_shop_site = build_shop_maps(
        fetch_platform_shop_rows()
    )

    result = main_df.copy()
    shop_series = result[shop_col].astype(str).str.strip()
    result[platform_col] = shop_series.map(platform_by_shop)

    if site_col is None:
        result[region_col] = shop_series.map(region_by_shop)
        return result

    result[shop_site_col] = shop_series + "-" + result[site_col].astype(str).str.strip()
    region_special = result[shop_site_col].map(region_by_shop_site)
    region_fallback = shop_series.map(region_by_shop)
    result[region_col] = region_special.where(region_special.notna(), region_fallback)
    return result


def build_region_to_platform_map(rows: list[dict] | None = None) -> dict[str, str]:
    """market_region → market_code（同一 key 后出现覆盖先出现）。"""
    if rows is None:
        rows = fetch_platform_shop_rows()
    mapping: dict[str, str] = {}
    for r in rows:
        region = str(r.get("market_region") or "").strip()
        code = str(r.get("market_code") or "").strip()
        if region and code:
            mapping[region] = code
    return mapping


def map_region_to_platform(
    main_df: pd.DataFrame,
    site_col: str = "站点",
    *,
    platform_col: str = "映射平台",
) -> pd.DataFrame:
    """
    按站点(market_region)写入「映射平台」(market_code)。
    未命中为空；列插入位置紧挨 site_col 之后（与原 sku_mappings 一致）。
    """
    if site_col not in main_df.columns:
        raise KeyError(f"主表缺少站点列：{site_col}")

    mapping = build_region_to_platform_map()
    result = main_df.copy()
    site_keys = _normalize_site_keys_for_lookup(
        result[site_col].astype(str).str.strip()
    )
    mapped = site_keys.map(mapping)
    if platform_col in result.columns:
        result[platform_col] = mapped
    else:
        insert_pos = result.columns.get_loc(site_col) + 1
        result.insert(insert_pos, platform_col, mapped)
    return result


def map_shop_to_region(
    main_df: pd.DataFrame,
    shop_col: str,
    *,
    region_col: str = "映射站点",
) -> pd.DataFrame:
    """
    按店铺英文名写入「映射站点」(market_region)。
    未命中为空；列插入位置紧挨 shop_col 之后。
    """
    if shop_col not in main_df.columns:
        raise KeyError(f"主表缺少店铺列：{shop_col}")

    _, region_by_shop, _ = build_shop_maps(fetch_platform_shop_rows())
    result = main_df.copy()
    mapped = result[shop_col].astype(str).str.strip().map(region_by_shop)
    if region_col in result.columns:
        result[region_col] = mapped
    else:
        insert_pos = result.columns.get_loc(shop_col) + 1
        result.insert(insert_pos, region_col, mapped)
    return result


def _build_fee_lookup(fee_df: pd.DataFrame) -> dict[str, dict]:
    """market_region（大小写不敏感）-> 费率行。"""
    lookup: dict[str, dict] = {}
    for row in fee_df.to_dict("records"):
        key = str(row["market_region"]).strip().casefold()
        if key:
            lookup[key] = row
    return lookup


def _rates_from_lookup(site_keys: pd.Series, lookup: dict[str, dict], field: str) -> pd.Series:
    return site_keys.map(lambda k: lookup.get(k, {}).get(field))


def _merge_db_excel_rates(db_rates: pd.Series, excel_rates: pd.Series) -> pd.Series:
    """数据库优先；仅当 DB 为空时用 Excel 兜底（0 视为有效值）。"""
    merged = db_rates.where(pd.notna(db_rates), excel_rates)
    # pymysql 返回 Decimal，须转为 float 才能与 Excel/pandas 数值列做运算
    return pd.to_numeric(merged, errors="coerce")


def map_site_vat_commission(
    main_df: pd.DataFrame,
    site_col: str = "站点",
    fee_df: pd.DataFrame | None = None,
    excel_path: str | Path | None = None,
    *,
    excel_fallback: bool = True,
    verbose: bool = True,
) -> pd.DataFrame:
    """
    按站点映射「映射平台费（佣金）」与「映射VAT税」列。

    优先级：platform_shop；excel_fallback=True 时再用桌面「VAT、平台费-映射.xlsx」兜底。
    列插入位置与 sku_mappings 一致：紧挨 site_col 之后。
    """
    if site_col not in main_df.columns:
        raise KeyError(f"主表缺少站点列：{site_col}")

    db_df = fetch_platform_shop_fee_df() if fee_df is None else fee_df
    db_lookup = _build_fee_lookup(db_df)

    result = main_df.copy()
    site_keys = _normalize_site_keys_for_lookup(
        result[site_col].astype(str).str.strip()
    ).str.casefold()

    db_comm = _rates_from_lookup(site_keys, db_lookup, "commission_rate")
    db_vat = _rates_from_lookup(site_keys, db_lookup, "vat_rate")

    commission_col = "映射平台费（佣金）"
    vat_col = "映射VAT税"
    insert_pos = result.columns.get_loc(site_col) + 1

    if excel_fallback:
        excel_df = fetch_excel_fee_df(excel_path)
        excel_lookup = _build_fee_lookup(excel_df)
        excel_comm = _rates_from_lookup(site_keys, excel_lookup, "commission_rate")
        excel_vat = _rates_from_lookup(site_keys, excel_lookup, "vat_rate")
        merged_comm = _merge_db_excel_rates(db_comm, excel_comm)
        merged_vat = _merge_db_excel_rates(db_vat, excel_vat)
    else:
        excel_df = pd.DataFrame()
        excel_comm = pd.Series(index=site_keys.index, dtype=float)
        excel_vat = pd.Series(index=site_keys.index, dtype=float)
        merged_comm = pd.to_numeric(db_comm, errors="coerce")
        merged_vat = pd.to_numeric(db_vat, errors="coerce")

    result.insert(insert_pos, commission_col, merged_comm)
    result.insert(insert_pos, vat_col, merged_vat)

    if verbose:
        from common.style import Color

        sites = site_keys.nunique()
        comm_db = int(pd.notna(db_comm).sum())
        vat_db = int(pd.notna(db_vat).sum())
        if excel_fallback:
            comm_excel = int((pd.isna(db_comm) & pd.notna(excel_comm)).sum())
            vat_excel = int((pd.isna(db_vat) & pd.notna(excel_vat)).sum())
            print(
                f"{Color.CYAN}[映射] 平台费/VAT：数据库优先，Excel 兜底"
                f"（行数 {len(result)}，站点 {sites}）{Color.RESET}"
            )
            print(
                f"  平台费：DB {comm_db} 行，Excel 兜底 {comm_excel} 行；"
                f"VAT：DB {vat_db} 行，Excel 兜底 {vat_excel} 行"
            )
            if excel_df.empty:
                print(f"{Color.YELLOW}  [提示] 未找到 Excel 兜底文件，仅使用数据库{Color.RESET}")
        else:
            print(
                f"{Color.CYAN}[映射] 平台费/VAT：仅数据库 platform_shop"
                f"（行数 {len(result)}，站点 {sites}）{Color.RESET}"
            )
            print(f"  平台费：DB {comm_db} 行；VAT：DB {vat_db} 行")

    return result
