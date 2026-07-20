import sys
import importlib.util
from pathlib import Path

import pandas as pd
import pymysql.cursors

# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.Z_method.style import Color
from A_报表.Z_method.sku_映射 import sku_mappings
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

_REPORT_PRA_ROOT = Path(__file__).resolve().parents[2] / "reportPRA"
if str(_REPORT_PRA_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPORT_PRA_ROOT))

from database.db_connection import get_db_manager  # noqa: E402  # pyright: ignore[reportMissingImports]

MMF_PRICE_TABLE = "mano_mmf_price"
_KEY_CHUNK = 200
PRODUCT_MAP_SKU_PATH = fr"{DESKTOP_ROOT}\MANO-MF 尾程.xlsx"

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-4)订单统计-{shared_date}.xlsx"
main_df = pd.read_excel(main_file_path)

# 筛选条件：包含 MF 站点、MF 仓库的，都要映射派送费
mask_mf_site = main_df["映射站点"].str.contains("MF", na=False)
mask_mf_category = main_df["派送费-映射分类"].str.startswith("MF", na=False)
mf_df = main_df[mask_mf_site & mask_mf_category].copy()
non_mf_df = main_df[~(mask_mf_site & mask_mf_category)].copy()


def _normalize_sku(sku) -> str:
    """
    与 sku_映射 / 价表查询一致的 SKU 规范化：
    - 去首尾空格；
    - 若以 -NW 结尾则剥掉（库存周转标记，不参与价表匹配）。
    匹配派送费时，订单 SKU 与价表 warehouse_sku / seller_sku 都按此规范后比对。
    """
    if pd.isna(sku):
        return ""
    s = str(sku).strip()
    if s.endswith("-NW"):
        s = s[:-3]
    return s


def _normalize_market_region(mapped_site) -> str:
    """
    订单「映射站点」→ 价表 market_region。
    空值返回 "" → 该行跳过 DB 匹配，留给 Excel 兜底。
    """
    if mapped_site is None or (isinstance(mapped_site, float) and pd.isna(mapped_site)):
        return ""
    return str(mapped_site).strip()


def _normalize_destination_platform(country) -> str:
    """
    订单「国家」→ 价表 destination_platform（2 位站点码，如 FR/DE/ES）。
    空值返回 "" → 该行跳过 DB 匹配。
    """
    if country is None or (isinstance(country, float) and pd.isna(country)):
        return ""
    return str(country).strip().upper()


def _normalize_platform(site) -> str:
    """
    订单「站点」→ 价表 platform（2 位站点码，如 FR/DE/ES）。
    空值返回 "" → 该行跳过 DB 匹配。
    """
    if site is None or (isinstance(site, float) and pd.isna(site)):
        return ""
    return str(site).strip().upper()


# 严格匹配键：(market_region, platform, destination_platform, 规范化SKU)
_MmfStrictKey = tuple[str, str, str, str]
# 降级匹配键：(platform, destination_platform, 规范化SKU)
_MmfPlatformKey = tuple[str, str, str]


class _MmfLookups:
    """mano_mmf_price 预加载查找表：严格 / B2B降级 / 无映射站点降级。"""

    __slots__ = ("wh_strict", "sk_strict", "wh_b2b", "sk_b2b", "wh_platform", "sk_platform")

    def __init__(self) -> None:
        self.wh_strict: dict[_MmfStrictKey, float] = {}
        self.sk_strict: dict[_MmfStrictKey, float] = {}
        self.wh_b2b: dict[_MmfPlatformKey, float] = {}
        self.sk_b2b: dict[_MmfPlatformKey, float] = {}
        self.wh_platform: dict[_MmfPlatformKey, float] = {}
        self.sk_platform: dict[_MmfPlatformKey, float] = {}


def _fetch_mmf_price_lookup(skus: list[str]) -> _MmfLookups:
    """
    批量预加载 mano_mmf_price，一次查询建成严格 + 降级查找表。

    SQL：SKU 命中 warehouse_sku/seller_sku，fee>0，platform/destination_platform 非空，
    ORDER BY dispatch_date DESC（同键保留最新价）。

    严格键 = (market_region, platform, destination_platform, sku)
    B2B降级键 = (platform, destination_platform, sku)，仅 market_region 含 B2B 的行
    无映射站点降级键 = (platform, destination_platform, sku)，不限 market_region
    """
    lookups = _MmfLookups()
    unique_skus = sorted({s for s in skus if s})
    if not unique_skus:
        return lookups

    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        for i in range(0, len(unique_skus), _KEY_CHUNK):
            chunk = unique_skus[i : i + _KEY_CHUNK]
            placeholders = ",".join(["%s"] * len(chunk))
            sql = f"""
                SELECT warehouse_sku, seller_sku, market_region, platform,
                       destination_platform, dispatch_fee_vat_exc, dispatch_date
                FROM `{MMF_PRICE_TABLE}`
                WHERE (warehouse_sku IN ({placeholders}) OR seller_sku IN ({placeholders}))
                  AND dispatch_fee_vat_exc > 0
                  AND TRIM(IFNULL(platform, '')) <> ''
                  AND TRIM(IFNULL(destination_platform, '')) <> ''
                ORDER BY dispatch_date DESC
            """
            cur.execute(sql, chunk + chunk)
            for row in cur.fetchall():
                market_region = str(row.get("market_region") or "").strip()
                platform = str(row.get("platform") or "").strip().upper()
                destination_platform = str(row.get("destination_platform") or "").strip().upper()
                if not platform or not destination_platform:
                    continue
                fee = float(row["dispatch_fee_vat_exc"])
                is_b2b_region = "B2B" in market_region.upper()
                platform_key_base = (platform, destination_platform)

                for sku_col, strict_lu, b2b_lu, plat_lu in (
                    ("warehouse_sku", lookups.wh_strict, lookups.wh_b2b, lookups.wh_platform),
                    ("seller_sku", lookups.sk_strict, lookups.sk_b2b, lookups.sk_platform),
                ):
                    sku_val = _normalize_sku(row.get(sku_col))
                    if not sku_val:
                        continue
                    plat_key = (*platform_key_base, sku_val)
                    if market_region:
                        strict_key = (market_region, *plat_key)
                        if strict_key not in strict_lu:
                            strict_lu[strict_key] = fee
                    if plat_key not in plat_lu:
                        plat_lu[plat_key] = fee
                    if is_b2b_region and plat_key not in b2b_lu:
                        b2b_lu[plat_key] = fee
    finally:
        cur.close()
        conn.close()

    return lookups


def _lookup_in_wh_sk(key, wh_lookup: dict, sk_lookup: dict) -> float | None:
    if key in wh_lookup:
        return wh_lookup[key]
    if key in sk_lookup:
        return sk_lookup[key]
    return None


def _lookup_dispatch_fee_strict(
    market_region: str,
    platform: str,
    destination_platform: str,
    sku: str,
    lookups: _MmfLookups,
) -> float | None:
    """严格匹配：映射站点 + 站点 + 国家 + SKU。"""
    if not market_region or not platform or not destination_platform or not sku:
        return None
    return _lookup_in_wh_sk(
        (market_region, platform, destination_platform, sku),
        lookups.wh_strict,
        lookups.sk_strict,
    )


def _lookup_dispatch_fee_fallback(
    mapped_site: str,
    platform: str,
    destination_platform: str,
    sku: str,
    lookups: _MmfLookups,
) -> float | None:
    """
    降级匹配（严格未命中时）：
    - 映射站点含 B2B：价表 market_region 含 B2B 即可，仍按 站点+国家+SKU
    - 否则：忽略映射站点，仅 站点+国家+SKU
    """
    if not platform or not destination_platform or not sku:
        return None
    plat_key = (platform, destination_platform, sku)
    if "B2B" in str(mapped_site or "").upper():
        return _lookup_in_wh_sk(plat_key, lookups.wh_b2b, lookups.sk_b2b)
    return _lookup_in_wh_sk(plat_key, lookups.wh_platform, lookups.sk_platform)


def _merge_mapped_tail_fee(df: pd.DataFrame) -> pd.Series:
    """从 sku_mappings 生成的「映射尾程-*」列合并出单个派送费。"""
    map_cols = [c for c in df.columns if c.startswith("映射尾程-")]
    if not map_cols:
        return pd.Series(index=df.index, dtype=float)

    def _first_non_null(row: pd.Series):
        non_null = row.dropna()
        return non_null.iloc[0] if not non_null.empty else None

    return df[map_cols].apply(_first_non_null, axis=1)


def apply_mmf_dispatch_fees_from_db(df: pd.DataFrame) -> pd.DataFrame:
    """
    优先从 DB 表 mano_mmf_price 填充「单个-MF-派送费」（不含 VAT 的派送单价）。

    ========== 匹配规则 ==========

    【严格】映射站点 + 站点 + 国家 + SKU
      映射站点 = market_region
      站点     = platform
      国家     = destination_platform
      SKU      = warehouse_sku 或 seller_sku（_normalize_sku）

    【降级】严格未命中时：
      - 映射站点含 B2B → 价表 market_region 含 B2B 即可，仍按 站点+国家+SKU
      - 否则           → 忽略映射站点，仅 站点+国家+SKU

    仍未命中 → NaN，由 apply_mmf_dispatch_fees_from_excel 兜底。
    """
    out = df.copy()
    if out.empty:
        out["单个-MF-派送费"] = pd.Series(dtype=float)
        return out

    normalized_skus = [_normalize_sku(s) for s in out["SKU"]]
    lookups = _fetch_mmf_price_lookup(normalized_skus)

    fees: list[float | None] = []
    missing_platform = set()
    missing_country = set()
    strict_hit = 0
    b2b_fallback_hit = 0
    platform_fallback_hit = 0

    for sku, (_, row) in zip(normalized_skus, out.iterrows()):
        market_region = _normalize_market_region(row.get("映射站点"))
        platform = _normalize_platform(row.get("站点"))
        destination_platform = _normalize_destination_platform(row.get("国家"))
        mapped_site_raw = str(row.get("映射站点") or "")

        if not platform:
            missing_platform.add(str(row.get("站点")))
            fees.append(None)
            continue
        if not destination_platform:
            missing_country.add(str(row.get("国家")))
            fees.append(None)
            continue
        if not sku:
            fees.append(None)
            continue

        fee = None
        if market_region:
            fee = _lookup_dispatch_fee_strict(
                market_region, platform, destination_platform, sku, lookups
            )
            if fee is not None:
                strict_hit += 1

        if fee is None:
            fee = _lookup_dispatch_fee_fallback(
                mapped_site_raw, platform, destination_platform, sku, lookups
            )
            if fee is not None:
                if "B2B" in mapped_site_raw.upper():
                    b2b_fallback_hit += 1
                else:
                    platform_fallback_hit += 1

        fees.append(fee)

    out["单个-MF-派送费"] = pd.to_numeric(fees, errors="coerce")

    hit = int(out["单个-MF-派送费"].notna().sum())
    total = len(out)
    print(
        f"{Color.CYAN}[B5] mano_mmf_price 映射：{hit}/{total} 行命中单个-MF-派送费"
        f"（严格 {strict_hit}，B2B降级 {b2b_fallback_hit}，无映射站点降级 {platform_fallback_hit}）{Color.RESET}"
    )
    if missing_platform:
        print(
            f"{Color.YELLOW}[B5] 以下行站点为空，将尝试 Excel 兜底："
            f"{', '.join(sorted(missing_platform))}{Color.RESET}"
        )
    if missing_country:
        print(
            f"{Color.YELLOW}[B5] 以下行国家为空，将尝试 Excel 兜底："
            f"{', '.join(sorted(missing_country))}{Color.RESET}"
        )
    return out


def apply_mmf_dispatch_fees_from_excel(df: pd.DataFrame, excel_path: str) -> pd.DataFrame:
    """数据库未命中的行，用 MANO-MF 尾程.xlsx 按映射站点补全。"""
    out = df.copy()
    miss_mask = out["单个-MF-派送费"].isna()
    if not miss_mask.any():
        return out

    miss_df = out.loc[miss_mask]
    filled_parts: list[pd.DataFrame] = []
    for site in miss_df["映射站点"].dropna().unique():
        site_df = miss_df[miss_df["映射站点"] == site]
        site_df_1 = sku_mappings(
            main_df=site_df,
            main_sku="SKU",
            map_sku_path=excel_path,
            map_old_sku=site,
            map_new_sku=f"尾程-{site}",
            map_sku_sheet="Sheet1",
        )
        filled_parts.append(site_df_1)

    if not filled_parts:
        return out

    filled_miss = pd.concat(filled_parts)
    excel_fees = pd.to_numeric(_merge_mapped_tail_fee(filled_miss), errors="coerce")
    still_miss = out.loc[filled_miss.index, "单个-MF-派送费"].isna()
    fill_idx = filled_miss.index[still_miss]
    out.loc[fill_idx, "单个-MF-派送费"] = excel_fees.loc[fill_idx]

    db_hit = int((~miss_mask).sum())
    excel_hit = int(out["单个-MF-派送费"].notna().sum()) - db_hit
    remain = int(out["单个-MF-派送费"].isna().sum())
    print(
        f"{Color.CYAN}[B5] MANO-MF 尾程.xlsx 兜底：补全 {excel_hit} 行；"
        f"{Color.YELLOW} 仍为空 {remain} 行{Color.RESET}"
    )

    print(
        f"{Color.YELLOW}[请检查]，「派送费-映射分类」含 MF 的「MF-派送费」是否有空 {Color.RESET}"
        f"{Color.GREEN} \n「单个-MF-派送费」 =VLOOKUP(G列,'[手动-二次映射.xlsx]MF-派送费'!$A:$B,2,FALSE) {Color.RESET}"
    )
    return out


mf_df_1 = apply_mmf_dispatch_fees_from_db(mf_df)
mf_df_1 = apply_mmf_dispatch_fees_from_excel(mf_df_1, PRODUCT_MAP_SKU_PATH)
mf_df_1["MF-派送费"] = mf_df_1["单个-MF-派送费"] * mf_df_1["仓库SKU销量"]

main_df_1 = pd.concat([mf_df_1, non_mf_df], ignore_index=True)
output_path = main_file_path.replace("已完成-4", "已完成-5")
main_df_1.to_excel(output_path, index=False)
print(f"\n 处理完成，output_path：{output_path}")
