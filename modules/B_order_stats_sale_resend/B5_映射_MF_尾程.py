import json
import sys
import importlib.util
from pathlib import Path

import pandas as pd
import pymysql.cursors

# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_PROJECT_ROOT = _epr_mod.bootstrap(__file__)

from common.style import Color
from common.sku_mapping import sku_mappings
from config.A0_set_date import shared_date, folder_name
from config.A0_paths import DESKTOP_ROOT

from database.db_connection import get_db_manager  # noqa: E402

MMF_PRICE_TABLE = "mano_mmf_price"
_KEY_CHUNK = 200

# False：屏蔽 MANO-MF 尾程.xlsx；DB 未命中直接走 JSON 兜底
USE_MMF_EXCEL_FALLBACK = False
PRODUCT_MAP_SKU_PATH = fr"{DESKTOP_ROOT}\MANO-MF 尾程.xlsx"

# 本机兜底（字段列表）：映射站点 / SKU / SKU-站点识别码 / 单个-MF-派送费
MMF_FEE_OVERRIDES_PATH = _PROJECT_ROOT / "runtime" / "local" / "mano_mf_fee.json"

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
    与 sku_mapping / 价表查询一致的 SKU 规范化：
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
    """数据库未命中的行，用 MANO-MF 尾程.xlsx 按映射站点补全。

    Excel 列为「映射站点」/「尾程-{映射站点}」成对出现（如 MANO-FR-OHPAMF）。
    价表无该站点列时跳过（不中断），留给 JSON 兜底。
    """
    out = df.copy()
    miss_mask = out["单个-MF-派送费"].isna()
    if not miss_mask.any():
        return out

    miss_df = out.loc[miss_mask]
    try:
        excel_cols = set(
            pd.read_excel(excel_path, sheet_name="Sheet1", nrows=0).columns.astype(str)
        )
    except Exception as exc:
        print(
            f"{Color.YELLOW}[B5] 无法读取 MANO-MF 尾程.xlsx，跳过 Excel 兜底："
            f"{exc}{Color.RESET}"
        )
        return out

    filled_parts: list[pd.DataFrame] = []
    skipped_sites: list[str] = []
    for site in miss_df["映射站点"].dropna().unique():
        site = str(site).strip()
        fee_col = f"尾程-{site}"
        if site not in excel_cols or fee_col not in excel_cols:
            skipped_sites.append(site)
            continue
        site_df = miss_df[miss_df["映射站点"] == site]
        site_df_1 = sku_mappings(
            main_df=site_df,
            main_sku="SKU",
            map_sku_path=excel_path,
            map_old_sku=site,
            map_new_sku=fee_col,
            map_sku_sheet="Sheet1",
        )
        filled_parts.append(site_df_1)

    if skipped_sites:
        print(
            f"{Color.YELLOW}[B5] MANO-MF 尾程.xlsx 无下列映射站点列，已跳过"
            f"（改走 JSON 兜底）：{', '.join(sorted(skipped_sites))}{Color.RESET}"
        )

    if not filled_parts:
        remain = int(out["单个-MF-派送费"].isna().sum())
        print(
            f"{Color.CYAN}[B5] MANO-MF 尾程.xlsx 兜底：无可映射站点列；"
            f"{Color.YELLOW}仍为空 {remain} 行{Color.RESET}"
        )
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
        f"{Color.YELLOW}仍为空 {remain} 行{Color.RESET}"
    )
    return out


_COL_MAP_SITE = "映射站点"
_COL_SKU = "SKU"
_COL_SKU_SITE = "SKU-站点识别码"
_COL_UNIT_FEE = "单个-MF-派送费"


def _norm_text(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).strip()


def _parse_unit_fee(val) -> float | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def _dump_mmf_fee_json(payload: dict, json_path: Path) -> None:
    """
    写出 MF 费用 JSON：items 中每个 {} 占一行，便于对照编辑。
    例：
      {
        "version": 1,
        "description": "...",
        "items": [
          {"映射站点": "...", "SKU": "...", "SKU-站点识别码": "...", "单个-MF-派送费": 1.23},
          {"映射站点": "...", "SKU": "...", "SKU-站点识别码": "...", "单个-MF-派送费": null}
        ]
      }
    """
    items = payload.get("items") if isinstance(payload, dict) else None
    if not isinstance(items, list):
        json_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        return

    meta = {k: v for k, v in payload.items() if k != "items"}
    lines = ["{"]
    for key, val in meta.items():
        lines.append(
            f"  {json.dumps(key, ensure_ascii=False)}: "
            f"{json.dumps(val, ensure_ascii=False)},"
        )
    item_field_order = (_COL_MAP_SITE, _COL_SKU, _COL_SKU_SITE, _COL_UNIT_FEE)
    lines.append('  "items": [')
    for i, row in enumerate(items):
        if isinstance(row, dict):
            ordered = {k: row.get(k) for k in item_field_order}
            for k, v in row.items():
                if k not in ordered:
                    ordered[k] = v
            row = ordered
        row_json = json.dumps(row, ensure_ascii=False, separators=(", ", ": "))
        suffix = "," if i < len(items) - 1 else ""
        lines.append(f"    {row_json}{suffix}")
    lines.append("  ]")
    lines.append("}")
    lines.append("")
    json_path.write_text("\n".join(lines), encoding="utf-8")


def _load_mmf_fee_overrides(json_path: Path) -> dict[str, float]:
    """
    读取 runtime/local JSON 字段列表（items 中每个 {} 一行亦可）：
      {
        "version": 1,
        "items": [
          {"映射站点": "...", "SKU": "...", "SKU-站点识别码": "...", "单个-MF-派送费": 1.23}
        ]
      }
    匹配键优先「SKU-站点识别码」；若为空则用「映射站点」+「SKU」。
    费用为 null / 非数字则跳过。
    """
    if not json_path.is_file():
        return {}
    try:
        payload = json.loads(json_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        print(f"{Color.YELLOW}[B5] 无法读取 JSON 兜底 {json_path}：{exc}{Color.RESET}")
        return {}

    items = payload.get("items") if isinstance(payload, dict) else None
    if not isinstance(items, list):
        print(
            f"{Color.YELLOW}[B5] JSON 缺少 items 列表，已跳过：{json_path}{Color.RESET}"
        )
        return {}

    out: dict[str, float] = {}
    for row in items:
        if not isinstance(row, dict):
            continue
        fee = _parse_unit_fee(row.get(_COL_UNIT_FEE))
        if fee is None:
            continue
        key = _norm_text(row.get(_COL_SKU_SITE))
        if not key:
            key = _norm_text(row.get(_COL_MAP_SITE)) + _norm_text(row.get(_COL_SKU))
        if not key:
            continue
        out[key] = fee
    return out


def apply_mmf_dispatch_fees_from_json(df: pd.DataFrame, json_path: Path) -> pd.DataFrame:
    """
    DB + Excel 仍未命中时，用 runtime/local JSON（items 字段列表）按
    「SKU-站点识别码」补全「单个-MF-派送费」。
    """
    out = df.copy()
    if out.empty or _COL_UNIT_FEE not in out.columns:
        return out
    if _COL_SKU_SITE not in out.columns:
        print(
            f"{Color.YELLOW}[B5] 缺少列「{_COL_SKU_SITE}」，跳过 JSON 兜底{Color.RESET}"
        )
        return out

    miss_mask = out[_COL_UNIT_FEE].isna()
    if not miss_mask.any():
        return out

    fee_map = _load_mmf_fee_overrides(json_path)
    if not fee_map:
        print(
            f"{Color.CYAN}[B5] JSON 兜底未启用或为空：{json_path}{Color.RESET}"
        )
        return out

    keys = out.loc[miss_mask, _COL_SKU_SITE].map(_norm_text)
    filled = keys.map(fee_map)
    hit_mask = filled.notna()
    n_hit = int(hit_mask.sum())
    if n_hit:
        hit_idx = filled.index[hit_mask]
        out.loc[hit_idx, _COL_UNIT_FEE] = pd.to_numeric(filled.loc[hit_idx], errors="coerce")

    remain = int(out[_COL_UNIT_FEE].isna().sum())
    print(
        f"{Color.CYAN}[B5] JSON 兜底（items 字段）：补全 {n_hit} 行；"
        f"{Color.YELLOW}仍为空 {remain} 行{Color.RESET}"
        f"\n  文件：{json_path}"
    )
    return out


def _read_mmf_fee_payload(json_path: Path) -> dict:
    """读取 mano_mf_fee.json；文件不存在或损坏时返回空骨架。"""
    default = {
        "version": 1,
        "description": (
            "MF 单个派送费本机兜底（不含 VAT）。"
            "字段：映射站点、SKU、SKU-站点识别码、单个-MF-派送费。"
        ),
        "items": [],
    }
    if not json_path.is_file():
        return default
    try:
        payload = json.loads(json_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        print(f"{Color.YELLOW}[B5] 读取 {json_path} 失败，将重建：{exc}{Color.RESET}")
        return default
    if not isinstance(payload, dict):
        return default
    if not isinstance(payload.get("items"), list):
        payload["items"] = []
    payload.setdefault("version", 1)
    payload.setdefault("description", default["description"])
    return payload


def _merge_missing_into_mmf_fee_json(df: pd.DataFrame, json_path: Path) -> int:
    """
    将仍缺「单个-MF-派送费」的记录直接追加进 mano_mf_fee.json。
    - 已存在的 SKU-站点识别码：保留原费用与原顺序，仅补空的 映射站点/SKU；
    - 不存在的：追加到 items 末尾，单个-MF-派送费=null，待手工填写后重跑 B5。
    返回新追加条数。
    """
    miss_df = df.loc[df[_COL_UNIT_FEE].isna()].copy()
    if miss_df.empty or _COL_SKU_SITE not in miss_df.columns:
        return 0

    pending: dict[str, dict] = {}
    for _, r in miss_df.iterrows():
        sku_site = _norm_text(r.get(_COL_SKU_SITE))
        if not sku_site or sku_site in pending:
            continue
        pending[sku_site] = {
            _COL_MAP_SITE: _norm_text(r.get(_COL_MAP_SITE)),
            _COL_SKU: _norm_text(r.get(_COL_SKU)),
            _COL_SKU_SITE: sku_site,
            _COL_UNIT_FEE: None,
        }
    if not pending:
        return 0

    json_path.parent.mkdir(parents=True, exist_ok=True)
    payload = _read_mmf_fee_payload(json_path)
    existing_items: list[dict] = []
    existing_keys: set[str] = set()

    for row in payload["items"]:
        if not isinstance(row, dict):
            continue
        key = _norm_text(row.get(_COL_SKU_SITE))
        if not key:
            key = _norm_text(row.get(_COL_MAP_SITE)) + _norm_text(row.get(_COL_SKU))
        if not key:
            continue
        existing_keys.add(key)
        # 用订单侧信息补全空字段，不覆盖已有费用
        if key in pending:
            src = pending[key]
            if not _norm_text(row.get(_COL_MAP_SITE)) and src[_COL_MAP_SITE]:
                row[_COL_MAP_SITE] = src[_COL_MAP_SITE]
            if not _norm_text(row.get(_COL_SKU)) and src[_COL_SKU]:
                row[_COL_SKU] = src[_COL_SKU]
            row[_COL_SKU_SITE] = key
        existing_items.append(row)

    # 新增待填一律追加到文档末尾（不重排已有 items，便于对照手工填写）
    n_added = 0
    for key, row in pending.items():
        if key in existing_keys:
            continue
        existing_items.append(row)
        n_added += 1

    payload["items"] = existing_items
    _dump_mmf_fee_json(payload, json_path)
    print(
        f"{Color.YELLOW}[B5] 已写入 {json_path}："
        f"新增待填 {n_added} 条（已追加到 items 末尾），"
        f"合计 {len(existing_items)} 条"
        f"（请填写「单个-MF-派送费」后重跑 B5）{Color.RESET}"
    )
    return n_added


mf_df_1 = apply_mmf_dispatch_fees_from_db(mf_df)
if USE_MMF_EXCEL_FALLBACK:
    mf_df_1 = apply_mmf_dispatch_fees_from_excel(mf_df_1, PRODUCT_MAP_SKU_PATH)
else:
    remain = int(mf_df_1["单个-MF-派送费"].isna().sum()) if not mf_df_1.empty else 0
    print(
        f"{Color.CYAN}[B5] 已屏蔽 MANO-MF 尾程.xlsx；"
        f"{Color.YELLOW}DB 未命中 {remain} 行改走 JSON 兜底{Color.RESET}"
    )
mf_df_1 = apply_mmf_dispatch_fees_from_json(mf_df_1, MMF_FEE_OVERRIDES_PATH)
mf_df_1["MF-派送费"] = mf_df_1["单个-MF-派送费"] * mf_df_1["仓库SKU销量"]

remain_empty = int(mf_df_1["单个-MF-派送费"].isna().sum()) if not mf_df_1.empty else 0
if remain_empty:
    _merge_missing_into_mmf_fee_json(mf_df_1, MMF_FEE_OVERRIDES_PATH)
    print(
        f"{Color.YELLOW}[请检查]「派送费-映射分类」含 MF 的「MF-派送费」仍有空 "
        f"（{remain_empty} 行）{Color.RESET}"
        f"\n{Color.GREEN}请直接编辑 {MMF_FEE_OVERRIDES_PATH}"
        f"\n  填写 items 中「单个-MF-派送费」后重跑 B5{Color.RESET}"
    )
else:
    print(
        f"{Color.GREEN}[B5] MF 行「单个-MF-派送费」已全部命中，无需 JSON 补数{Color.RESET}"
    )

main_df_1 = pd.concat([mf_df_1, non_mf_df], ignore_index=True)
output_path = main_file_path.replace("已完成-4", "已完成-5")
main_df_1.to_excel(output_path, index=False)
print(f"\n 处理完成，output_path：{output_path}")
