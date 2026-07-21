"""B6：非 MF 尾程派送费映射。

输入：桌面「(已完成-5)订单统计-*.xlsx」（B5 产出）。
输出：「(已完成-5-1)订单统计-*.xlsx」。

处理逻辑概要：
1. 筛出「派送运费=0 且 fba费用=0 且 无 transaction-FBA 派送费」的行 → 需要定价映射。
2. 优先用 DB「goods_delivery_fee」匹配（SKU+仓+国+渠道），dispatch_fee 按 currency 换算为 EUR →「映射-单个-定价派送费」。
3. 仍未命中的行，再按「派送费-映射分类」用欧洲平台定价表做 SKU→单价映射（跳过含 MF 的分类，MF 由 B5 处理）。
4. 合并 FBA/HY/4PX 单价列（不覆盖 DB 已命中）→「映射-单个-定价派送费」×「仓库SKU销量」→「映射-定价派送费」。
5. 仍空时用本机 non_mf_fee.json 兜底；再空则把待填项追加进该 JSON，供下次手工补价后重跑。
6. 汇总写出最终「派送运费」（MF 行用 MF-派送费；ZHG 分销强制为 0）。
"""

from __future__ import annotations

import json
import importlib.util
import re
import sys
from pathlib import Path

import numpy as np
import pandas as pd
import pymysql.cursors

# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_PROJECT_ROOT = _epr_mod.bootstrap(__file__)

from common.sku_mapping import sku_mappings
from config.A0_set_date import (
    shared_date,
    folder_name,
    RMB_di_EUR,
    USD_to_EUR,
    CAD_to_EUR,
    kc_to_EUR,
    zl_to_EUR,
    Ft_to_EUR,
    Lei_to_EUR,
    kr_to_EUR,
)
from common.style import Color
from config.A0_paths import DESKTOP_ROOT

_REPORT_PRA_ROOT = next(
    (p / "reportPRA" for p in Path(__file__).resolve().parents if (p / "reportPRA").is_dir()),
    None,
)
if _REPORT_PRA_ROOT and str(_REPORT_PRA_ROOT) not in sys.path:
    sys.path.append(str(_REPORT_PRA_ROOT))

from database.db_connection import get_db_manager  # noqa: E402  # pyright: ignore[reportMissingImports]

# 本机兜底（替代原 VLOOKUP / 手动-二次映射）：定价表缺价时人工补「映射-单个-定价派送费」
NON_MF_FEE_PATH = _PROJECT_ROOT / "runtime" / "local" / "non_mf_fee.json"

# goods_delivery_fee（与 app.delivery_fee_hy 写入侧一致）
FEE_TABLE = "goods_delivery_fee"
ZIPCODE_STORE = "*"
_KEY_CHUNK = 200
_WAREHOUSE_BRACKET_RE = re.compile(r"\[.*\]\s*$")
SHIPPING_METHOD_MAP_PATH = Path(_PROJECT_ROOT) / "api" / "hy_oms" / "shipping_method_map.json"

# 匹配键：(provider_code, product_sku, dispatch_warehouse, destination_country, dispatch_channel)
_GdfKey = tuple[str, str, str, str, str]
# 无 provider 降级键
_GdfKeyNoProv = tuple[str, str, str, str]

# JSON items 与订单表共用的字段名
_COL_FEE_CLASS = "派送费-映射分类"
_COL_SKU = "SKU"
_COL_SKU_SITE = "SKU-站点识别码"
_COL_UNIT_FEE = "映射-单个-定价派送费"


def _norm_text(val) -> str:
    """空值 → ""；其余转 str 并 strip，便于做匹配键。"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).strip()


def _normalize_sku(sku) -> str:
    """与 delivery_fee_hy / sku_mapping 一致：strip；剥尾缀 -NW。"""
    if sku is None or (isinstance(sku, float) and pd.isna(sku)):
        return ""
    s = str(sku).strip()
    if s.endswith("-NW"):
        s = s[:-3]
    return s


def _parse_unit_fee(val) -> float | None:
    """解析单价；无法转 float 时返回 None（该条 JSON 不参与映射）。"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def _erp_warehouse_code(warehouse) -> str:
    """HY-OTTO-DE-01[中文名] → HY-OTTO-DE-01"""
    s = _norm_text(warehouse)
    if not s:
        return ""
    return _WAREHOUSE_BRACKET_RE.sub("", s).strip()


def _oms_warehouse_code(erp_code: str) -> str:
    """ERP 仓码 → OMS warehouse_code：-03 → DE03，其余默认 DEHY。"""
    code = erp_code.upper()
    if code.endswith("-03") or code in {"DE03", "DEHY-03", "DEHY03"}:
        return "DE03"
    return "DEHY"


def _destination_country(country, fee_class) -> str:
    """订单「国家」优先；空则从「派送费-映射分类」如 HY-DE 取 DE。"""
    c = _norm_text(country).upper()
    if c:
        return c
    fc = _norm_text(fee_class)
    if fc.upper().startswith("HY-") and len(fc) > 3:
        return fc.split("-", 1)[1].strip().upper()
    return ""


def _provider_from_fee_class(fee_class) -> str:
    """派送费-映射分类 → goods_delivery_fee.provider_code。"""
    fc = _norm_text(fee_class).upper()
    if fc.startswith("HY"):
        return "HY"
    if fc.startswith("4PX"):
        return "4PX"
    if fc.startswith("FBA"):
        return "FBA"
    return ""


def _load_shipping_method_map(path: Path = SHIPPING_METHOD_MAP_PATH) -> dict[str, str]:
    """读取 shipping_method_map.json：aliases 优先于 map → OMS shipping_method code。"""
    if not path.is_file():
        print(
            f"{Color.YELLOW}[B6] 缺少运输方式映射 {path}，"
            f"goods_delivery_fee 匹配将跳过渠道键{Color.RESET}"
        )
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        print(f"{Color.YELLOW}[B6] 无法解析运输方式映射 {path}：{exc}{Color.RESET}")
        return {}
    if not isinstance(payload, dict):
        return {}

    out: dict[str, str] = {}
    for section in ("map", "aliases"):
        raw = payload.get(section)
        if not isinstance(raw, dict):
            continue
        for k, v in raw.items():
            key = str(k).strip()
            val = str(v).strip() if v is not None else ""
            if key and val:
                out[key] = val
    return out


def _fee_to_eur(fee: float, currency: str | None) -> float | None:
    """
    将 goods_delivery_fee.dispatch_fee 换算为报表币种 EUR。
    currency 空 → 视为 EUR（HY OMS 常见）；未知币种 → None（不采纳，留给定价表/JSON）。
    """
    cur = _norm_text(currency).upper()
    if not cur or cur in {"EUR", "EU", "€"}:
        return fee
    if cur in {"USD", "US$", "US"}:
        return fee * USD_to_EUR
    if cur in {"RMB", "CNY", "CNH"}:
        return fee / RMB_di_EUR
    if cur == "CAD":
        return fee * CAD_to_EUR
    if cur == "CZK":
        return fee * kc_to_EUR
    if cur == "PLN":
        return fee * zl_to_EUR
    if cur == "HUF":
        return fee * Ft_to_EUR
    if cur == "RON":
        return fee * Lei_to_EUR
    if cur == "SEK":
        return fee * kr_to_EUR
    return None


def _fetch_goods_delivery_fee_lookup(
    skus: list[str],
) -> tuple[dict[_GdfKey, float], dict[_GdfKeyNoProv, float]]:
    """
    批量预加载 goods_delivery_fee。

    匹配键与 delivery_fee_hy 写入侧一致：
      product_sku + dispatch_warehouse + destination_country + dispatch_channel
      （destination_zipcode = '*'；同键取 dispatch_date 最新一行）

    返回：
      by_prov[(provider, sku, wh, country, channel)] → fee_EUR
      by_any[(sku, wh, country, channel)] → fee_EUR（无 provider 降级）
    """
    by_prov: dict[_GdfKey, float] = {}
    by_any: dict[_GdfKeyNoProv, float] = {}
    unique = sorted({s for s in skus if s})
    if not unique:
        return by_prov, by_any

    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor(pymysql.cursors.DictCursor)
    unknown_ccy: set[str] = set()
    try:
        for i in range(0, len(unique), _KEY_CHUNK):
            chunk = unique[i : i + _KEY_CHUNK]
            placeholders = ",".join(["%s"] * len(chunk))
            sql = f"""
                SELECT provider_code, product_sku, dispatch_warehouse,
                       destination_country, dispatch_channel,
                       dispatch_fee, currency, dispatch_date
                FROM `{FEE_TABLE}`
                WHERE product_sku IN ({placeholders})
                  AND destination_zipcode = %s
                  AND dispatch_fee IS NOT NULL
                  AND dispatch_fee > 0
                ORDER BY dispatch_date DESC
            """
            cur.execute(sql, [*chunk, ZIPCODE_STORE])
            for row in cur.fetchall():
                sku = _normalize_sku(row.get("product_sku"))
                wh = _norm_text(row.get("dispatch_warehouse"))
                country = _norm_text(row.get("destination_country")).upper()
                channel = _norm_text(row.get("dispatch_channel"))
                provider = _norm_text(row.get("provider_code")).upper()
                if not sku or not wh or not country or not channel:
                    continue
                try:
                    raw_fee = float(row["dispatch_fee"])
                except (TypeError, ValueError):
                    continue
                fee_eur = _fee_to_eur(raw_fee, row.get("currency"))
                if fee_eur is None:
                    ccy = _norm_text(row.get("currency")) or "(空)"
                    unknown_ccy.add(ccy)
                    continue
                any_key: _GdfKeyNoProv = (sku, wh, country, channel)
                if any_key not in by_any:
                    by_any[any_key] = fee_eur
                if provider:
                    prov_key: _GdfKey = (provider, sku, wh, country, channel)
                    if prov_key not in by_prov:
                        by_prov[prov_key] = fee_eur
    finally:
        cur.close()
        conn.close()

    if unknown_ccy:
        print(
            f"{Color.YELLOW}[B6] goods_delivery_fee 存在未知币种未换算为 EUR，已跳过："
            f"{', '.join(sorted(unknown_ccy))}{Color.RESET}"
        )
    return by_prov, by_any


def apply_fees_from_goods_delivery_fee(df: pd.DataFrame) -> pd.DataFrame:
    """
    优先从 goods_delivery_fee 填充「映射-单个-定价派送费」（已换算 EUR）。

    订单侧键构造（与 app.delivery_fee_hy 一致）：
      SKU           → product_sku（剥 -NW）
      仓库          → ERP 仓码去括号 → OMS DEHY/DE03
      国家/分类     → destination_country
      运输方式      → shipping_method_map → dispatch_channel
      派送费-映射分类 → provider_code（HY/4PX/FBA）

    先按 provider+四元组严格匹配，再降级为无 provider 四元组。
    """
    out = df.copy()
    if out.empty:
        out[_COL_UNIT_FEE] = pd.Series(dtype=float)
        return out
    if _COL_UNIT_FEE not in out.columns:
        out[_COL_UNIT_FEE] = np.nan

    shipping_map = _load_shipping_method_map()
    normalized_skus = [_normalize_sku(s) for s in out[_COL_SKU]]
    by_prov, by_any = _fetch_goods_delivery_fee_lookup(normalized_skus)

    fees: list[float | None] = []
    strict_hit = 0
    any_hit = 0
    skip_no_channel = 0
    skip_no_key = 0

    has_warehouse = "仓库" in out.columns
    has_country = "国家" in out.columns
    has_shipping = "运输方式" in out.columns

    for sku, (_, row) in zip(normalized_skus, out.iterrows()):
        # 已有单价（例如上游预填）不覆盖
        existing = row.get(_COL_UNIT_FEE)
        if existing is not None and not (isinstance(existing, float) and pd.isna(existing)):
            try:
                fees.append(float(existing))
                continue
            except (TypeError, ValueError):
                pass

        if not sku:
            fees.append(None)
            skip_no_key += 1
            continue

        erp_wh = _erp_warehouse_code(row.get("仓库")) if has_warehouse else ""
        oms_wh = _oms_warehouse_code(erp_wh) if erp_wh else ""
        country = (
            _destination_country(row.get("国家") if has_country else "", row.get(_COL_FEE_CLASS))
        )
        shipping_cn = _norm_text(row.get("运输方式")) if has_shipping else ""
        channel = shipping_map.get(shipping_cn, "") if shipping_cn else ""
        provider = _provider_from_fee_class(row.get(_COL_FEE_CLASS))

        if not oms_wh or not country:
            fees.append(None)
            skip_no_key += 1
            continue
        if not channel:
            fees.append(None)
            skip_no_channel += 1
            continue

        fee = None
        if provider:
            fee = by_prov.get((provider, sku, oms_wh, country, channel))
            if fee is not None:
                strict_hit += 1
        if fee is None:
            fee = by_any.get((sku, oms_wh, country, channel))
            if fee is not None:
                any_hit += 1
        fees.append(fee)

    out[_COL_UNIT_FEE] = pd.to_numeric(fees, errors="coerce")
    hit = int(out[_COL_UNIT_FEE].notna().sum())
    total = len(out)
    print(
        f"{Color.CYAN}[B6] goods_delivery_fee 映射：{hit}/{total} 行命中{_COL_UNIT_FEE}"
        f"（provider严格 {strict_hit}，无provider降级 {any_hit}；"
        f"缺渠道 {skip_no_channel}，缺仓/国/SKU {skip_no_key}）{Color.RESET}"
        f"\n  费用已按 currency 换算为 EUR（空币种按 EUR）"
    )
    return out


def _dump_fee_json(payload: dict, json_path: Path) -> None:
    """写出 JSON：meta 缩进；items 中每个 {} 占一行，便于人工逐条填价。"""
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
    # 固定字段顺序，方便 diff / 手工编辑
    field_order = (_COL_FEE_CLASS, _COL_SKU, _COL_SKU_SITE, _COL_UNIT_FEE)
    lines.append('  "items": [')
    for i, row in enumerate(items):
        if isinstance(row, dict):
            ordered = {k: row.get(k) for k in field_order}
            for k, v in row.items():
                if k not in ordered:
                    ordered[k] = v
            row = ordered
        row_json = json.dumps(row, ensure_ascii=False, separators=(", ", ": "))
        suffix = "," if i < len(items) - 1 else ""
        lines.append(f"    {row_json}{suffix}")
    lines.extend(["  ]", "}", ""])
    json_path.write_text("\n".join(lines), encoding="utf-8")


def _load_non_mf_fee_overrides(json_path: Path) -> dict[str, float]:
    """
    读取 non_mf_fee.json → {匹配键: 单价}。

    items 形如：
      [{派送费-映射分类, SKU, SKU-站点识别码, 映射-单个-定价派送费}, ...]

    匹配键优先「SKU-站点识别码」；为空则降级为「派送费-映射分类 + SKU」。
    """
    if not json_path.is_file():
        return {}
    try:
        payload = json.loads(json_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        print(f"{Color.YELLOW}[B6] 无法读取 JSON 兜底 {json_path}：{exc}{Color.RESET}")
        return {}

    items = payload.get("items") if isinstance(payload, dict) else None
    if not isinstance(items, list):
        print(f"{Color.YELLOW}[B6] JSON 缺少 items 列表，已跳过：{json_path}{Color.RESET}")
        return {}

    out: dict[str, float] = {}
    for row in items:
        if not isinstance(row, dict):
            continue
        fee = _parse_unit_fee(row.get(_COL_UNIT_FEE))
        if fee is None:
            continue  # 待填(null)或非法值：本次不命中
        key = _norm_text(row.get(_COL_SKU_SITE))
        if not key:
            key = _norm_text(row.get(_COL_FEE_CLASS)) + _norm_text(row.get(_COL_SKU))
        if key:
            out[key] = fee
    return out


def _apply_non_mf_fees_from_json(df: pd.DataFrame, json_path: Path, need_map: pd.Series) -> pd.DataFrame:
    """
    对「需要定价映射、单价仍空、且非 MF」的行，用 JSON 补全「映射-单个-定价派送费」。

    need_map：与脚本开头筛「待映射行」的 mask 一致（派送运费/fba 均为 0 且无 transaction-FBA）。
    """
    out = df.copy()
    if _COL_UNIT_FEE not in out.columns:
        out[_COL_UNIT_FEE] = np.nan

    miss_mask = (
        need_map.reindex(out.index, fill_value=False)
        & out[_COL_UNIT_FEE].isna()
        & ~out[_COL_FEE_CLASS].astype(str).str.startswith("MF", na=False)
    )
    if not miss_mask.any():
        return out

    fee_map = _load_non_mf_fee_overrides(json_path)
    if not fee_map:
        print(f"{Color.CYAN}[B6] JSON 兜底未启用或为空：{json_path}{Color.RESET}")
        return out

    miss_idx = out.index[miss_mask]
    if _COL_SKU_SITE in out.columns:
        keys = out.loc[miss_idx, _COL_SKU_SITE].map(_norm_text)
    else:
        keys = pd.Series("", index=miss_idx)
    # 识别码为空时降级：派送费-映射分类 + SKU（与 _load 侧键规则一致）
    empty_key = keys.eq("")
    if empty_key.any():
        idx = keys.index[empty_key]
        keys.loc[idx] = (
            out.loc[idx, _COL_FEE_CLASS].map(_norm_text)
            + out.loc[idx, _COL_SKU].map(_norm_text)
        )

    filled = keys.map(fee_map)
    hit_mask = filled.notna()
    n_hit = int(hit_mask.sum())
    if n_hit:
        hit_idx = filled.index[hit_mask]
        out.loc[hit_idx, _COL_UNIT_FEE] = pd.to_numeric(filled.loc[hit_idx], errors="coerce")

    remain = int((miss_mask & out[_COL_UNIT_FEE].isna()).sum())
    print(
        f"{Color.CYAN}[B6] JSON 兜底（items）：补全 {n_hit} 行；"
        f"{Color.YELLOW}仍为空 {remain} 行{Color.RESET}"
        f"\n  文件：{json_path}"
    )
    return out


def _read_non_mf_fee_payload(json_path: Path) -> dict:
    """读已有兜底 JSON；损坏或不存在时返回带空 items 的默认结构。"""
    default = {
        "version": 1,
        "description": (
            "非MF 单个定价派送费本机兜底。"
            "字段：派送费-映射分类、SKU、SKU-站点识别码、映射-单个-定价派送费。"
        ),
        "items": [],
    }
    if not json_path.is_file():
        return default
    try:
        payload = json.loads(json_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        print(f"{Color.YELLOW}[B6] 读取 {json_path} 失败，将重建：{exc}{Color.RESET}")
        return default
    if not isinstance(payload, dict):
        return default
    if not isinstance(payload.get("items"), list):
        payload["items"] = []
    payload.setdefault("version", 1)
    payload.setdefault("description", default["description"])
    return payload


def _merge_missing_into_non_mf_fee_json(df: pd.DataFrame, json_path: Path) -> int:
    """
    将「派送运费」仍空的非 MF 行追加进 non_mf_fee.json，供人工填价。

    - 已存在的键：保留原「映射-单个-定价派送费」，仅补全缺失的分类/SKU/识别码字段。
    - 新键：追加一条，映射-单个-定价派送费=null（需手工填写后重跑 B6）。
    返回本次新增条数。
    """
    miss_mask = (
        df["派送运费"].isna()
        & ~df[_COL_FEE_CLASS].astype(str).str.startswith("MF", na=False)
    )
    miss_df = df.loc[miss_mask]
    if miss_df.empty:
        return 0

    # 按匹配键去重：同一 SKU-站点识别码只生成一条待填
    pending: dict[str, dict] = {}
    for _, r in miss_df.iterrows():
        sku_site = _norm_text(r.get(_COL_SKU_SITE))
        fee_class = _norm_text(r.get(_COL_FEE_CLASS))
        sku = _norm_text(r.get(_COL_SKU))
        key = sku_site or (fee_class + sku)
        if not key or key in pending:
            continue
        pending[key] = {
            _COL_FEE_CLASS: fee_class,
            _COL_SKU: sku,
            _COL_SKU_SITE: sku_site or key,
            _COL_UNIT_FEE: None,
        }
    if not pending:
        return 0

    json_path.parent.mkdir(parents=True, exist_ok=True)
    payload = _read_non_mf_fee_payload(json_path)
    existing_items: list[dict] = []
    existing_keys: set[str] = set()

    for row in payload["items"]:
        if not isinstance(row, dict):
            continue
        key = _norm_text(row.get(_COL_SKU_SITE))
        if not key:
            key = _norm_text(row.get(_COL_FEE_CLASS)) + _norm_text(row.get(_COL_SKU))
        if not key:
            continue
        existing_keys.add(key)
        # 已有条目：用订单侧信息补全空字段，不覆盖已有单价
        if key in pending:
            src = pending[key]
            if not _norm_text(row.get(_COL_FEE_CLASS)) and src[_COL_FEE_CLASS]:
                row[_COL_FEE_CLASS] = src[_COL_FEE_CLASS]
            if not _norm_text(row.get(_COL_SKU)) and src[_COL_SKU]:
                row[_COL_SKU] = src[_COL_SKU]
            if not _norm_text(row.get(_COL_SKU_SITE)):
                row[_COL_SKU_SITE] = src[_COL_SKU_SITE]
        existing_items.append(row)

    n_added = 0
    for key, row in pending.items():
        if key in existing_keys:
            continue
        existing_items.append(row)
        n_added += 1

    existing_items.sort(key=lambda x: _norm_text(x.get(_COL_SKU_SITE)))
    payload["items"] = existing_items
    _dump_fee_json(payload, json_path)
    print(
        f"{Color.YELLOW}[B6] 已写入 {json_path}："
        f"新增待填 {n_added} 条，合计 {len(existing_items)} 条"
        f"（请填写「{_COL_UNIT_FEE}」后重跑 B6）{Color.RESET}"
    )
    return n_added


def merge_prefix_columns(df, prefix, new_col_name):
    """将以 prefix 开头的多列按行取第一个非空值，合并为 new_col_name，并删除原列。"""
    cols = [col for col in df.columns if col.startswith(prefix)]
    if cols:
        df[new_col_name] = df[cols].apply(
            lambda row: next((v for v in row if pd.notna(v)), None), axis=1
        )
        df = df.drop(columns=cols)
    else:
        df[new_col_name] = None
    return df


def _to_num(s: pd.Series) -> pd.Series:
    """费用列转数值：去千分位逗号，空串/字面 nan/None → NaN。"""
    return pd.to_numeric(
        s.astype(str).str.replace(",", "", regex=False).replace({"": np.nan, "nan": np.nan, "None": np.nan}),
        errors="coerce",
    )


# ---------------------------------------------------------------------------
# 主流程（脚本级执行）
# ---------------------------------------------------------------------------

# TODO 文件路径！！！依赖 A0 日期 / 桌面目录结构
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-5)订单统计-{shared_date}.xlsx"
main_df = pd.read_excel(main_file_path)
# 清掉历史「映射尾程-*」列，避免上次跑批残留干扰本次 sku_mappings 写出
main_df = main_df.drop(columns=[col for col in main_df.columns if col.startswith("映射尾程-")])

# 待映射：派送运费、fba 都为 0，且 transaction 未给出 FBA 派送费（HY/4PX/FBA 仓常见）
mask = (main_df["派送运费"] == 0) & (main_df["fba费用"] == 0) & (main_df["映射transaction-FBA-派送运费"].isna())
main_df_map = main_df[mask].copy()
main_df_not_map = main_df[~mask].copy()

# ---------- ① goods_delivery_fee（优先，费用已转 EUR）----------
main_df_map = apply_fees_from_goods_delivery_fee(main_df_map)
db_hit_mask = main_df_map[_COL_UNIT_FEE].notna()
main_df_map_db = main_df_map.loc[db_hit_mask].copy()
main_df_map_need_excel = main_df_map.loc[~db_hit_mask].copy()
print(
    f"{Color.CYAN}[B6] DB 已命中 {len(main_df_map_db)} 行，"
    f"其余 {len(main_df_map_need_excel)} 行进入欧洲平台定价表{Color.RESET}"
)

# ---------- ② 欧洲平台定价表（仅 DB 未命中）----------
# 表头第 2 行（iloc[1]）为「派送费-映射分类」列名，与订单分类对齐后做 VLOOKUP 式映射
product_map_sku_dir = r"\\Betohow\数据报表\2-定价表"
product_map_sku_file = "欧洲平台定价表 2026.0708.xlsx"
product_map_sku_path = fr"{product_map_sku_dir}\{product_map_sku_file}"
header_list = pd.read_excel(product_map_sku_path, sheet_name="基础表").iloc[1].fillna("").tolist()

main_df_map_excel = pd.DataFrame()
if not main_df_map_need_excel.empty:
    unique_site_classes = main_df_map_need_excel[_COL_FEE_CLASS].unique()
    for site_class in unique_site_classes:
        site_df = main_df_map_need_excel[main_df_map_need_excel[_COL_FEE_CLASS] == site_class]
        # 分类名出现在定价表表头，且不是 MF（MF 由 B5 处理）→ 按 SKU 映射该列单价
        if site_class in header_list and "MF" not in str(site_class):
            site_df_1 = sku_mappings(
                main_df=site_df,
                main_sku="SKU",
                map_sku_path=product_map_sku_path,
                map_old_sku="百途鸿SKU",
                map_new_sku=site_class,  # 定价表中对应站点/仓的单价列
                map_sku_sheet="基础表",
            )
        else:
            # 定价表无此列，或分类含 MF：原样带回，后续靠 JSON 兜底 / 其它费用列
            site_df_1 = site_df
        if main_df_map_excel.empty:
            main_df_map_excel = site_df_1
        else:
            main_df_map_excel = pd.concat([main_df_map_excel, site_df_1], ignore_index=True)

# DB 命中行 + 定价表结果
parts = [p for p in (main_df_map_db, main_df_map_excel) if not p.empty]
main_df_map_1 = pd.concat(parts, ignore_index=True) if parts else main_df_map.iloc[0:0].copy()

# 中间产物：仅含本轮「待映射」行的映射结果，便于核对
output_path = main_file_path.replace("已完成-5", "已完成-51(定价表-映射尾程)")
main_df_map_1.to_excel(output_path, index=False)

# sku_mappings 可能按仓写出 映射FBA-*/映射HY-*/映射4PX-*，收成三列单价
prefix_mapping = [
    ("映射FBA-", "单个-FBA-派送费"),
    ("映射HY-", "单个-HY-派送费"),
    ("映射4PX-", "单个-4PX-派送费"),
]
for prefix, new_name in prefix_mapping:
    main_df_map_1 = merge_prefix_columns(main_df_map_1, prefix, new_name)

# 拼回未参与定价映射的行，恢复完整订单表
main_df_1 = pd.concat([main_df_map_1, main_df_not_map], ignore_index=True)

# 三仓单价：仅当恰好 1 列非空时采纳；已有 DB/先前「映射-单个-定价派送费」优先，不被覆盖
cols = ["单个-FBA-派送费", "单个-HY-派送费", "单个-4PX-派送费"]
existing_cols = [c for c in cols if c in main_df_1.columns]
for col in existing_cols:
    main_df_1[col] = pd.to_numeric(main_df_1[col], errors="coerce")

if _COL_UNIT_FEE not in main_df_1.columns:
    main_df_1[_COL_UNIT_FEE] = np.nan
else:
    main_df_1[_COL_UNIT_FEE] = pd.to_numeric(main_df_1[_COL_UNIT_FEE], errors="coerce")

prior_unit_fee = main_df_1[_COL_UNIT_FEE].copy()
non_null_count = (
    main_df_1[existing_cols].notna().sum(axis=1) if existing_cols else pd.Series(0, index=main_df_1.index)
)
excel_unit_fee = (
    main_df_1[existing_cols].sum(axis=1, numeric_only=True) if existing_cols else pd.Series(np.nan, index=main_df_1.index)
)
excel_unit_fee = excel_unit_fee.where(non_null_count == 1, other=np.nan)
# DB（或上游）已有值优先；否则用定价表合并结果
main_df_1[_COL_UNIT_FEE] = prior_unit_fee.where(prior_unit_fee.notna(), excel_unit_fee)

# 定价表仍未命中 → 用本机 JSON 补单价（仅 need_price_map 且非 MF）
need_price_map = (
    (main_df_1["派送运费"].fillna(0) == 0)
    & (main_df_1["fba费用"].fillna(0) == 0)
    & (main_df_1["映射transaction-FBA-派送运费"].isna())
)
main_df_1 = _apply_non_mf_fees_from_json(main_df_1, NON_MF_FEE_PATH, need_price_map)

# 行金额 = 单价 × 仓库SKU销量（sale / resend 共用）
main_df_1["映射-定价派送费"] = main_df_1[_COL_UNIT_FEE] * main_df_1["仓库SKU销量"]

# 原派送运费改名后，下面用多来源费用重算「派送运费」
main_df_1 = main_df_1.rename(columns={"派送运费": "原-派送运费"})

fee_cols = [
    "fba费用",
    "原-派送运费",
    "映射transaction-FBA-派送运费",
    "映射-定价派送费",
]

for _c in [*fee_cols, "MF-派送费"]:
    if _c in main_df_1.columns:
        main_df_1[_c] = _to_num(main_df_1[_c])

# MF：沿用 B5 的 MF-派送费；非 MF：四源费用相加
mask_mf = main_df_1[_COL_FEE_CLASS].astype(str).str.startswith("MF", na=False)
main_df_1["派送运费"] = np.where(
    mask_mf,
    main_df_1["MF-派送费"],
    main_df_1[fee_cols].sum(axis=1, numeric_only=True),
)
main_df_1["派送运费"] = main_df_1["派送运费"].replace(0, np.nan)
# 分销(ZHG)：派送运费强制为 0（业务约定，非缺失）
main_df_1.loc[main_df_1[_COL_FEE_CLASS].str.startswith("ZHG", na=False), "派送运费"] = 0

# 非 MF 仍空 → 追加进 JSON 待人工填「映射-单个-定价派送费」后重跑
remain_empty = int(
    (
        main_df_1["派送运费"].isna()
        & ~main_df_1[_COL_FEE_CLASS].astype(str).str.startswith("MF", na=False)
    ).sum()
)
if remain_empty:
    _merge_missing_into_non_mf_fee_json(main_df_1, NON_MF_FEE_PATH)
else:
    print(f"{Color.GREEN}[B6] 非 MF「派送运费」已全部命中，无需 JSON 补数{Color.RESET}")

output_path = main_file_path.replace("已完成-5", "已完成-5-1")
main_df_1.to_excel(output_path, index=False)
print(f"处理完成，output_path：{output_path}")

if remain_empty:
    print(
        f"{Color.YELLOW}[请检查] 非 MF「派送运费」仍有空（{remain_empty} 行）；"
        f"请填写 {NON_MF_FEE_PATH} 中「映射-单个-定价派送费」后重跑 B6{Color.RESET}"
    )
print("================================================================================")
print(f'{Color.YELLOW}~~~（注意"仓库SKU销量"的数量）~~~~~{Color.RESET}')
print('~~~~~~~~~~~~~~~~~"欧洲平台定价表"没有的话，联系：李杨，更新定价表的数据')
