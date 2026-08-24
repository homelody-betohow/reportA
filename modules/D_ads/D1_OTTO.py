"""
D1_OTTO.py — OTTO 广告 CSV → (处理完成)OTTO广告.xlsx

流程：
  1. 读桌面 OTTO 广告 CSV（跳过说明行，分号分隔）
  2. 落 (原始预览).xlsx，便于核对新旧导出表头
  3. 清洗 Ausgaben，去掉 0 花费
  4. 货号 Artikelnummer → product_sku_mapping(otto/platform) 换仓库 SKU
     （命中替换；未命中保留原 SKU；原 SKU/商品ID 也不在 product_sku 则黄字）
     bundle：按 component_info 的 qty 重复拼成 A,B,B…，再经 split_one_rows_data 按份数分摊广告费
  5. 组合 SKU（+ / ,）拆行均摊广告费
  6. 25- 商品ID → product_sku.product_uid 映射为产品编码
  7. 固定站点 OTTO-BTH，生成识别码，写出处理完成 Excel

用法：
  python modules/D_ads/D1_OTTO.py
"""
from __future__ import annotations

import csv
import importlib.util
import json
import re
from pathlib import Path
from typing import Any, Iterable

import chardet
import pandas as pd
import pymysql.cursors

# 须在 import config/common 之前：加载项目根到 sys.path
_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.platform_shop import map_region_to_platform
from common.split_rows_data_SKU import split_one_rows_data
from common.style import Color
from config.A0_paths import DESKTOP_ROOT
from config.A0_set_date import folder_name, shared_date
from database.db_connection import get_db_manager

PRODUCT_SKU_TABLE = "product_sku"
PSM_TABLE = "product_sku_mapping"
PARTNER_CODE_OTTO = "otto"
SITE_OTTO = "OTTO-BTH"
_KEY_CHUNK = 200
_SKU_SPLIT_RE = re.compile(r"[+,]")

# ---------------------------------------------------------------------------
# 小工具
# ---------------------------------------------------------------------------


def _as_text(val: Any) -> str:
    if pd.isna(val):
        return ""
    text = str(val).strip()
    return "" if text.lower() in ("nan", "none") else text


def _unique_keys(values: Iterable[Any]) -> list[str]:
    return sorted({_as_text(v) for v in values if _as_text(v)})


def _parse_cell(cell: str) -> Any:
    cell = cell.strip()
    try:
        return int(cell)
    except ValueError:
        try:
            return float(cell)
        except ValueError:
            return cell


def _parse_ausgaben(val: Any) -> float:
    if isinstance(val, str) and val.strip():
        return float(val.replace("€", "").replace(",", ".").strip())
    return 0.0


def _split_sku_parts(sku_val: Any) -> list[str]:
    """把 'A+B' / 'A, B' 拆成零件。"""
    text = _as_text(sku_val)
    if not text:
        return []
    return [p.strip() for p in _SKU_SPLIT_RE.split(text) if p.strip()]


def _sibling(src: str | Path, name: str) -> Path:
    return Path(src).parent / name


def _db_dict_chunks(keys: list[str], sql_template: str, params_prefix: tuple = ()) -> list[dict]:
    """按 IN 列表分片查库，返回 DictCursor 行列表。sql_template 含 {placeholders}。"""
    if not keys:
        return []
    rows: list[dict] = []
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            for i in range(0, len(keys), _KEY_CHUNK):
                chunk = keys[i : i + _KEY_CHUNK]
                placeholders = ", ".join(["%s"] * len(chunk))
                cur.execute(
                    sql_template.format(placeholders=placeholders),
                    (*params_prefix, *chunk),
                )
                rows.extend(cur.fetchall())
    finally:
        conn.close()
    return rows


# ---------------------------------------------------------------------------
# product_sku_mapping / product_sku
# ---------------------------------------------------------------------------


def _parse_component_info(component_info: Any) -> list:
    """
    把 component_info 规范成组件 list。
    兼容：
      - [{product_sku, qty}, ...]
      - {"items": [{product_sku, qty}, ...]}
    """
    if component_info is None or component_info == "":
        return []
    data = component_info
    if isinstance(data, (bytes, bytearray)):
        data = data.decode("utf-8")
    if isinstance(data, str):
        try:
            data = json.loads(data)
        except json.JSONDecodeError:
            return []
    if isinstance(data, dict):
        items = data.get("items")
        return items if isinstance(items, list) else []
    return data if isinstance(data, list) else []


def _component_qty(item: dict) -> int:
    """读取组件 qty；缺省/非法按 1；<=0 视为 0（不参与分摊）。"""
    raw = item.get("qty", 1)
    if raw is None or raw == "":
        return 1
    try:
        qty = int(float(raw))
    except (TypeError, ValueError):
        return 1
    return max(qty, 0)


def _component_skus_by_qty(component_info: Any) -> list[str]:
    """
    bundle 组件按 qty 展开为 SKU 列表，供拼成 A,B,B 后走 split_one_rows_data。
    例：[{product_sku:A,qty:1},{product_sku:B,qty:2}] → [A,B,B]
    均摊后 A=1/3、B 两行各 1/3（合计 2/3）。
    """
    parts: list[str] = []
    for item in _parse_component_info(component_info):
        if isinstance(item, dict):
            sku = _as_text(item.get("product_sku"))
            if not sku:
                continue
            qty = _component_qty(item)
            if qty <= 0:
                continue
            parts.extend([sku] * qty)
        elif isinstance(item, str) and item.strip():
            parts.append(item.strip())
    return parts


def _warehouse_sku_from_row(row: dict) -> str:
    mapping_type = _as_text(row.get("mapping_type")).lower()
    if mapping_type == "bundle":
        # 按 qty 重复拼接，后续 split_one_rows_data 按份数均摊 = 按数量分摊
        return ",".join(_component_skus_by_qty(row.get("component_info")))
    return _as_text(row.get("product_sku"))


def fetch_otto_artikel_to_warehouse_sku(artikel_nrs: list[Any]) -> dict[str, str]:
    """
    product_sku_mapping（otto / platform）：货号(seller_sku) → 仓库SKU。
    同一 seller_sku 取 updated_at 最新一条。
    """
    keys = _unique_keys(artikel_nrs)
    sql = f"""
        SELECT seller_sku, product_sku, mapping_type, component_info, updated_at
        FROM `{PSM_TABLE}`
        WHERE partner_code = %s
          AND partner_type = 'platform'
          AND is_active = 1
          AND seller_sku IN ({{placeholders}})
        ORDER BY updated_at DESC
    """
    mapping: dict[str, str] = {}
    for row in _db_dict_chunks(keys, sql, (PARTNER_CODE_OTTO,)):
        seller_sku = _as_text(row.get("seller_sku"))
        warehouse_sku = _warehouse_sku_from_row(row)
        if seller_sku and warehouse_sku and seller_sku not in mapping:
            mapping[seller_sku] = warehouse_sku
    return mapping


def fetch_existing_product_keys(skus: list[Any]) -> set[str]:
    """
    在 product_sku 表中存在的键（未删除）。
    同时匹配 product_sku 与 product_uid（OTTO 源表常见 25- 商品ID）。
    """
    keys = _unique_keys(skus)
    if not keys:
        return set()
    key_set = set(keys)
    sql = f"""
        SELECT product_sku, product_uid
        FROM `{PRODUCT_SKU_TABLE}`
        WHERE is_deleted = 0
          AND (
            product_sku IN ({{placeholders}})
            OR product_uid IN ({{placeholders}})
          )
    """
    # 同一 placeholders 用两次，参数要翻倍：由专用循环处理
    found: set[str] = set()
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            for i in range(0, len(keys), _KEY_CHUNK):
                chunk = keys[i : i + _KEY_CHUNK]
                placeholders = ", ".join(["%s"] * len(chunk))
                cur.execute(
                    sql.format(placeholders=placeholders),
                    (*chunk, *chunk),
                )
                for row in cur.fetchall():
                    for col in ("product_sku", "product_uid"):
                        val = _as_text(row.get(col))
                        if val in key_set:
                            found.add(val)
    finally:
        conn.close()
    return found


def fetch_first_sku_by_uid(uids: list[Any]) -> dict[str, str]:
    """product_uid → 最新 product_sku（按 id 降序，取第一条）。"""
    keys = _unique_keys(uids)
    sql = f"""
        SELECT product_uid, product_sku
        FROM `{PRODUCT_SKU_TABLE}`
        WHERE product_uid IN ({{placeholders}})
          AND is_deleted = 0
          AND product_sku IS NOT NULL
          AND TRIM(product_sku) <> ''
        ORDER BY id DESC
    """
    mapping: dict[str, str] = {}
    for row in _db_dict_chunks(keys, sql):
        uid = _as_text(row.get("product_uid"))
        sku = _as_text(row.get("product_sku"))
        if uid and sku and uid not in mapping:
            mapping[uid] = sku
    return mapping


# ---------------------------------------------------------------------------
# 流水线步骤
# ---------------------------------------------------------------------------


def load_otto_csv(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    raw = path.read_bytes()
    encoding = chardet.detect(raw)["encoding"]
    print(f"文件的编码是: {encoding}")

    with path.open(encoding=encoding) as f:
        next(f)
        next(f)
        reader = csv.reader(f, delimiter=";")
        rows = [[_parse_cell(c) for c in row] for row in reader]

    if not rows:
        raise ValueError(f"OTTO CSV 无数据行：{path}")
    return pd.DataFrame(rows[1:], columns=rows[0])


def save_raw_preview(df: pd.DataFrame, csv_path: str | Path) -> Path:
    out = _sibling(csv_path, f"(原始预览){Path(csv_path).stem}.xlsx")
    df.to_excel(out, index=False)
    print(f"列名: {list(df.columns)}")
    print(f"原始预览已保存到{out}")
    return out


def clean_otto_ads(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["Ausgaben"] = out["Ausgaben"].apply(_parse_ausgaben)
    out = out[out["Ausgaben"] != 0].copy()
    for col in out.columns:
        out[col] = out[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
    return out


def apply_artikel_mapping(df: pd.DataFrame) -> pd.DataFrame:
    """
    货号 → 仓库SKU：命中替换；未命中保留原 SKU。
    未命中且原 SKU/商品ID 也不在 product_sku 时黄字告警。
    """
    out = df.copy()
    artikel = out["Artikelnummer"].map(_as_text)
    sku_map = fetch_otto_artikel_to_warehouse_sku(artikel.tolist())
    print(f"[DB] product_sku_mapping(otto/platform) 命中 {len(sku_map)} 条 货号 → 仓库SKU")

    mapped = artikel.map(sku_map)
    hit_mask = mapped.notna()
    out.loc[hit_mask, "SKU"] = mapped[hit_mask]
    miss_mask = ~hit_mask & artikel.ne("")
    print(f"[DB] 已替换 SKU {int(hit_mask.sum())} 行；映射未命中保留原 SKU {int(miss_mask.sum())} 行")

    if not miss_mask.any():
        return out

    miss_rows = out.loc[miss_mask, ["Artikelnummer", "SKU"]]
    candidates: list[str] = []
    for sku_val in miss_rows["SKU"]:
        candidates.extend(_split_sku_parts(sku_val))
    existing = fetch_existing_product_keys(candidates)
    print(f"[DB] product_sku 兜底校验：候选 {len(set(candidates))} 个，命中 {len(existing)} 个")

    warn_artikels: list[str] = []
    for _, row in miss_rows.iterrows():
        art = _as_text(row["Artikelnummer"])
        parts = _split_sku_parts(row["SKU"])
        if not art:
            continue
        if not parts or any(p not in existing for p in parts):
            warn_artikels.append(art)
    warn_artikels = sorted(set(warn_artikels))

    if warn_artikels:
        print(
            f"{Color.YELLOW}[检查] 货号映射未命中且原 SKU 不在 product_sku，"
            f"示例：{warn_artikels[:10]}{Color.RESET}"
        )
    else:
        print("[DB] 映射未命中行的原 SKU 均已在 product_sku，无需告警")
    return out


def apply_product_uid_mapping(df: pd.DataFrame, sku_col: str = "SKU") -> pd.DataFrame:
    """
    SKU 以 25- 开头时视为 product_uid，映射为最新 product_sku。
    未命中保留原值；带 -NW 时剥后缀查库再缀回。
    """
    out = df.copy()
    if sku_col not in out.columns:
        raise KeyError(f"主表缺少列 {sku_col!r}，当前列: {list(out.columns)}")

    series = out[sku_col].map(_as_text)
    uid_mask = series.str.startswith("25-", na=False)
    if not uid_mask.any():
        return out

    work = series.where(uid_mask, "")
    nw_mask = work.str.endswith("-NW", na=False)
    lookup = work.mask(nw_mask, work.str.replace(r"-NW$", "", regex=True))

    uid_sku_map = fetch_first_sku_by_uid(lookup[uid_mask].tolist())
    print(f"[DB] product_sku 命中 {len(uid_sku_map)} 条 product_uid → 最新 product_sku")

    mapped = lookup.map(uid_sku_map)
    mapped = mapped.mask(nw_mask & mapped.notna(), mapped.astype(str) + "-NW")
    miss = uid_mask & mapped.isna()
    out.loc[uid_mask & mapped.notna(), sku_col] = mapped[uid_mask & mapped.notna()]

    n_miss = int(miss.sum())
    if n_miss:
        preview_cols = [c for c in (sku_col, "Artikelnummer", "Ausgaben") if c in out.columns]
        preview = out.loc[miss, preview_cols].head(10)
        print(
            f"{Color.YELLOW}[检查] 商品ID 有 {n_miss} 行未命中 product_sku"
            f"（已保留原 SKU），请核对：{Color.RESET}"
        )
        print(preview.to_string(index=False))
    return out


def build_output(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.insert(out.columns.get_loc("SKU") + 1, "SKU-站点识别码", SITE_OTTO + out["SKU"].astype(str))
    out["站点"] = SITE_OTTO
    out = map_region_to_platform(out, site_col="站点")
    out.insert(
        out.columns.get_loc("SKU-站点识别码") + 1,
        "SKU-平台识别码",
        out["映射平台"].astype(str) + out["SKU"].astype(str),
    )
    out = out[
        [
            "Artikelnummer",
            "SKU",
            "站点",
            "映射平台",
            "SKU-站点识别码",
            "SKU-平台识别码",
            "Ausgaben",
        ]
    ].rename(columns={"Ausgaben": "广告费(非AMZ)"})
    return out


def main() -> None:
    csv_path = (
        Path(DESKTOP_ROOT)
        / f"{folder_name}{shared_date}"
        / "广告"
        / "OTTO"
        / f"OTTO-广告数据-{shared_date}.csv"
    )

    df = load_otto_csv(csv_path)
    save_raw_preview(df, csv_path)
    df = clean_otto_ads(df)
    df = apply_artikel_mapping(df)

    df = split_one_rows_data(input_df=df, data_column="SKU", value_column="Ausgaben")
    mid_path = _sibling(csv_path, f"(已完成-1){csv_path.name}")
    df.to_csv(mid_path, index=False)
    print(f"处理完成，结果已保存到{mid_path}")

    df = apply_product_uid_mapping(df, sku_col="SKU")
    result = build_output(df)

    out_path = _sibling(csv_path, "(处理完成)OTTO广告.xlsx")
    result.to_excel(out_path, index=False)
    print(f"处理完成，结果已保存到{out_path}")


if __name__ == "__main__":
    main()
