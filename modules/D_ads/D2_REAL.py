"""
D2_REAL.py — REAL 广告 CSV → (处理完成)REAL广告.xlsx

流程：
  1. 读桌面 REAL 各站点 csv（按文件名识别站点），合并；CZ 单独读并换汇
  2. 清洗 Cost (€)，去掉 0 花费
  3. EAN → product_sku_mapping(real/platform).seller_sku 换仓库 SKU
     （命中写入 SKU；未命中 SKU 留空，并写入待维护行：product_sku 空、is_active=1、check_state=0）
     bundle：按 component_info 的 qty 重复拼成 A,B,B…，再经 split_one_rows_data 按份数分摊
  4. 组合 SKU（+ / ,）拆行均摊广告费
  5. 生成 SKU-站点识别码 / SKU-平台识别码，写出处理完成 Excel

用法：
  python modules/D_ads/D2_REAL.py
"""
from __future__ import annotations

import glob
import hashlib
import importlib.util
import json
import os
from pathlib import Path
from typing import Any, Iterable

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
from config.A0_set_date import folder_name, kc_to_EUR, shared_date
from database.db_connection import get_db_manager

PSM_TABLE = "product_sku_mapping"
PLATFORM_SHOP_TABLE = "platform_shop"
PARTNER_CODE_REAL = "real"
PARTNER_TYPE_PLATFORM = "platform"
SOURCE_TYPE_REAL_ADS = "REALAds"
_KEY_CHUNK = 200
_SELLER_SKU_MAX_LEN = 128

# line_hash 参与列（与 product_sku_mapping 表注释一致）
_PSM_HASH_FIELDS = (
    "partner_code",
    "partner_type",
    "partner_name",
    "shop_hash",
    "seller_sku",
    "warehouse_sku",
    "mapping_type",
    "product_sku",
)

_SITE_ALIASES = {
    "REAL-FB": "REAL-DE-FB",
}

_SITE_MARKERS = (
    ("REAL-DE-FB", "REAL-DE-FB"),
    ("REAL-IT-FB", "REAL-IT-FB"),
    ("REAL-CZ-FB", "REAL-CZ-FB"),
    ("REAL-BTH", "REAL-BTH"),
)


# ---------------------------------------------------------------------------
# 小工具
# ---------------------------------------------------------------------------


def _as_text(val: Any) -> str:
    if pd.isna(val):
        return ""
    # EAN 常被读成 int/float，避免 1.23e11
    if isinstance(val, float) and val == int(val):
        val = int(val)
    text = str(val).strip()
    return "" if text.lower() in ("nan", "none") else text


def _unique_keys(values: Iterable[Any]) -> list[str]:
    return sorted({_as_text(v) for v in values if _as_text(v)})


def _sibling(src: str | Path, name: str) -> Path:
    return Path(src).parent / name


def _db_dict_chunks(
    keys: list[str], sql_template: str, params_prefix: tuple = ()
) -> list[dict]:
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
# CSV 读取 / CZ 花费列
# ---------------------------------------------------------------------------


def _read_csv_lines(file_path: str | Path) -> list[str]:
    """按常见编码读取 REAL 导出 csv，避免 Kč 等表头因编码错误变成乱码。"""
    raw = Path(file_path).read_bytes()
    fallback = None
    for encoding in ("utf-8-sig", "utf-8", "cp1250", "cp1252", "latin-1"):
        try:
            text = raw.decode(encoding)
        except UnicodeDecodeError:
            continue
        lines = text.splitlines()
        if lines and "Cost" in lines[0]:
            return lines
        if fallback is None:
            fallback = lines
    return fallback or []


def _site_from_filename(file_name: str) -> str:
    for marker, site in _SITE_MARKERS:
        if marker in file_name:
            return site
    return ""


def csv_to_df(file_path: str | Path) -> pd.DataFrame:
    lines = _read_csv_lines(file_path)
    new_data = []
    for line in lines:
        cells = line.strip().split(";")
        new_row = []
        for cell in cells:
            cell = cell.strip().replace('"', "").replace("\ufeff", "").replace(",", ".")
            try:
                new_row.append(int(cell))
            except ValueError:
                new_row.append(cell)
        new_data.append(new_row)

    if not new_data:
        raise ValueError(f"REAL CSV 无数据：{file_path}")

    site = _site_from_filename(os.path.basename(str(file_path)))
    if not site:
        raise SystemExit(f"无法获取到对应的站点，请检查文件名，程序终止！！！ path={file_path}")

    df = pd.DataFrame(new_data[1:], columns=new_data[0])
    df["站点"] = site
    return df


def _find_cz_cost_col(columns) -> tuple[str | None, str | None]:
    """定位 CZ 花费列：优先 Cost (Kč)，兼容编码差异或已是欧元的表头。"""
    cols = list(columns)
    if "Cost (Kč)" in cols:
        return "Cost (Kč)", "kc"
    if "Cost (€)" in cols:
        return "Cost (€)", "eur"
    for col in cols:
        col_str = str(col)
        lower = col_str.lower()
        if "cost" in lower and ("kč" in lower or "kc" in lower or "czk" in lower):
            return col, "kc"
        if "cost" in lower and "€" in col_str:
            return col, "eur"
    for col in cols:
        col_str = str(col)
        if col_str.startswith("Cost (") and col_str.endswith(")"):
            inner = col_str[6:-1].lower()
            if "eur" in inner or "€" in col_str:
                return col, "eur"
            return col, "kc"
    return None, None


# ---------------------------------------------------------------------------
# product_sku_mapping
# ---------------------------------------------------------------------------


def _parse_component_info(component_info: Any) -> list:
    """兼容 [{...}] 与 {"items": [{...}]}。"""
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
    bundle 按 qty 展开：[{A,1},{B,2}] → [A,B,B]
    后续 split_one_rows_data 按份数均摊 = 按数量分摊。
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
        return ",".join(_component_skus_by_qty(row.get("component_info")))
    return _as_text(row.get("product_sku"))


def fetch_real_ean_rows(eans: list[Any]) -> list[dict]:
    """拉取 real/platform 下相关 EAN 的映射行（updated_at 降序）。"""
    keys = _unique_keys(eans)
    sql = f"""
        SELECT seller_sku, market_region, product_sku, mapping_type, component_info, updated_at
        FROM `{PSM_TABLE}`
        WHERE partner_code = %s
          AND partner_type = 'platform'
          AND is_active = 1
          AND (
                (mapping_type = 'single' AND TRIM(product_sku) <> '')
                OR (mapping_type = 'bundle' AND component_info IS NOT NULL)
              )
          AND seller_sku IN ({{placeholders}})
        ORDER BY updated_at DESC
    """
    return _db_dict_chunks(keys, sql, (PARTNER_CODE_REAL,))


def resolve_ean_sku_for_row(
    ean: str, site: str, rows_by_ean: dict[str, list[dict]]
) -> str:
    """
    优先 market_region == 站点；否则取该 EAN 最新一条。
    REAL-FB 与 REAL-DE-FB 视为可互通。
    """
    rows = rows_by_ean.get(ean) or []
    if not rows:
        return ""
    site_norm = "REAL-DE-FB" if site == "REAL-FB" else site
    for row in rows:
        region = _as_text(row.get("market_region"))
        region_norm = "REAL-DE-FB" if region == "REAL-FB" else region
        if site_norm and region_norm == site_norm:
            return _warehouse_sku_from_row(row)
    return _warehouse_sku_from_row(rows[0])


def _normalize_site(site: str) -> str:
    site = _as_text(site)
    return _SITE_ALIASES.get(site, site)


def _psm_line_hash(
    *,
    partner_name: str,
    shop_hash: str,
    seller_sku: str,
    product_sku: str = "",
) -> str:
    record = {
        "partner_code": PARTNER_CODE_REAL,
        "partner_type": PARTNER_TYPE_PLATFORM,
        "partner_name": partner_name,
        "shop_hash": shop_hash,
        "seller_sku": seller_sku,
        "warehouse_sku": "",
        "mapping_type": "single",
        "product_sku": product_sku,
    }
    parts = [str(record.get(k) or "").strip() for k in _PSM_HASH_FIELDS]
    return hashlib.sha256("\x1f".join(parts).encode("utf-8")).hexdigest()


def _fetch_shop_by_region() -> dict[str, dict[str, str]]:
    """market_region → {shop_hash, partner_name, market_code}。"""
    sql = f"""
        SELECT shop_hash, shop_alias, market_region, market_code
        FROM `{PLATFORM_SHOP_TABLE}`
        WHERE platform = %s
          AND market_region IS NOT NULL
          AND TRIM(market_region) <> ''
    """
    mapping: dict[str, dict[str, str]] = {}
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            cur.execute(sql, (PARTNER_CODE_REAL,))
            for row in cur.fetchall():
                region = _as_text(row.get("market_region"))
                if not region:
                    continue
                mapping[region] = {
                    "shop_hash": _as_text(row.get("shop_hash")),
                    "partner_name": _as_text(row.get("shop_alias")) or "REAL",
                    "market_code": _as_text(row.get("market_code")) or "REAL",
                }
    finally:
        conn.close()
    return mapping


def _fetch_existing_real_ean_keys(eans: list[str]) -> set[tuple[str, str]]:
    """已有 real/platform 行，键为 (seller_sku, shop_hash)。"""
    keys = _unique_keys(eans)
    if not keys:
        return set()
    existing: set[tuple[str, str]] = set()
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            for i in range(0, len(keys), _KEY_CHUNK):
                chunk = keys[i : i + _KEY_CHUNK]
                placeholders = ", ".join(["%s"] * len(chunk))
                sql = f"""
                    SELECT seller_sku, shop_hash
                    FROM `{PSM_TABLE}`
                    WHERE partner_code = %s
                      AND partner_type = %s
                      AND seller_sku IN ({placeholders})
                """
                cur.execute(sql, (PARTNER_CODE_REAL, PARTNER_TYPE_PLATFORM, *chunk))
                for row in cur.fetchall():
                    ean = _as_text(row.get("seller_sku"))
                    if ean:
                        existing.add((ean, _as_text(row.get("shop_hash"))))
    finally:
        conn.close()
    return existing


def _insert_pending_real_mappings(pairs: list[tuple[str, str]]) -> tuple[int, int]:
    """
    未命中 EAN 写入 product_sku_mapping（real / platform）。

    待人工维护：product_sku 空串、is_active=1、check_state=0、mapping_type=single。
    填好 product_sku 后把 check_state 改为 1。
    同一 EAN + 店铺已存在则跳过。
    返回 (新建条数, 已存在条数)。
    """
    shop_map = _fetch_shop_by_region()
    pending: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()
    warn_no_shop: set[str] = set()
    for ean, site in pairs:
        ean = _as_text(ean)[:_SELLER_SKU_MAX_LEN]
        site = _normalize_site(site)
        if not ean:
            continue
        shop = shop_map.get(site)
        if not shop:
            warn_no_shop.add(site or "(空)")
            shop = {"shop_hash": "", "partner_name": "REAL", "market_code": "REAL"}
        key = (ean, shop["shop_hash"])
        if key in seen:
            continue
        seen.add(key)
        pending.append(
            {
                "ean": ean,
                "site": site,
                "shop_hash": shop["shop_hash"],
                "partner_name": shop["partner_name"],
                "market_code": shop["market_code"],
            }
        )
    if warn_no_shop:
        print(
            f"{Color.YELLOW}[提示] 以下站点在 platform_shop 无匹配，"
            f"shop_hash 置空：{sorted(warn_no_shop)}{Color.RESET}"
        )
    if not pending:
        return 0, 0

    existing = _fetch_existing_real_ean_keys([p["ean"] for p in pending])
    to_insert = [p for p in pending if (p["ean"], p["shop_hash"]) not in existing]
    n_exist = len(pending) - len(to_insert)
    if not to_insert:
        return 0, n_exist

    sql = f"""
        INSERT INTO `{PSM_TABLE}` (
            line_hash, partner_code, partner_type, partner_name, shop_hash,
            market_region, market_code, seller_ean, seller_sku, warehouse_sku,
            mapping_type, product_sku, component_info, source_type, is_active,
            check_state
        ) VALUES (
            %s, %s, %s, %s, %s,
            %s, %s, %s, %s, '',
            'single', '', NULL, %s, 1, 0
        )
    """
    n_new = 0
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor() as cur:
            for rec in to_insert:
                line_hash = _psm_line_hash(
                    partner_name=rec["partner_name"],
                    shop_hash=rec["shop_hash"],
                    seller_sku=rec["ean"],
                )
                try:
                    cur.execute(
                        sql,
                        (
                            line_hash,
                            PARTNER_CODE_REAL,
                            PARTNER_TYPE_PLATFORM,
                            rec["partner_name"],
                            rec["shop_hash"],
                            rec["site"] or None,
                            rec["market_code"],
                            rec["ean"],
                            rec["ean"],
                            SOURCE_TYPE_REAL_ADS,
                        ),
                    )
                    if cur.rowcount:
                        n_new += 1
                except (pymysql.err.IntegrityError, pymysql.err.OperationalError) as exc:
                    print(
                        f"{Color.RED}[写入失败] EAN={rec['ean']!r} "
                        f"site={rec['site']!r} → {exc}{Color.RESET}"
                    )
                    err = str(exc)
                    if "chk_psm_mapping_payload" in err:
                        print(
                            f"{Color.YELLOW}[提示] 请先执行 "
                            f"database/alter/alter_product_sku_mapping_allow_empty_product_sku.sql"
                            f"{Color.RESET}"
                        )
        conn.commit()
    finally:
        conn.close()
    return n_new, n_exist


def apply_ean_mapping(df: pd.DataFrame) -> pd.DataFrame:
    """
    EAN → 仓库SKU：命中写入 SKU；未命中 SKU 留空并黄字告警。
    优先按「站点 = market_region」匹配。
    """
    out = df.copy()
    ean = out["EAN"].map(_as_text)
    site = out["站点"].map(_as_text) if "站点" in out.columns else pd.Series([""] * len(out))

    db_rows = fetch_real_ean_rows(ean.tolist())
    rows_by_ean: dict[str, list[dict]] = {}
    for row in db_rows:
        key = _as_text(row.get("seller_sku"))
        if key:
            rows_by_ean.setdefault(key, []).append(row)
    print(
        f"[DB] product_sku_mapping(real/platform) 命中 "
        f"{len(rows_by_ean)} 个 EAN（{len(db_rows)} 行）"
    )

    mapped_vals = [
        resolve_ean_sku_for_row(a, s, rows_by_ean) or pd.NA
        for a, s in zip(ean.tolist(), site.tolist())
    ]
    mapped = pd.Series(mapped_vals, index=out.index)
    hit_mask = mapped.notna() & mapped.astype(str).str.strip().ne("")
    out["SKU"] = mapped.where(hit_mask, pd.NA)
    miss_mask = ~hit_mask & ean.ne("")
    print(f"[DB] 已映射 SKU {int(hit_mask.sum())} 行；未命中 {int(miss_mask.sum())} 行")

    if miss_mask.any():
        miss_pairs = list(
            {
                (a, s)
                for a, s in zip(ean[miss_mask].tolist(), site[miss_mask].tolist())
                if a
            }
        )
        miss_eans = sorted({a for a, _ in miss_pairs})
        print(
            f"{Color.YELLOW}[检查] EAN 映射未命中示例：{miss_eans[:10]}{Color.RESET}"
        )
        n_new, n_exist = _insert_pending_real_mappings(miss_pairs)
        print(
            f"{Color.YELLOW}[待维护] 已写入 `{PSM_TABLE}`："
            f"新建 {n_new} 条；库中已有 {n_exist} 条"
            f"（共 {len(miss_pairs)} 个 EAN+站点）{Color.RESET}"
        )
        print(
            f"{Color.CYAN}请人工维护 `{PSM_TABLE}`：\n"
            f"  - partner_code={PARTNER_CODE_REAL}, partner_type={PARTNER_TYPE_PLATFORM}, "
            f"source_type={SOURCE_TYPE_REAL_ADS}\n"
            f"  - 条件示例：WHERE partner_code='{PARTNER_CODE_REAL}' "
            f"AND partner_type='{PARTNER_TYPE_PLATFORM}' AND check_state=0 "
            f"AND source_type='{SOURCE_TYPE_REAL_ADS}'\n"
            f"  - 新建行 is_active=1、check_state=0；seller_sku / seller_ean 为 EAN\n"
            f"  - 填写 product_sku 后把 check_state 改为 1，再重新执行本脚本{Color.RESET}"
        )
    return out


def load_all_real_ads(folder: Path) -> tuple[pd.DataFrame, Path]:
    """合并非 CZ csv + 可选 CZ 文件；返回 (df, 任一源文件路径用于写输出)。"""
    paths = [
        f
        for f in glob.glob(str(folder / f"*{shared_date}.csv"))
        if not Path(f).name.startswith("REAL-CZ-FB")
    ]
    print(paths)
    if not paths:
        raise FileNotFoundError(f"未找到 REAL 广告 csv：{folder}\\*{shared_date}.csv")

    df = pd.concat([csv_to_df(p) for p in paths], ignore_index=True)

    # DE/IT 为欧元；CZ 为克朗，需换汇（汇率见 A0_set_date.kc_to_EUR）
    cz_path = folder / f"REAL-CZ-FB-广告数据-{shared_date}.csv"
    if cz_path.is_file():
        cz_df = csv_to_df(cz_path)
        cz_cost_col, cz_currency = _find_cz_cost_col(cz_df.columns)
        if cz_cost_col is None:
            raise KeyError(
                f"REAL-CZ-FB 未找到花费列（期望 Cost (Kč) 或 Cost (€)），"
                f"实际列名：{cz_df.columns.tolist()}"
            )
        cz_df[cz_cost_col] = cz_df[cz_cost_col].astype(float)
        if cz_currency == "kc":
            cz_df[cz_cost_col] = cz_df[cz_cost_col] * kc_to_EUR
            cz_df = cz_df.rename(columns={cz_cost_col: "Cost (€)"})
            cz_df.columns = cz_df.columns.str.replace("Kč", "€", regex=False)
            cz_df.columns = cz_df.columns.str.replace("Kc", "€", regex=False)
        elif cz_cost_col != "Cost (€)":
            cz_df = cz_df.rename(columns={cz_cost_col: "Cost (€)"})
        df = pd.concat([df, cz_df], ignore_index=True)
    else:
        print(f"未找到 CZ 广告文件，跳过：{cz_path}")

    return df, Path(paths[0])


def clean_real_ads(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in out.columns:
        out[col] = out[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
    out = out[~out["Cost (€)"].isin(["0.00", "0", 0, 0.0])].copy()
    out["Cost (€)"] = out["Cost (€)"].astype(float)
    out = out[out["Cost (€)"] != 0].copy()
    return out


def build_output(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.insert(
        out.columns.get_loc("SKU") + 1,
        "SKU-站点识别码",
        out["站点"].astype(str) + out["SKU"].astype(str),
    )
    out = map_region_to_platform(out, site_col="站点")
    out.insert(
        out.columns.get_loc("SKU-站点识别码") + 1,
        "SKU-平台识别码",
        out["映射平台"].astype(str) + out["SKU"].astype(str),
    )
    out = out[
        [
            "EAN",
            "SKU",
            "站点",
            "映射平台",
            "SKU-站点识别码",
            "SKU-平台识别码",
            "Cost (€)",
        ]
    ].rename(columns={"Cost (€)": "广告费(非AMZ)"})
    return out


def main() -> None:
    folder = Path(DESKTOP_ROOT) / f"{folder_name}{shared_date}" / "广告" / "REAL"
    df, sample_path = load_all_real_ads(folder)
    df = clean_real_ads(df)
    df = apply_ean_mapping(df)

    mid_path = _sibling(sample_path, "(已完成-1)REAL广告.xlsx")
    df.to_excel(mid_path, index=False)
    print(f"处理完成，结果已保存到{mid_path}")

    df = split_one_rows_data(
        input_df=df, data_column="SKU", value_column="Cost (€)"
    )
    result = build_output(df)

    out_path = _sibling(sample_path, "(处理完成)REAL广告.xlsx")
    result.to_excel(out_path, index=False)
    print(f"处理完成，结果已保存到{out_path}")


if __name__ == "__main__":
    main()
