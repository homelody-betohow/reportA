"""
D2_REAL.py — REAL 广告 CSV → (处理完成)REAL广告.xlsx

流程：
  1. 读桌面 REAL 各站点 csv（按文件名识别站点），合并；CZ 单独读并换汇
  2. 清洗 Cost (€)，去掉 0 花费
  3. EAN → product_sku_mapping(real/platform).seller_sku 换仓库 SKU
     （命中写入 SKU；未命中 SKU 留空并黄字列 EAN）
     bundle：按 component_info 的 qty 重复拼成 A,B,B…，再经 split_one_rows_data 按份数分摊
  4. 组合 SKU（+ / ,）拆行均摊广告费
  5. 生成 SKU-站点识别码 / SKU-平台识别码，写出处理完成 Excel

用法：
  python modules/D_ads/D2_REAL.py
"""
from __future__ import annotations

import glob
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
PARTNER_CODE_REAL = "real"
_KEY_CHUNK = 200

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
        miss_eans = sorted({a for a in ean[miss_mask].tolist() if a})
        print(
            f"{Color.YELLOW}[检查] EAN 映射未命中示例：{miss_eans[:10]}{Color.RESET}"
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
