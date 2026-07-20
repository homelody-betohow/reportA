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
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

_REPORT_PRA_ROOT = Path(__file__).resolve().parents[2] / "reportPRA"
if str(_REPORT_PRA_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPORT_PRA_ROOT))

from database.db_connection import get_db_manager  # noqa: E402  # pyright: ignore[reportMissingImports]

SHIPPED_TABLE = "sales_order_shipped"
_KEY_CHUNK = 200

main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-1)订单统计-{shared_date}.xlsx"
main_df = pd.read_excel(main_file_path)


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


def fetch_order_site_from_db(order_nos: list[str], shop_name_en: str) -> dict[str, str]:
    """从 sales_order_shipped 按 order_no 查站点/国家，优先 country，为空时用 platform_site。"""
    order_nos = sorted({str(x).strip() for x in order_nos if x and str(x).strip()})
    if not order_nos:
        return {}

    sql = f"""
        SELECT order_no, country, platform_site
        FROM `{SHIPPED_TABLE}`
        WHERE shop_name_en = %s
          AND order_no IN ({{placeholders}})
        ORDER BY id ASC
    """
    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        rows = _chunked_in_query(cur, sql, order_nos, (shop_name_en,))
    finally:
        cur.close()
        conn.close()

    mapping: dict[str, str] = {}
    for row in rows:
        order_no = str(row.get("order_no") or "").strip()
        if not order_no or order_no in mapping:
            continue
        site = str(row.get("country") or "").strip() or str(row.get("platform_site") or "").strip()
        if site:
            mapping[order_no] = site
    return mapping


def fetch_order_platform_sku_from_db(order_nos: list[str]) -> dict[str, str]:
    """从 sales_order_shipped 按 order_no 查 platform_sku（不限店铺，每单保留首条）。"""
    order_nos = sorted({str(x).strip() for x in order_nos if x and str(x).strip()})
    if not order_nos:
        return {}

    sql = f"""
        SELECT order_no, platform_sku
        FROM `{SHIPPED_TABLE}`
        WHERE order_no IN ({{placeholders}})
          AND platform_sku IS NOT NULL
          AND TRIM(platform_sku) <> ''
        ORDER BY id ASC
    """
    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        rows = _chunked_in_query(cur, sql, order_nos)
    finally:
        cur.close()
        conn.close()

    mapping: dict[str, str] = {}
    for row in rows:
        order_no = str(row.get("order_no") or "").strip()
        if not order_no or order_no in mapping:
            continue
        mapping[order_no] = str(row.get("platform_sku") or "").strip()
    return mapping


# 一、映射 REAL-FB 没有站点（国家）的
fb_mask = (main_df["店铺英文名"] == "FB_REAL") & (main_df["站点"] == "UNKNOW")
fb_order_nos = main_df.loc[fb_mask, "订单号"].dropna().astype(str).str.strip().tolist()
site_map = fetch_order_site_from_db(fb_order_nos, "FB_REAL")

mapped_site = main_df["订单号"].astype(str).str.strip().map(site_map)
if site_map:
    print(f"{Color.GREEN}[DB] REAL-FB 从 sales_order_shipped 查到 {len(site_map)} 个订单站点{Color.RESET}")
    update_mask = fb_mask & mapped_site.notna()
    main_df.loc[update_mask, "站点"] = mapped_site[update_mask]

unmapped_fb = fb_mask & mapped_site.isna()
# if unmapped_fb.any():
#     print(f"{Color.YELLOW}[DB] REAL-FB 仍有 {unmapped_fb.sum()} 单未映射站点，请检查 sales_order_shipped{Color.RESET}")

# 二、映射所有重发订单对应原订单的平台SKU
resend_mask = main_df["订单类型"] == "重发订单"
shops_with_unmapped: list[str] = []
_FOCUS_RESEND_SHOPS = {"LM_BC_FR", "LM_RP_FR"}

if not resend_mask.any():
    print("\n没有重发订单！\n")
else:
    order_no_clean = main_df.loc[resend_mask, "订单号"].astype(str).str.replace(r"-\d$", "", regex=True)
    orig_order_nos = order_no_clean.dropna().unique().tolist()
    sku_map = fetch_order_platform_sku_from_db(orig_order_nos)

    print(f"{Color.GREEN}[DB] 重发订单：从 sales_order_shipped 查到 {len(sku_map)} 个原单平台SKU{Color.RESET}")

    matched = resend_mask & order_no_clean.isin(sku_map)
    main_df.loc[matched, "平台sku"] = order_no_clean[matched].map(sku_map)
    main_df.loc[matched, "订单号"] += '——已映射"' + main_df.loc[matched, "平台sku"].astype(str) + '"'

    unmatched = resend_mask & ~matched
    if unmatched.any():
        # 只提示“之前需要重点检查的店铺”，避免把所有店铺都列出来造成干扰
        shops_with_unmapped = sorted(
            set(main_df.loc[unmatched, "店铺英文名"].dropna().unique().tolist()) & _FOCUS_RESEND_SHOPS
        )
        # print(f"{Color.YELLOW}[DB] 仍有 {unmatched.sum()} 条重发订单未映射平台SKU{Color.RESET}")

# 保存结果
output_path = main_file_path.replace("已完成-1", "已完成-1-1")
main_df.to_excel(output_path, index=False)
print(f"处理完成，文件另存为：{output_path}")
print("-" * 100)
if shops_with_unmapped:
    for shop in shops_with_unmapped:
        print(f'{Color.YELLOW}~~~~~~~~~~~~~~~~~请检查，"店铺英文名" == {shop}，"重发订单"是否都已映射 "平台SKU"！！！{Color.RESET}')
elif resend_mask.any():
    print(f"{Color.GREEN}[成功] 所有重发订单均已映射平台SKU！{Color.RESET}")
