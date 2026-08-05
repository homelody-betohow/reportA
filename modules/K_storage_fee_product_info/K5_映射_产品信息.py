"""
K5_映射_产品信息 — 订单统计补全产品属性（16→17）

【流水线位置】
  上游：(已完成-16)订单统计（K1_4；月报时 MANO 仓租可能已由 C2 写入「FBA仓租费」）
  下游：(已完成-17) → K6 分销收尾 + 无平台仓租分摊

【映射来源】
  1. 运营模式 / 供应商 / 二级分类 / 三级分类 / 产品状态
     ← DB `product_sku`
       （ops_model / supplier_abbr / category_lv2 / category_lv3 /
        amz_lifecycle 或 local_lifecycle）
  2. 产品状态仍为空 ← 桌面「信息-映射.xlsx」兜底
  3. 业务规则收尾：U88、-NW、智慧谷/易速、分销等

用法：
  python modules/K_storage_fee_product_info/K5_映射_产品信息.py
"""

from __future__ import annotations

import importlib.util
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

from common.sku_mapping import sku_mappings  # noqa: E402
from common.style import Color  # noqa: E402
from config.A0_paths import DESKTOP_ROOT  # noqa: E402
from config.A0_set_date import folder_name, shared_date  # noqa: E402
from database.db_connection import get_db_manager  # noqa: E402

PRODUCT_SKU_TABLE = "product_sku"
_KEY_CHUNK = 500
_BLANK = frozenset({"", "nan", "None", "NaN", "none"})

INFO_MAP_XLSX = Path(DESKTOP_ROOT) / "信息-映射.xlsx"


def _norm_sku_key(series: pd.Series) -> pd.Series:
    """与 sku_mappings 一致：转字符串、去空格、剥 -NW 再查库。"""
    s = series.map(lambda v: "" if pd.isna(v) else str(v).strip())
    s = s.mask(s.isin(_BLANK), "")
    return s.str.replace(r"-NW$", "", regex=True)


def _chunked_in_query(cur, sql_template: str, keys: list[str]) -> list[dict]:
    if not keys:
        return []
    results: list[dict] = []
    for i in range(0, len(keys), _KEY_CHUNK):
        chunk = keys[i : i + _KEY_CHUNK]
        placeholders = ",".join(["%s"] * len(chunk))
        cur.execute(sql_template.format(placeholders=placeholders), tuple(chunk))
        results.extend(cur.fetchall())
    return results


def fetch_product_sku_attrs(sku_keys: list[str]) -> dict[str, dict[str, str]]:
    """
    从 product_sku 批量查产品属性。
    返回 product_sku → {
      ops_model, supplier_abbr, category_lv2, category_lv3,
      amz_lifecycle, local_lifecycle
    }
    """
    sku_keys = sorted({str(x).strip() for x in sku_keys if x and str(x).strip()})
    if not sku_keys:
        return {}

    sql = f"""
        SELECT product_sku, ops_model, supplier_abbr,
               category_lv2, category_lv3,
               amz_lifecycle, local_lifecycle
        FROM `{PRODUCT_SKU_TABLE}`
        WHERE product_sku IN ({{placeholders}})
          AND COALESCE(is_deleted, 0) = 0
    """
    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        rows = _chunked_in_query(cur, sql, sku_keys)
    finally:
        cur.close()
        conn.close()

    out: dict[str, dict[str, str]] = {}
    for row in rows:
        sku = str(row.get("product_sku") or "").strip()
        if not sku:
            continue
        out[sku] = {
            "ops_model": str(row.get("ops_model") or "").strip(),
            "supplier_abbr": str(row.get("supplier_abbr") or "").strip(),
            "category_lv2": str(row.get("category_lv2") or "").strip(),
            "category_lv3": str(row.get("category_lv3") or "").strip(),
            "amz_lifecycle": str(row.get("amz_lifecycle") or "").strip(),
            "local_lifecycle": str(row.get("local_lifecycle") or "").strip(),
        }
    return out


def apply_product_info_from_db(df: pd.DataFrame) -> pd.DataFrame:
    """
    按 SKU 写入运营模式 / 供应商 / 二级分类 / 三级分类 / 产品状态（来自 product_sku）。
    产品状态：平台含 AMAZON → amz_lifecycle，否则 → local_lifecycle。
    """
    if "SKU" not in df.columns:
        raise KeyError("订单统计缺少列「SKU」")
    if "平台" not in df.columns:
        raise KeyError("订单统计缺少列「平台」")

    out = df.copy()
    for col in ("运营模式", "供应商", "二级分类", "三级分类", "产品状态"):
        if col in out.columns:
            out = out.drop(columns=[col])

    keys = _norm_sku_key(out["SKU"])
    unique_keys = sorted({k for k in keys.tolist() if k})
    mapping = fetch_product_sku_attrs(unique_keys)
    is_amazon = out["平台"].astype(str).str.contains("AMAZON", case=False, na=False)

    ops_list: list[str | None] = []
    supplier_list: list[str | None] = []
    lv2_list: list[str | None] = []
    lv3_list: list[str | None] = []
    status_list: list[str | None] = []
    hit_ops = 0
    hit_cat = 0
    hit_status = 0

    for k, amazon in zip(keys.tolist(), is_amazon.tolist()):
        attrs = mapping.get(k) if k else None
        if not attrs:
            ops_list.append(None)
            supplier_list.append(None)
            lv2_list.append(None)
            lv3_list.append(None)
            status_list.append(None)
            continue

        ops = attrs["ops_model"] or None
        supplier = attrs["supplier_abbr"] or None
        lv2 = attrs["category_lv2"] or None
        lv3 = attrs["category_lv3"] or None
        status = (attrs["amz_lifecycle"] if amazon else attrs["local_lifecycle"]) or None

        ops_list.append(ops)
        supplier_list.append(supplier)
        lv2_list.append(lv2)
        lv3_list.append(lv3)
        status_list.append(status)

        if ops or supplier:
            hit_ops += 1
        if lv2 or lv3:
            hit_cat += 1
        if status:
            hit_status += 1

    sku_pos = out.columns.get_loc("SKU")
    out.insert(sku_pos + 1, "运营模式", ops_list)
    out.insert(sku_pos + 2, "供应商", supplier_list)
    out.insert(sku_pos + 3, "二级分类", lv2_list)
    out.insert(sku_pos + 4, "三级分类", lv3_list)
    out.insert(sku_pos + 5, "产品状态", status_list)
    print(
        f"[DB] product_sku 映射："
        f"唯一SKU={len(unique_keys)}，"
        f"运营/供应商命中行={hit_ops}/{len(out)}，"
        f"分类命中行={hit_cat}/{len(out)}，"
        f"产品状态命中行={hit_status}/{len(out)}"
    )
    return out


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

main_file_path = (
    fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计"
    fr"\(已完成-16)订单统计-{shared_date}.xlsx"
)
main_file_df = pd.read_excel(main_file_path)

# 1. 运营模式、供应商、二级/三级分类、产品状态 ← product_sku
main_file_df_5 = apply_product_info_from_db(main_file_df)

# 2. 产品状态为空 → 桌面「信息-映射.xlsx」兜底
no_state_mask = main_file_df_5["产品状态"].isna()
no_state_df = main_file_df_5.loc[no_state_mask].copy()
if not no_state_df.empty and INFO_MAP_XLSX.is_file():
    no_state_df_1 = sku_mappings(
        main_df=no_state_df,
        main_sku="SKU",
        map_sku_path=str(INFO_MAP_XLSX),
        map_old_sku="SKU",
        map_new_sku="产品状态",
        map_sku_sheet="产品状态",
    )
    # 按原 index 写回，避免 DataFrame.update 对不齐
    filled = no_state_df_1["映射产品状态"]
    main_file_df_5.loc[filled.index, "产品状态"] = filled
elif not no_state_df.empty:
    print(f"{Color.YELLOW}[警告] 未找到 {INFO_MAP_XLSX}，跳过产品状态兜底映射{Color.RESET}")

# 产品状态为空，且 分销为 是 的行，产品状态改为 分销
if "分销" in main_file_df_5.columns:
    main_file_df_5.loc[
        (main_file_df_5["产品状态"].isna()) & (main_file_df_5["分销"] == "是"),
        "产品状态",
    ] = "分销"

# 业务规则收尾
main_file_df_5.loc[
    main_file_df_5["SKU"].astype(str).str.startswith("U88"),
    ["产品状态", "二级分类", "三级分类"],
] = ["新品", "其他", "其他"]
main_file_df_5.loc[
    main_file_df_5["SKU"].astype(str).str.endswith("-NW"), "产品状态"
] = "不保留老品"
main_file_df_5.loc[
    main_file_df_5["SKU"].astype(str).str.match(r"^(25|SN25|207)"), "供应商"
] = "智慧谷"
main_file_df_5.loc[
    main_file_df_5["供应商"].isin(["易速", "智慧谷"]),
    ["产品状态", "二级分类", "三级分类"],
] = "分销"
main_file_df_5.loc[
    main_file_df_5["产品状态"] == "分销",
    ["运营模式", "二级分类", "三级分类"],
] = ["自运营", "分销", "分销"]

# 保存
output_path = main_file_path.replace("已完成-16", "已完成-17")
try:
    main_file_df_5.to_excel(output_path, index=False)
    print(f"处理完成，文件另存为：{output_path}")
except PermissionError:
    output_path_2 = output_path.replace(".xlsx", "-另存.xlsx")
    main_file_df_5.to_excel(output_path_2, index=False)
    print(f"处理完成（原文件被占用，已另存），文件为：{output_path_2}")

print(
    f"{Color.YELLOW}[注意]检查————供应商、运营模式、产品状态、三级分类，"
    f"是不是都有了，没有的部分手动去判断、填写进去！！！{Color.RESET}"
)
print("新品 二、三级分类，空着；分销 供应商目前都是：智慧谷！")
if folder_name == "月报":
    print(
        f"{Color.RED}~~~~~~~~~~~~~~~~~站点：AMAZON-UK，如果只有 '仓租'，"
        f"就把'仓租'放到AMAZON-DE，然后删掉 AMAZON-UK！{Color.RESET}"
    )
