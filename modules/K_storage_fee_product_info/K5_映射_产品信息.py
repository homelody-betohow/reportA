"""
K5_映射_产品信息 — 订单统计补全产品属性（16→17）

【流水线位置】
  上游：(已完成-16)订单统计（K1_4；月报时 MANO 仓租可能已由 C2 写入「FBA仓租费」）
  下游：(已完成-17) → K6 分销收尾 + 无平台仓租分摊

【映射来源】
  1. 运营模式 / 供应商 ← DB `product_sku`（ops_model / supplier_abbr）
  2. 二级分类 / 三级分类 / 产品状态 ← 产品信息库2025.xlsx（+ 桌面「信息-映射.xlsx」兜底）
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

PRODUCT_INFO_XLSX = r"\\Betohow\数据报表\数据库\产品信息库2025.xlsx"
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


def fetch_ops_and_supplier(sku_keys: list[str]) -> dict[str, tuple[str, str]]:
    """
    从 product_sku 批量查运营模式 / 供应商简称。
    返回 product_sku → (ops_model, supplier_abbr)；空串不当作成有效值。
    """
    sku_keys = sorted({str(x).strip() for x in sku_keys if x and str(x).strip()})
    if not sku_keys:
        return {}

    sql = f"""
        SELECT product_sku, ops_model, supplier_abbr
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

    out: dict[str, tuple[str, str]] = {}
    for row in rows:
        sku = str(row.get("product_sku") or "").strip()
        if not sku:
            continue
        ops = str(row.get("ops_model") or "").strip()
        supplier = str(row.get("supplier_abbr") or "").strip()
        out[sku] = (ops, supplier)
    return out


def apply_ops_and_supplier_from_db(df: pd.DataFrame) -> pd.DataFrame:
    """按 SKU 写入「运营模式」「供应商」（来自 product_sku）。"""
    if "SKU" not in df.columns:
        raise KeyError("订单统计缺少列「SKU」")

    out = df.copy()
    for col in ("运营模式", "供应商"):
        if col in out.columns:
            out = out.drop(columns=[col])

    keys = _norm_sku_key(out["SKU"])
    unique_keys = sorted({k for k in keys.tolist() if k})
    mapping = fetch_ops_and_supplier(unique_keys)

    ops_list: list[str | None] = []
    supplier_list: list[str | None] = []
    hit = 0
    for k in keys:
        if k and k in mapping:
            ops, supplier = mapping[k]
            ops_list.append(ops or None)
            supplier_list.append(supplier or None)
            if ops or supplier:
                hit += 1
        else:
            ops_list.append(None)
            supplier_list.append(None)

    sku_pos = out.columns.get_loc("SKU")
    out.insert(sku_pos + 1, "运营模式", ops_list)
    out.insert(sku_pos + 2, "供应商", supplier_list)
    print(
        f"[DB] product_sku 映射运营模式/供应商："
        f"唯一SKU={len(unique_keys)}，命中行={hit}/{len(out)}"
    )
    return out


def _replace_mapped_col(df: pd.DataFrame, mapped_col: str, target_col: str) -> pd.DataFrame:
    """把「映射X」落到目标列（先删已有同名列，避免重复列名）。"""
    out = df.copy()
    if target_col in out.columns:
        out = out.drop(columns=[target_col])
    return out.rename(columns={mapped_col: target_col})


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

main_file_path = (
    fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计"
    fr"\(已完成-16)订单统计-{shared_date}.xlsx"
)
main_file_df = pd.read_excel(main_file_path)

# 1. 运营模式、供应商 ← product_sku
main_file_df_2 = apply_ops_and_supplier_from_db(main_file_df)

# 2. 二级分类、三级分类、产品状态 ← 产品信息库2025.xlsx
product_map_sku_path = PRODUCT_INFO_XLSX

main_file_df_3 = sku_mappings(
    main_df=main_file_df_2,
    main_sku="SKU",
    map_sku_path=product_map_sku_path,
    map_old_sku="产品编码",
    map_new_sku="二级分类",
    map_sku_sheet="产品信息表",
)
main_file_df_3 = _replace_mapped_col(main_file_df_3, "映射二级分类", "二级分类")

main_file_df_4 = sku_mappings(
    main_df=main_file_df_3,
    main_sku="SKU",
    map_sku_path=product_map_sku_path,
    map_old_sku="产品编码",
    map_new_sku="三级分类",
    map_sku_sheet="产品信息表",
)
main_file_df_4 = _replace_mapped_col(main_file_df_4, "映射三级分类", "三级分类")

# 映射产品状态：平台是否包含 AMAZON
df_amazon = main_file_df_4[
    main_file_df_4["平台"].str.contains("AMAZON", case=False, na=False)
]
df_not_amazon = main_file_df_4[
    ~main_file_df_4["平台"].str.contains("AMAZON", case=False, na=False)
]

df_amazon_1 = sku_mappings(
    main_df=df_amazon,
    main_sku="SKU",
    map_sku_path=product_map_sku_path,
    map_old_sku="产品编码",
    map_new_sku="AMZ新老品",
    map_sku_sheet="产品信息表",
)
df_amazon_1 = _replace_mapped_col(df_amazon_1, "映射AMZ新老品", "产品状态")

df_not_amazon_1 = sku_mappings(
    main_df=df_not_amazon,
    main_sku="SKU",
    map_sku_path=product_map_sku_path,
    map_old_sku="产品编码",
    map_new_sku="本土平台新老品",
    map_sku_sheet="产品信息表",
)
df_not_amazon_1 = _replace_mapped_col(df_not_amazon_1, "映射本土平台新老品", "产品状态")

main_file_df_5 = pd.concat([df_amazon_1, df_not_amazon_1]).reset_index(drop=True)

# 3. 产品状态为空 → 桌面「信息-映射.xlsx」兜底
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
