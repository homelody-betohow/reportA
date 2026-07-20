"""
G2_1_Temu罚款映射.py — TEMU 罚款汇总映射仓库 SKU 并按发货数量分摊费用

功能：
  1. 读取 G2_1_Temu罚款合并.py 生成的 TEMU-罚款汇总-{shared_date}.xlsx
  2. 按 订单编号 关联 sales_order_shipped（platform=semitemu, ref_no=订单编号）
  3. 一单多品拆分为多行，新增 SKU、发货数量、店铺名称、SKU-站点识别
  4. 支出金额、结算金额按发货数量比例分摊
  5. 输出：(已完成)TEMU-罚款汇总-{shared_date}.xlsx
"""

import importlib.util
import sys
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import pymysql.cursors

# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.Z_method.style import Color
from A_报表.A0_设置_时间段.A0_set_date import folder_name, shared_date
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

_REPORT_PRA_ROOT = Path(__file__).resolve().parents[2] / "reportPRA"
if str(_REPORT_PRA_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPORT_PRA_ROOT))

from database.db_connection import get_db_manager  # noqa: E402  # pyright: ignore[reportMissingImports]

SHIPPED_TABLE = "sales_order_shipped"
PLATFORM_SEMITEMU = "semitemu"
_KEY_CHUNK = 200

ORDER_NO_COL = "订单编号"
AMOUNT_COL = "支出金额"
SETTLE_AMOUNT_COL = "结算金额"
SKU_COL = "SKU"
QTY_COL = "发货数量"
SHOP_COL = "店铺"
SHOP_NAME_COL = "店铺名称"
SKU_SITE_ID_COL = "SKU-站点识别"
EXPENSE_TYPE_COL = "支出类型"

INPUT_NAME = f"TEMU-罚款汇总-{shared_date}.xlsx"
OUTPUT_NAME = f"(已完成)TEMU-罚款汇总-{shared_date}.xlsx"
FEE_COLUMNS = (AMOUNT_COL, SETTLE_AMOUNT_COL)

TEMU_FINE_DIR = Path(fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\TEMU-罚款")


def _strip_df_strings(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
    return df


def _norm_ref_no(val) -> str:
    if pd.isna(val):
        return ""
    return str(val).strip()


def _to_qty(val) -> int:
    if pd.isna(val):
        return 0
    try:
        return max(int(float(val)), 0)
    except (TypeError, ValueError):
        return 0


def _build_sku_site_id(shop, sku) -> str | None:
    """SKU-站点识别 = 店铺 + SKU（与订单统计 SKU-站点识别码 一致）。"""
    if pd.isna(shop) or pd.isna(sku):
        return None
    shop_s = str(shop).strip()
    sku_s = str(sku).strip()
    if not shop_s or not sku_s:
        return None
    return shop_s + sku_s


def _chunked_in_query(cur, sql_template: str, keys: list[str], extra_params: tuple = ()) -> list[dict]:
    if not keys:
        return []
    results: list[dict] = []
    for i in range(0, len(keys), _KEY_CHUNK):
        chunk = keys[i : i + _KEY_CHUNK]
        placeholders = ",".join(["%s"] * len(chunk))
        cur.execute(sql_template.format(placeholders=placeholders), extra_params + tuple(chunk))
        results.extend(cur.fetchall())
    return results


def fetch_shipped_lines_by_ref_no(ref_nos: list[str]) -> dict[str, list[dict]]:
    """按 ref_no 批量查询 semitemu 发货明细，返回 ref_no -> 行列表。"""
    ref_nos = sorted({_norm_ref_no(x) for x in ref_nos if _norm_ref_no(x)})
    if not ref_nos:
        return {}

    sql = f"""
        SELECT ref_no, shop_name_en, warehouse_sku, warehouse_sku_qty
        FROM `{SHIPPED_TABLE}`
        WHERE platform = %s
          AND ref_no IN ({{placeholders}})
          AND warehouse_sku IS NOT NULL
          AND TRIM(warehouse_sku) <> ''
        ORDER BY id ASC
    """
    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        rows = _chunked_in_query(cur, sql, ref_nos, extra_params=(PLATFORM_SEMITEMU,))
    finally:
        cur.close()
        conn.close()

    grouped: dict[str, list[dict]] = {}
    for row in rows:
        ref_no = _norm_ref_no(row.get("ref_no"))
        warehouse_sku = str(row.get("warehouse_sku") or "").strip()
        if not ref_no or not warehouse_sku:
            continue
        grouped.setdefault(ref_no, []).append(
            {
                "shop_name_en": str(row.get("shop_name_en") or "").strip(),
                "warehouse_sku": warehouse_sku,
                "warehouse_sku_qty": _to_qty(row.get("warehouse_sku_qty")),
            }
        )
    return grouped


def _allocate_fee_values(total: float, ratios: list[float]) -> list[float]:
    """按 ratio 分摊金额，末行补差，保证合计等于原值。"""
    if not ratios:
        return []
    if total is None or (isinstance(total, float) and np.isnan(total)):
        return [None] * len(ratios)

    total_f = float(total)
    if len(ratios) == 1:
        return [float(np.round(total_f, 2))]

    parts: list[float] = []
    allocated = 0.0
    for ratio in ratios[:-1]:
        part = float(np.round(total_f * ratio, 2))
        parts.append(part)
        allocated += part
    parts.append(float(np.round(total_f - allocated, 2)))
    return parts


def _split_row_by_shipped(fine_row: pd.Series, shipped_lines: list[dict]) -> list[dict]:
    qty_list = [_to_qty(line.get("warehouse_sku_qty")) for line in shipped_lines]
    total_qty = sum(qty_list)
    if total_qty > 0:
        ratios = [q / total_qty for q in qty_list]
    else:
        ratios = [1 / len(shipped_lines)] * len(shipped_lines)

    fee_alloc: dict[str, list[float]] = {}
    for fee_col in FEE_COLUMNS:
        if fee_col in fine_row.index:
            fee_alloc[fee_col] = _allocate_fee_values(fine_row[fee_col], ratios)

    out_rows: list[dict] = []
    base = fine_row.to_dict()
    for idx, line in enumerate(shipped_lines):
        row = base.copy()
        row[SKU_COL] = line["warehouse_sku"]
        row[QTY_COL] = line.get("warehouse_sku_qty")
        row[SHOP_NAME_COL] = line.get("shop_name_en") or None
        row[SKU_SITE_ID_COL] = _build_sku_site_id(fine_row.get(SHOP_COL), line["warehouse_sku"])
        for fee_col, values in fee_alloc.items():
            row[fee_col] = values[idx]
        out_rows.append(row)
    return out_rows


def _map_fines_to_skus(fine_df: pd.DataFrame, shipped_map: dict[str, list[dict]]) -> tuple[pd.DataFrame, list[str]]:
    if ORDER_NO_COL not in fine_df.columns:
        raise KeyError(f"罚款汇总缺少列 {ORDER_NO_COL!r}")

    mapped_rows: list[dict] = []
    unmatched_refs: list[str] = []

    for _, row in fine_df.iterrows():
        ref_no = _norm_ref_no(row.get(ORDER_NO_COL))
        shipped_lines = shipped_map.get(ref_no, [])
        if not shipped_lines:
            unmatched_refs.append(ref_no or "(空订单编号)")
            new_row = row.to_dict()
            new_row[SKU_COL] = None
            new_row[QTY_COL] = None
            new_row[SHOP_NAME_COL] = None
            new_row[SKU_SITE_ID_COL] = None
            mapped_rows.append(new_row)
            continue
        mapped_rows.extend(_split_row_by_shipped(row, shipped_lines))

    result_df = pd.DataFrame(mapped_rows)
    return result_df, unmatched_refs


def _order_output_columns(df: pd.DataFrame) -> pd.DataFrame:
    front_cols = [c for c in (SHOP_COL, SHOP_NAME_COL, EXPENSE_TYPE_COL, ORDER_NO_COL) if c in df.columns]
    sku_cols = [c for c in (SKU_COL, SKU_SITE_ID_COL, QTY_COL) if c in df.columns]
    middle_cols = [c for c in df.columns if c not in front_cols + sku_cols]

    if ORDER_NO_COL in middle_cols:
        pos = middle_cols.index(ORDER_NO_COL) + 1
        middle_cols = middle_cols[:pos] + sku_cols + middle_cols[pos:]
    else:
        middle_cols = sku_cols + middle_cols

    return df[front_cols + middle_cols]


def _resolve_input_path() -> Path:
    candidates = [
        TEMU_FINE_DIR / INPUT_NAME,
        TEMU_FINE_DIR / f"(处理完成){INPUT_NAME}",
    ]
    for path in candidates:
        if path.is_file():
            return path
    raise FileNotFoundError(
        f"未找到罚款汇总文件，请先运行 G2_1_Temu罚款合并.py。尝试路径：{candidates}"
    )


def _save_excel(df: pd.DataFrame, output_path: Path) -> Path:
    try:
        df.to_excel(output_path, index=False, engine="openpyxl")
        return output_path
    except PermissionError:
        alt = output_path.with_name(output_path.stem + "-另存.xlsx")
        df.to_excel(alt, index=False, engine="openpyxl")
        print(f"{Color.YELLOW}[提示]{Color.RESET} 目标文件被占用，已另存为：{alt}")
        return alt


def main() -> None:
    input_path = _resolve_input_path()
    output_path = TEMU_FINE_DIR / OUTPUT_NAME

    fine_df = _strip_df_strings(pd.read_excel(input_path))
    print(f"{Color.CYAN}[读取]{Color.RESET} {input_path.name}：{len(fine_df)} 行")

    ref_nos = fine_df[ORDER_NO_COL].map(_norm_ref_no).tolist()
    shipped_map = fetch_shipped_lines_by_ref_no(ref_nos)
    hit_cnt = sum(1 for ref in {_norm_ref_no(x) for x in ref_nos if _norm_ref_no(x)} if ref in shipped_map)
    print(
        f"{Color.GREEN}[DB]{Color.RESET} sales_order_shipped 命中 "
        f"{hit_cnt} / {len({_norm_ref_no(x) for x in ref_nos if _norm_ref_no(x)})} 个订单编号"
    )

    mapped_df, unmatched_refs = _map_fines_to_skus(fine_df, shipped_map)
    mapped_df = _order_output_columns(mapped_df)

    saved = _save_excel(mapped_df, output_path)
    print(
        f"{Color.GREEN}[完成]{Color.RESET} {len(fine_df)} 行 → {len(mapped_df)} 行（按 SKU 拆分后）"
    )
    if unmatched_refs:
        unique_unmatched = sorted(set(unmatched_refs))
        print(
            f"{Color.YELLOW}[检查]{Color.RESET} {len(unique_unmatched)} 个订单编号未在 "
            f"sales_order_shipped(semitemu) 中命中："
        )
        preview = ", ".join(unique_unmatched[:10])
        suffix = " ..." if len(unique_unmatched) > 10 else ""
        print(f"  {preview}{suffix}")

    print(f"处理完成，output_path：{saved}")


if __name__ == "__main__":
    main()
