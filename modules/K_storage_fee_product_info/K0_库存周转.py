"""
K0_库存周转.py — 从 snapshot_market_turnover 反向生成库存动销明细 Excel

输出：
  {DESKTOP_ROOT}\\{folder_name}{shared_date}\\仓租\\{yyyy.m.d}库存动销明细.xlsx
  sheet：
    1. 各平台SKU库存动销明细
    2. 各平台商品库存周转明细（由 sheet1 按 平台+商品ID 汇总，货值除外）
    3. 商品库存周转明细（由 sheet1 按 商品ID 汇总，货值除外；不含平台列）
    4. 库存周转汇总透视（三块：按平台 / 产品状态 / 销售负责人；含货值×数量）

供 K1/K2 仓租分摊读取（需含列：商品ID、平台、SKU、在库（可调拨））。

用法：
  python modules/K_storage_fee_product_info/K0_库存周转.py
  python modules/K_storage_fee_product_info/K0_库存周转.py --date 2026.7.18
  python modules/K_storage_fee_product_info/K0_库存周转.py -d 2026-07-18
"""

from __future__ import annotations

import argparse
import importlib.util
import re
import warnings
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import pymysql.cursors
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from config.A0_paths import DESKTOP_ROOT  # noqa: E402
from config.A0_set_date import folder_name, ku_cun_date, report_date, shared_date  # noqa: E402
from database.db_connection import get_db_manager  # noqa: E402

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")

SHEET_SKU = "各平台SKU库存周转明细"
SHEET_PRODUCT_BY_MARKET = "各平台商品ID周转明细"
SHEET_PRODUCT = "商品ID周转明细"
SHEET_PIVOT = "库存周转汇总"
FILE_SUFFIX = "库存动销明细.xlsx"

# 商品级汇总时不求和的列（货值按组取首行；周转天数汇总后重算）
_PRODUCT_SKIP_SUM = frozenset({"货值", "海外周转-月", "总库存周转-月"})
# 商品级汇总时取首行的文本/标识列（非分组键时生效）
_PRODUCT_FIRST_COLS = (
    "平台",
    "商品ID",
    "SKU",
    "产品状态",
    "运营经理",
    "运营负责人",
    "货值",
    "销售站点",
    "供应商",
)

# 库存/销量：整数
_INT_COLS = (
    "计划库存-禁调",
    "在途库存-禁调",
    "可售库存-禁调",
    "计划库存-可调",
    "在途库存-可调",
    "可售库存-可调",
    "总计划库存",
    "总在途库存",
    "总可售库存",
    "参考月销量",
    "海外库存",
    "总库存",
    "在库（可调拨）",
)
# 货值、周转天数：浮点
_FLOAT_COLS = (
    "货值",
    "海外周转-月",
    "总库存周转-月",
)
# 禁调列：灰白背景
_DIR_COLS = (
    "计划库存-禁调",
    "在途库存-禁调",
    "可售库存-禁调",
)
_DIR_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
_HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
_PIVOT_HEADER_FILL = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
_PIVOT_SALES_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
_HEADER_FONT = Font(bold=True, color="FFFFFF", name="微软雅黑", size=10)
_PIVOT_HEADER_FONT = Font(bold=True, color="000000", name="微软雅黑", size=10)
_BODY_FONT = Font(name="微软雅黑", size=10)
_TOTAL_FONT = Font(bold=True, name="微软雅黑", size=10)
_HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
_BODY_ALIGN = Alignment(horizontal="center", vertical="center")
_THIN_BORDER = Border(
    left=Side(style="thin", color="B4B4B4"),
    right=Side(style="thin", color="B4B4B4"),
    top=Side(style="thin", color="B4B4B4"),
    bottom=Side(style="thin", color="B4B4B4"),
)
_COL_WIDTH_MIN = 8
_COL_WIDTH_MAX = 28
_COL_WIDTH_PAD = 2

# 透视表：沿用 sheet1 原列名求和；另计货值金额（单价货值 × 数量）
_PIVOT_QTY_COLS = (
    "参考月销量",
    "计划库存-禁调",
    "在途库存-禁调",
    "可售库存-禁调",
    "计划库存-可调",
    "在途库存-可调",
    "可售库存-可调",
)
_PIVOT_AMOUNT_COLS = (
    "货值(在途+在库)",  # sum(货值 × 海外库存)
    "货值(整体库存)",  # sum(货值 × 总库存)
)
_PIVOT_VALUE_COLS = _PIVOT_QTY_COLS + _PIVOT_AMOUNT_COLS
_PIVOT_BLANK_LABELS = {
    "产品状态": "空白",
    "运营负责人": "无负责人",
}
_PIVOT_SECTIONS = (
    ("平台", "平台"),
    ("产品状态", "产品状态"),
    ("销售负责人", "运营负责人"),  # 表头用销售负责人，数据取运营负责人
)

# SQL 别名 → 导出列名
_RENAME_COLS = {
    "计划库存_禁调": "计划库存-禁调",
    "在途库存_禁调": "在途库存-禁调",
    "可售库存_禁调": "可售库存-禁调",
    "计划库存_可调拨": "计划库存-可调",
    "在途库存_可调": "在途库存-可调",
    "可售库存_可调": "可售库存-可调",
}

MARKET_TURNOVER_SQL = """
SELECT
    t.market_code          AS `销售市场`,
    t.product_uid          AS `商品ID`,
    t.product_sku          AS `SKU`,
    t.sku_lifecycle        AS `产品状态`,
    t.ops_leader           AS `运营经理`,
    t.ops_owner            AS `运营负责人`,
    t.supplier_name        AS `供应商`,
    t.cost_price_cny       AS `货值`,

    t.dir_planned_qty      AS `计划库存-禁调`,
    t.dir_onway_qty        AS `在途库存-禁调`,
    t.dir_sellable_qty     AS `可售库存-禁调`,

    t.trf_planned_qty      AS `计划库存-可调`,
    t.trf_onway_qty        AS `在途库存-可调`,
    t.trf_sellable_qty     AS `可售库存-可调`,

    t.total_planned_qty    AS `总计划库存`,
    t.total_onway_qty      AS `总在途库存`,
    t.total_sellable_qty   AS `总可售库存`,
    t.ref_month_sales_qty  AS `参考月销量`,
    t.market_region        AS `销售站点`,
    (t.total_onway_qty + t.total_sellable_qty) AS `海外库存`,
    (t.total_onway_qty + t.total_sellable_qty) / t.ref_month_sales_qty AS `海外周转-月`,
    (t.total_onway_qty + t.total_sellable_qty + t.total_planned_qty) AS `总库存`,
    (t.total_onway_qty + t.total_sellable_qty + t.total_planned_qty) / t.ref_month_sales_qty AS `总库存周转-月`
FROM
    snapshot_market_turnover t
WHERE
    t.snapshot_date = %s
ORDER BY
    t.market_code, t.product_sku
"""


def _format_ku_cun_date(d: date) -> str:
    """与 A0_set_date.ku_cun_date 一致：yyyy.m.d（月/日不补零）。"""
    return f"{d.year}.{d.month}.{d.day}"


def parse_snapshot_date(raw: str) -> date:
    """
    解析快照日，支持：
      2026.7.18 / 2026.07.18
      2026-7-18 / 2026-07-18
      2026/7/18 / 20260718
    """
    s = str(raw).strip()
    if not s:
        raise ValueError("日期不能为空")

    if re.fullmatch(r"\d{8}", s):
        return datetime.strptime(s, "%Y%m%d").date()

    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    # 兼容不补零：2026.7.18 / 2026-7-8
    m = re.fullmatch(r"(\d{4})[./-](\d{1,2})[./-](\d{1,2})", s)
    if m:
        y, mo, d = map(int, m.groups())
        return date(y, mo, d)

    raise ValueError(
        f"无法解析日期：{raw!r}，请使用如 2026.7.18 / 2026-07-18 / 20260718"
    )


def _fetch_market_turnover(snapshot_date: date) -> pd.DataFrame:
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            cur.execute(MARKET_TURNOVER_SQL, (snapshot_date,))
            rows = cur.fetchall()
    finally:
        conn.close()

    if not rows:
        return pd.DataFrame(columns=list(_export_column_order()))

    df = pd.DataFrame(rows)
    df = _attach_legacy_columns(df)
    return _coerce_numeric_columns(df)


def _attach_legacy_columns(df: pd.DataFrame) -> pd.DataFrame:
    """对齐导出列名，并补充 K1/K2 口径：平台、在库（可调拨）。"""
    out = df.rename(columns=_RENAME_COLS).copy()
    if "平台" not in out.columns and "销售市场" in out.columns:
        out["平台"] = out["销售市场"]
    if "在库（可调拨）" not in out.columns and "可售库存-可调" in out.columns:
        out["在库（可调拨）"] = out["可售库存-可调"]
    return out


def _display_width(value) -> int:
    """估算单元格显示宽度（中文约 2 个字符宽）。"""
    s = "" if value is None else str(value)
    width = 0
    for ch in s:
        width += 2 if ord(ch) > 127 else 1
    return width


def _autosize_columns(ws) -> None:
    for col_cells in ws.columns:
        letter = get_column_letter(col_cells[0].column)
        max_w = 0
        for cell in col_cells:
            max_w = max(max_w, _display_width(cell.value))
        ws.column_dimensions[letter].width = min(
            _COL_WIDTH_MAX, max(_COL_WIDTH_MIN, max_w + _COL_WIDTH_PAD)
        )


def _apply_sheet_styles(ws, columns: list[str]) -> None:
    """表头样式、禁调列灰底、边框、居中、冻结首行、自动列宽。"""
    dir_idxs = {i for i, name in enumerate(columns, start=1) if name in _DIR_COLS}
    float_idxs = {i for i, name in enumerate(columns, start=1) if name in _FLOAT_COLS}
    int_idxs = {i for i, name in enumerate(columns, start=1) if name in _INT_COLS}

    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = _THIN_BORDER
            cell.font = _HEADER_FONT if cell.row == 1 else _BODY_FONT
            cell.alignment = _HEADER_ALIGN if cell.row == 1 else _BODY_ALIGN

            if cell.row == 1:
                cell.fill = _HEADER_FILL
            elif cell.column in dir_idxs:
                cell.fill = _DIR_FILL

            if cell.row > 1 and cell.column in float_idxs and cell.value is not None:
                cell.number_format = "0.00"
            elif cell.row > 1 and cell.column in int_idxs and cell.value is not None:
                cell.number_format = "0"

    _autosize_columns(ws)


def _coerce_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    """确保库存/天数/货值以数字类型写入 Excel（避免 Decimal/object 被当成文本）。"""
    out = df.copy()
    for col in _INT_COLS:
        if col not in out.columns:
            continue
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0).astype("int64")
    for col in _FLOAT_COLS:
        if col not in out.columns:
            continue
        out[col] = pd.to_numeric(out[col], errors="coerce").astype("float64")
    return out


def _export_column_order() -> list[str]:
    return [
        "平台",
        "商品ID",
        "SKU",
        "产品状态",
        "运营经理",
        "运营负责人",
        "货值",
        "计划库存-禁调",
        "在途库存-禁调",
        "可售库存-禁调",
        "计划库存-可调",
        "在途库存-可调",
        "可售库存-可调",
        "总计划库存",
        "总在途库存",
        "总可售库存",
        "参考月销量",
        "销售站点",
        "海外库存",
        "海外周转-月",
        "总库存",
        "总库存周转-月",
        "供应商",
    ]


def _fill_blank_product_id(df: pd.DataFrame) -> pd.DataFrame:
    """商品ID 为空时默认等于 SKU。"""
    out = df.copy()
    if "商品ID" not in out.columns or "SKU" not in out.columns:
        return out
    blank = out["商品ID"].isna() | (
        out["商品ID"].astype(str).str.strip().isin(("", "nan", "None", "NaN"))
    )
    out.loc[blank, "商品ID"] = out.loc[blank, "SKU"].astype(str)
    return out


def _build_product_turnover(
    sku_df: pd.DataFrame,
    *,
    group_keys: list[str],
    drop_cols: tuple[str, ...] = (),
) -> pd.DataFrame:
    """
    由各平台SKU库存动销明细汇总商品级周转明细：
    商品ID 为空 → SKU；按 group_keys 分组，数值列求和（货值除外），周转天数重算。
    """
    drop_set = frozenset(drop_cols)
    cols = [
        c
        for c in _export_column_order()
        if c in sku_df.columns and c not in drop_set
    ]
    if sku_df.empty:
        return pd.DataFrame(columns=cols)

    src = _fill_blank_product_id(sku_df)
    for key in group_keys:
        if key not in src.columns:
            raise KeyError(f"商品级汇总缺少分组列：{key}")

    agg: dict[str, str] = {}
    for col in _PRODUCT_FIRST_COLS:
        if (
            col in src.columns
            and col not in group_keys
            and col not in drop_set
        ):
            agg[col] = "first"

    sum_candidates = (
        list(_INT_COLS) + [c for c in _FLOAT_COLS if c not in _PRODUCT_SKIP_SUM]
    )
    for col in sum_candidates:
        if col in src.columns and col not in group_keys and col not in agg:
            agg[col] = "sum"

    grouped = src.groupby(group_keys, as_index=False, dropna=False).agg(agg)

    # 海外库存 / 总库存：用汇总后分量重算，与分量保持一致
    if {"总在途库存", "总可售库存"}.issubset(grouped.columns):
        grouped["海外库存"] = grouped["总在途库存"] + grouped["总可售库存"]
    if {"总在途库存", "总可售库存", "总计划库存"}.issubset(grouped.columns):
        grouped["总库存"] = (
            grouped["总在途库存"] + grouped["总可售库存"] + grouped["总计划库存"]
        )

    # 周转天数：汇总后再除，避免对比率求和；销量为 0 时结果为 NaN
    sales = grouped["参考月销量"] if "参考月销量" in grouped.columns else None
    if sales is not None:
        sales_safe = sales.mask(sales.eq(0))
        if "海外库存" in grouped.columns:
            grouped["海外周转-月"] = grouped["海外库存"] / sales_safe
        if "总库存" in grouped.columns:
            grouped["总库存周转-月"] = grouped["总库存"] / sales_safe

    grouped = _coerce_numeric_columns(grouped)
    out_cols = [
        c
        for c in _export_column_order()
        if c in grouped.columns and c not in drop_set
    ]
    return grouped[out_cols]


def _normalize_pivot_dim(series: pd.Series, *, blank_label: str) -> pd.Series:
    """空维度值替换为空白/无负责人等标签。"""
    s = series.astype(object).where(series.notna(), "")
    s = s.map(lambda v: str(v).strip())
    blank = s.isin(("", "nan", "None", "NaN"))
    return s.mask(blank, blank_label)


def _prepare_pivot_source(sku_df: pd.DataFrame) -> pd.DataFrame:
    """透视前准备：数量列转数值，并预计算行级货值金额。"""
    src = sku_df.copy()
    for col in _PIVOT_QTY_COLS:
        if col not in src.columns:
            src[col] = 0
        src[col] = pd.to_numeric(src[col], errors="coerce").fillna(0)

    unit_cost = pd.to_numeric(src["货值"], errors="coerce").fillna(0) if "货值" in src.columns else 0

    if "海外库存" in src.columns:
        overseas = pd.to_numeric(src["海外库存"], errors="coerce").fillna(0)
    else:
        overseas = (
            pd.to_numeric(src.get("总在途库存", 0), errors="coerce").fillna(0)
            + pd.to_numeric(src.get("总可售库存", 0), errors="coerce").fillna(0)
        )

    if "总库存" in src.columns:
        total_inv = pd.to_numeric(src["总库存"], errors="coerce").fillna(0)
    else:
        total_inv = overseas + pd.to_numeric(
            src.get("总计划库存", 0), errors="coerce"
        ).fillna(0)

    src["货值(在途+在库)"] = unit_cost * overseas
    src["货值(整体库存)"] = unit_cost * total_inv
    return src


def _build_one_pivot_section(
    sku_df: pd.DataFrame,
    *,
    source_col: str,
    dim_label: str,
) -> pd.DataFrame:
    """按单一维度汇总一块透视表（含总计行）。"""
    headers = [dim_label, *_PIVOT_VALUE_COLS]
    if sku_df.empty or source_col not in sku_df.columns:
        return pd.DataFrame(columns=headers)

    src = _prepare_pivot_source(sku_df)
    blank_label = _PIVOT_BLANK_LABELS.get(source_col, "空白")
    src["_dim"] = _normalize_pivot_dim(src[source_col], blank_label=blank_label)

    sum_cols = {col: "sum" for col in _PIVOT_VALUE_COLS}
    grouped = src.groupby("_dim", as_index=False, dropna=False).agg(sum_cols)
    grouped = grouped.rename(columns={"_dim": dim_label})
    grouped = grouped.sort_values(by=dim_label, kind="mergesort").reset_index(drop=True)

    total = {dim_label: "总计"}
    for col in _PIVOT_VALUE_COLS:
        total[col] = pd.to_numeric(grouped[col], errors="coerce").fillna(0).sum()
    out = pd.concat([grouped, pd.DataFrame([total])], ignore_index=True)
    return out[headers]


def _build_pivot_sections(sku_df: pd.DataFrame) -> list[pd.DataFrame]:
    """生成三块透视：平台 / 产品状态 / 销售负责人。"""
    return [
        _build_one_pivot_section(sku_df, source_col=src, dim_label=label)
        for label, src in _PIVOT_SECTIONS
    ]


def _write_pivot_sheet(ws, sections: list[pd.DataFrame]) -> None:
    """把三块透视纵向写入同一 sheet，块之间空一行。"""
    row_cursor = 1
    max_col = 1 + len(_PIVOT_VALUE_COLS)

    for section_idx, section_df in enumerate(sections):
        headers = list(section_df.columns)
        for col_idx, name in enumerate(headers, start=1):
            cell = ws.cell(row=row_cursor, column=col_idx, value=name)
            cell.font = _PIVOT_HEADER_FONT
            cell.alignment = _HEADER_ALIGN
            cell.border = _THIN_BORDER
            cell.fill = (
                _PIVOT_SALES_FILL if name == "参考月销量" else _PIVOT_HEADER_FILL
            )
        ws.row_dimensions[row_cursor].height = 30
        header_row = row_cursor
        row_cursor += 1

        for _, rec in section_df.iterrows():
            is_total = str(rec.iloc[0]) == "总计"
            for col_idx, name in enumerate(headers, start=1):
                val = rec[name]
                if pd.isna(val):
                    val = None
                elif name in _PIVOT_QTY_COLS and val is not None:
                    val = int(round(float(val)))
                elif name in _PIVOT_AMOUNT_COLS and val is not None:
                    val = float(val)
                cell = ws.cell(row=row_cursor, column=col_idx, value=val)
                cell.font = _TOTAL_FONT if is_total else _BODY_FONT
                cell.alignment = _BODY_ALIGN
                cell.border = _THIN_BORDER
                if name in _PIVOT_AMOUNT_COLS and val is not None:
                    cell.number_format = "0.00"
                elif name in _PIVOT_QTY_COLS and val is not None:
                    cell.number_format = "0"
            row_cursor += 1

        if section_idx < len(sections) - 1:
            row_cursor += 1

        if section_idx == 0:
            ws.freeze_panes = f"A{header_row + 1}"

    ws.column_dimensions["A"].width = 16
    for col_idx in range(2, max_col + 1):
        letter = get_column_letter(col_idx)
        max_w = _COL_WIDTH_MIN
        for row in ws.iter_rows(
            min_row=1, max_row=ws.max_row, min_col=col_idx, max_col=col_idx
        ):
            for cell in row:
                max_w = max(max_w, _display_width(cell.value))
        ws.column_dimensions[letter].width = min(
            _COL_WIDTH_MAX, max(12, max_w + _COL_WIDTH_PAD)
        )


def export_turnover(snapshot_date: date, date_label: str) -> Path:
    output_dir = Path(fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租")
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{date_label}{FILE_SUFFIX}"

    df = _fetch_market_turnover(snapshot_date)
    cols = [c for c in _export_column_order() if c in df.columns]
    if not df.empty:
        df = df[cols]

    by_market_df = _build_product_turnover(df, group_keys=["平台", "商品ID"])
    by_product_df = _build_product_turnover(
        df, group_keys=["商品ID"], drop_cols=("平台",)
    )
    pivot_sections = _build_pivot_sections(df)

    detail_sheets = (
        (SHEET_SKU, df, cols),
        (SHEET_PRODUCT_BY_MARKET, by_market_df, list(by_market_df.columns)),
        (SHEET_PRODUCT, by_product_df, list(by_product_df.columns)),
    )
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, sheet_df, sheet_cols in detail_sheets:
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
            _apply_sheet_styles(writer.sheets[sheet_name], sheet_cols)

        # 第 4 个 sheet：三块纵向透视
        ws = writer.book.create_sheet(SHEET_PIVOT)
        _write_pivot_sheet(ws, pivot_sections)

    pivot_rows = sum(max(0, len(s) - 1) for s in pivot_sections)  # 不含总计
    print(
        f"库存动销明细已导出：SKU {len(df)} 行 / "
        f"各平台商品 {len(by_market_df)} 行 / 商品 {len(by_product_df)} 行 / "
        f"透视维度行 {pivot_rows}，"
        f"快照日 {snapshot_date}，文件另存为：{output_path}"
    )
    return output_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="从 snapshot_market_turnover 导出 SKU/商品/透视 库存动销与周转明细"
    )
    parser.add_argument(
        "-d",
        "--date",
        dest="date",
        default=None,
        help="快照日（默认取 A0_set_date.report_date）。例：2026.7.18 / 2026-07-18",
    )
    args = parser.parse_args(argv)

    if args.date:
        snapshot_date = parse_snapshot_date(args.date)
        date_label = _format_ku_cun_date(snapshot_date)
    else:
        snapshot_date = report_date.date()
        date_label = ku_cun_date

    export_turnover(snapshot_date, date_label)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
