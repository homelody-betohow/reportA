"""将 ``product_sku`` 同步到钉钉「产品信息」表格（DB → 钉钉）。

支持两种模式：

1. **更新列**（默认）：表格已有行按 ``SKU`` 对齐，回写白名单有差异的单元格。
2. **追加**：表格中尚不存在的 ``SKU`` 按表头追加整行（供易仓全量同步新增使用）。

匹配键::
    product_sku  →  SKU

文档 ID（workbookId / nodeId）::
    Obva6QBXJwjBxoE2sM62MrzGVn4qY5Pr

用法（项目根目录）::

    # 默认：用 DB 更新白名单全部字段
    python app/ding-disk/product_sku_push.py
    python app/ding-disk/product_sku_push.py --dry-run

    # 只更新部分字段 / 指定 SKU
    python app/ding-disk/product_sku_push.py --up-field supplier_abbr
    python app/ding-disk/product_sku_push.py --sku XPM0213
    python app/ding-disk/product_sku_push.py --sku XPM0213 --up-field hs_code --dry-run

    # 追加库中有、表中无的 SKU
    python app/ding-disk/product_sku_push.py --append
    python app/ding-disk/product_sku_push.py --append --sku NEW001 --dry-run

    # 追加 + 更新
    python app/ding-disk/product_sku_push.py --append --update

    python app/ding-disk/product_sku_push.py --list-sheets
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple

import pandas as pd
import pymysql.cursors

_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
assert _spec.loader is not None
_spec.loader.exec_module(_epr_mod)
_PROJECT_ROOT = _epr_mod.bootstrap(__file__)

from database.db_connection import get_db_manager  # noqa: E402
from api.ding_disk.exceptions import DingDiskError  # noqa: E402
from api.ding_disk.workbook import (  # noqa: E402
    Workbook,
    clean_cell,
    filter_by_column,
)

# 与 productInfo.py 保持一致
WORKBOOK_ID = "Obva6QBXJwjBxoE2sM62MrzGVn4qY5Pr"
DEFAULT_SHEET = "Sheet1"
TABLE = "product_sku"
SKU_SHEET_COL = "SKU"
FIX_DB_COL = "product_sku"
BATCH_SIZE = 500


@dataclass(frozen=True)
class FieldMap:
    """字段名 → 钉钉列 → product_sku 列。"""

    sheet_col: str
    db_col: str


# 白名单：DB → 钉钉（更新 / 追加共用；列名与钉钉表头一致，含「图片URl」笔误）
UP_FIELD_MAP: Mapping[str, FieldMap] = {
    "product_uid": FieldMap(sheet_col="商品ID", db_col="product_uid"),
    "product_name_en": FieldMap(sheet_col="英文标题", db_col="product_name_en"),
    "warehouse_ref": FieldMap(sheet_col="仓库识别码", db_col="warehouse_ref"),
    "category_lv1": FieldMap(sheet_col="一级分类", db_col="category_lv1"),
    "category_lv2": FieldMap(sheet_col="二级分类", db_col="category_lv2"),
    "category_lv3": FieldMap(sheet_col="三级分类", db_col="category_lv3"),
    "category_code": FieldMap(sheet_col="品类编码", db_col="category_code"),
    "ean_code": FieldMap(sheet_col="EAN码", db_col="ean_code"),
    "supplier_abbr": FieldMap(sheet_col="供应商", db_col="supplier_abbr"),
    "supplier_name": FieldMap(sheet_col="供应商全称", db_col="supplier_name"),
    "product_color": FieldMap(sheet_col="颜色", db_col="product_color"),
    "product_img": FieldMap(sheet_col="图片URl", db_col="product_img"),
    "declare_price": FieldMap(sheet_col="申报价值USD", db_col="declare_price_usd"),
    "declare_name_cn": FieldMap(sheet_col="申报中文", db_col="declare_name_cn"),
    "declare_name_en": FieldMap(sheet_col="申报英文", db_col="declare_name_en"),
    "hs_code": FieldMap(sheet_col="海关编码", db_col="hs_code"),
    "amz_lifecycle": FieldMap(sheet_col="AMZ状态", db_col="amz_lifecycle"),
    "local_lifecycle": FieldMap(sheet_col="本土状态", db_col="local_lifecycle"),
    "accounting_class": FieldMap(sheet_col="核算分类", db_col="accounting_class"),
    "carton_qty": FieldMap(sheet_col="内箱数量", db_col="carton_qty"),
    "product_unit": FieldMap(sheet_col="单位", db_col="product_unit"),
    "purchase_moq": FieldMap(sheet_col="最小起定量", db_col="purchase_moq"),
    "purchase_lead_days": FieldMap(sheet_col="采购交期", db_col="purchase_lead_days"),
    "purchase_price": FieldMap(sheet_col="采购价CNY", db_col="purchase_price"),
    "cost_price_cny": FieldMap(sheet_col="成本价CNY", db_col="cost_price_cny"),
    "unit_weight_g": FieldMap(sheet_col="单件净重(g)", db_col="unit_weight_g"),
    "carton_gross_g": FieldMap(sheet_col="外箱毛重(g)", db_col="carton_gross_g"),
    "inner_box_l_cm": FieldMap(sheet_col="内箱长(cm)", db_col="inner_box_l_cm"),
    "inner_box_w_cm": FieldMap(sheet_col="内箱宽(cm)", db_col="inner_box_w_cm"),
    "inner_box_h_cm": FieldMap(sheet_col="内箱高(cm)", db_col="inner_box_h_cm"),
    "outer_box_l_cm": FieldMap(sheet_col="外箱长(cm)", db_col="outer_box_l_cm"),
    "outer_box_w_cm": FieldMap(sheet_col="外箱宽(cm)", db_col="outer_box_w_cm"),
    "outer_box_h_cm": FieldMap(sheet_col="外箱高(cm)", db_col="outer_box_h_cm"),
    "first_leg_eu_au_cny": FieldMap(sheet_col="EU/AU头程运费", db_col="first_leg_eu_au_cny"),
    "first_leg_us_cny": FieldMap(sheet_col="US头程运费", db_col="first_leg_us_cny"),
    "first_leg_uk_cny": FieldMap(sheet_col="UK头程运费", db_col="first_leg_uk_cny"),
    "duty_eu_cny": FieldMap(sheet_col="EU单件关税", db_col="duty_eu_cny"),
    "duty_us_cny": FieldMap(sheet_col="US单件关税", db_col="duty_us_cny"),
    "duty_uk_cny": FieldMap(sheet_col="UK单件关税", db_col="duty_uk_cny"),
    "ops_model": FieldMap(sheet_col="运营模式", db_col="ops_model"),
    "ops_tax_rate": FieldMap(sheet_col="运营税率", db_col="ops_tax_rate"),
}


@dataclass
class UpdateStats:
    sheet_rows: int = 0
    unique_skus: int = 0
    empty_value: int = 0
    missing_in_db: int = 0
    unchanged: int = 0
    updated: int = 0
    skipped_missing_col: bool = False


@dataclass
class AppendStats:
    db_rows: int = 0
    sheet_skus: int = 0
    already_in_sheet: int = 0
    to_append: int = 0
    appended: int = 0
    missing_sku_col: bool = False


def resolve_up_fields(keys: Optional[Sequence[str]]) -> List[str]:
    """解析待更新字段；未指定则返回白名单全部。"""
    if not keys:
        return list(UP_FIELD_MAP.keys())
    unknown = [k for k in keys if k not in UP_FIELD_MAP]
    if unknown:
        known = ", ".join(sorted(UP_FIELD_MAP))
        raise ValueError(f"未知字段 {unknown}；可选: {known}")
    wanted = set(keys)
    return [k for k in UP_FIELD_MAP if k in wanted]


def _sheet_to_db_col() -> Dict[str, str]:
    """钉钉列名 → DB 列名（含匹配键）。"""
    mapping = {SKU_SHEET_COL: FIX_DB_COL}
    for field in UP_FIELD_MAP.values():
        mapping[field.sheet_col] = field.db_col
    return mapping


def fetch_product_map(
    skus: Sequence[str],
    db_cols: Sequence[str],
) -> Dict[str, Dict[str, str]]:
    """批量读取 ``product_sku`` → ``{sku: {db_col: value}}``。"""
    if not skus or not db_cols:
        return {}
    col_sql = ", ".join(f"`{c}`" for c in db_cols)
    result: Dict[str, Dict[str, str]] = {}
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            for i in range(0, len(skus), BATCH_SIZE):
                chunk = list(skus[i : i + BATCH_SIZE])
                placeholders = ",".join(["%s"] * len(chunk))
                cur.execute(
                    f"SELECT product_sku, {col_sql} FROM `{TABLE}` "
                    f"WHERE product_sku IN ({placeholders}) AND is_deleted = 0",
                    chunk,
                )
                for row in cur.fetchall():
                    sku = clean_cell(row["product_sku"])
                    if not sku:
                        continue
                    result[sku] = {c: clean_cell(row.get(c)) for c in db_cols}
    finally:
        conn.close()
    return result


def fetch_product_rows(skus: Optional[Sequence[str]] = None) -> List[Dict[str, Any]]:
    """读取 product_sku 行；``skus`` 非空时仅这些 SKU。"""
    db_cols = [FIX_DB_COL] + [f.db_col for f in UP_FIELD_MAP.values()]
    seen: List[str] = []
    for c in db_cols:
        if c not in seen:
            seen.append(c)

    db = get_db_manager()
    conn = db.get_connection()
    rows: List[Dict[str, Any]] = []
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            cur.execute(f"SHOW COLUMNS FROM `{TABLE}`")
            existing_cols = {r["Field"] for r in cur.fetchall()}
            select_cols = [c for c in seen if c in existing_cols]
            if FIX_DB_COL not in select_cols:
                select_cols.insert(0, FIX_DB_COL)
            col_sql = ", ".join(f"`{c}`" for c in select_cols)

            if skus:
                wanted = [clean_cell(s) for s in skus if clean_cell(s)]
                if not wanted:
                    return []
                for i in range(0, len(wanted), BATCH_SIZE):
                    chunk = wanted[i : i + BATCH_SIZE]
                    placeholders = ",".join(["%s"] * len(chunk))
                    cur.execute(
                        f"SELECT {col_sql} FROM `{TABLE}` "
                        f"WHERE `{FIX_DB_COL}` IN ({placeholders}) "
                        f"AND is_deleted = 0 "
                        f"ORDER BY id ASC",
                        chunk,
                    )
                    rows.extend(cur.fetchall())
            else:
                cur.execute(
                    f"SELECT {col_sql} FROM `{TABLE}` "
                    f"WHERE TRIM(`{FIX_DB_COL}`) <> '' AND is_deleted = 0 "
                    f"ORDER BY id ASC"
                )
                rows.extend(cur.fetchall())
    finally:
        conn.close()
    return rows


def _sheet_col_index(df: pd.DataFrame, sheet_col: str) -> int:
    """DataFrame 列位置 = 钉钉表 0-based 列下标（按 A1 已用区域读取时）。"""
    return list(df.columns).index(sheet_col)


def sheet_sku_set(df: pd.DataFrame) -> set[str]:
    """表格已有 SKU 集合。"""
    if SKU_SHEET_COL not in df.columns:
        return set()
    return {
        s
        for s in df[SKU_SHEET_COL].map(clean_cell).tolist()
        if s
    }


def build_append_dataframe(
    sheet_columns: Sequence[str],
    db_rows: Sequence[Mapping[str, Any]],
) -> pd.DataFrame:
    """按钉钉表头列顺序构造待追加 DataFrame。"""
    sheet_to_db = _sheet_to_db_col()
    records: List[Dict[str, str]] = []
    for row in db_rows:
        rec: Dict[str, str] = {}
        for col in sheet_columns:
            db_col = sheet_to_db.get(col)
            if not db_col:
                rec[col] = ""
                continue
            rec[col] = clean_cell(row.get(db_col))
        records.append(rec)
    return pd.DataFrame(records, columns=list(sheet_columns))


def sync_fields_to_sheet(
    wb: Workbook,
    sheet: str,
    df: pd.DataFrame,
    field_keys: Sequence[str],
    *,
    dry_run: bool = False,
    allow_empty: bool = False,
    skus: Optional[Sequence[str]] = None,
) -> Dict[str, UpdateStats]:
    """用 DB 值回写钉钉表；按 SKU 对齐行，仅改有差异的单元格。"""
    if SKU_SHEET_COL not in df.columns:
        raise KeyError(f"表格缺少列: [{SKU_SHEET_COL}]；实际列={list(df.columns)}")

    work = df
    excel_rows: Optional[List[int]] = None

    if skus:
        work, missing = filter_by_column(df, SKU_SHEET_COL, skus)
        if missing:
            print(f"[WARN] 表格中未找到 SKU: {', '.join(missing)}", file=sys.stderr)
        if work.empty:
            print("[WARN] 无匹配行，跳过更新", file=sys.stderr)
            return {}
        excel_rows = [int(idx) + 2 for idx in work.index.tolist()]

    sku_series = work[SKU_SHEET_COL].map(clean_cell)
    wanted_skus = [s for s in sku_series.tolist() if s]
    db_cols = [UP_FIELD_MAP[k].db_col for k in field_keys]
    seen_cols: List[str] = []
    for c in db_cols:
        if c not in seen_cols:
            seen_cols.append(c)
    db_map = fetch_product_map(wanted_skus, seen_cols)

    results: Dict[str, UpdateStats] = {}
    for key in field_keys:
        field = UP_FIELD_MAP[key]
        if field.sheet_col not in work.columns:
            results[key] = UpdateStats(sheet_rows=len(work), skipped_missing_col=True)
            print(f"[SKIP] {key}: 表格无列[{field.sheet_col}]", file=sys.stderr)
            continue

        col_idx = _sheet_col_index(df, field.sheet_col)
        stats = UpdateStats(sheet_rows=len(work), unique_skus=len(set(wanted_skus)))
        updates: List[Tuple[int, Any]] = []

        for i, sku in enumerate(sku_series.tolist()):
            if not sku:
                continue
            excel_row = excel_rows[i] if excel_rows is not None else (i + 2)
            sheet_val = clean_cell(work.iloc[i][field.sheet_col])
            db_row = db_map.get(sku)
            if db_row is None:
                stats.missing_in_db += 1
                continue
            db_val = db_row.get(field.db_col, "")
            if not db_val and not allow_empty:
                stats.empty_value += 1
                continue
            if db_val == sheet_val:
                stats.unchanged += 1
                continue
            cell_val: Any = db_val if db_val != "" else ""
            updates.append((excel_row, cell_val))

        stats.updated = len(updates)
        results[key] = stats
        verb = "would_update" if dry_run else "updated"
        print(
            f"[UPDATE] {TABLE}.{field.db_col} → [{field.sheet_col}]  "
            f"unique={stats.unique_skus} empty_skipped={stats.empty_value} "
            f"missing_in_db={stats.missing_in_db} unchanged={stats.unchanged} "
            f"{verb}={stats.updated}"
        )
        if not dry_run and updates:
            wb.write_column_updates(sheet, col_idx, updates)

    return results


def sync_new_rows_to_sheet(
    wb: Workbook,
    sheet: str,
    df: pd.DataFrame,
    *,
    dry_run: bool = False,
    skus: Optional[Sequence[str]] = None,
    preview: int = -1,
) -> AppendStats:
    """把库中有、表格无的 SKU 追加到钉钉表。"""
    stats = AppendStats()
    if SKU_SHEET_COL not in df.columns:
        stats.missing_sku_col = True
        return stats

    existing = sheet_sku_set(df)
    stats.sheet_skus = len(existing)

    db_rows = fetch_product_rows(skus)
    stats.db_rows = len(db_rows)

    new_rows: List[Dict[str, Any]] = []
    for row in db_rows:
        sku = clean_cell(row.get(FIX_DB_COL))
        if not sku:
            continue
        if sku in existing:
            stats.already_in_sheet += 1
            continue
        new_rows.append(row)

    stats.to_append = len(new_rows)
    if not new_rows:
        return stats

    append_df = build_append_dataframe(list(df.columns), new_rows)
    if preview >= 0:
        print(f"=== 待追加  shape={append_df.shape} ===")
        print(
            append_df.head(preview).to_string(index=False)
            if preview > 0
            else append_df.to_string(index=False)
        )
        print()

    if dry_run:
        return stats

    wb.append_dataframe(sheet, append_df, include_header=False)
    stats.appended = len(new_rows)
    return stats


def build_parser() -> argparse.ArgumentParser:
    up_choices = list(UP_FIELD_MAP.keys())
    parser = argparse.ArgumentParser(
        description="将 product_sku 同步到钉钉「产品信息」表格（DB → 钉钉：更新列 / 追加）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "匹配键: product_sku → SKU\n"
            "默认：更新表格已有行的白名单字段\n"
            "白名单字段:\n  "
            + ", ".join(up_choices)
            + "\n\n示例:\n"
            "  python app/ding-disk/product_sku_push.py\n"
            "  python app/ding-disk/product_sku_push.py --dry-run\n"
            "  python app/ding-disk/product_sku_push.py --sku XPM0213\n"
            "  python app/ding-disk/product_sku_push.py --up-field supplier_abbr\n"
            "  python app/ding-disk/product_sku_push.py --append --sku NEW001\n"
            "  python app/ding-disk/product_sku_push.py --append --update\n"
        ),
    )
    parser.add_argument("--workbook-id", default=WORKBOOK_ID, help="表格文档 ID")
    parser.add_argument("--list-sheets", action="store_true", help="仅列出工作表")
    parser.add_argument("--sheet", default=None, help=f"工作表名称；默认 {DEFAULT_SHEET}")
    parser.add_argument("--raw", action="store_true", help="--list-sheets 时缩进 JSON")
    parser.add_argument(
        "--sku",
        action="append",
        dest="skus",
        default=None,
        help="仅处理指定 SKU（可重复）；默认处理表格/库内全部",
    )
    parser.add_argument(
        "--append",
        action="store_true",
        help="追加库中有、表中无的新行",
    )
    parser.add_argument(
        "--update",
        action="store_true",
        help="更新表格已有行的白名单全部字段（未指定 --append 时为默认）",
    )
    parser.add_argument(
        "--up-field",
        action="append",
        choices=up_choices,
        dest="up_fields",
        default=None,
        help="只更新指定字段（可重复）；指定即进入更新模式",
    )
    parser.add_argument("--dry-run", action="store_true", help="统计变更，不写钉钉")
    parser.add_argument(
        "--allow-empty",
        action="store_true",
        help="允许用 DB 空值覆盖钉钉单元格（默认跳过空值）",
    )
    parser.add_argument(
        "--preview",
        type=int,
        default=-1,
        help="预览行数；更新模式预览表格，追加模式预览待追加；0=全量；默认 -1 不打印",
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)

    try:
        workbook_id = (args.workbook_id or WORKBOOK_ID).strip()
        wb = Workbook(workbook_id)
        sheets = wb.list_sheets()

        if args.list_sheets:
            payload: Dict[str, Any] = {"workbookId": workbook_id, "sheets": sheets}
            print(json.dumps(payload, ensure_ascii=False, indent=2 if args.raw else None))
            return 0

        if not sheets:
            print(f"[WARN] 表格无工作表: {workbook_id}", file=sys.stderr)
            return 1

        sheet = (args.sheet or "").strip() or DEFAULT_SHEET
        df = wb.read_sheet(sheet, header=True)

        do_append = bool(args.append)
        # 兼容旧用法：无 --append 时默认更新；有 --up-field/--update 也更新
        do_update = bool(args.update) or bool(args.up_fields) or not do_append

        if args.preview >= 0 and do_update and not do_append:
            print(f"=== {sheet}  shape={df.shape} ===")
            print(
                df.head(args.preview).to_string(index=False)
                if args.preview > 0
                else df.to_string(index=False)
            )
            print()

        mode = "dry-run" if args.dry_run else "write"
        sku_hint = f" skus={','.join(args.skus)}" if args.skus else ""
        actions = []
        if do_append:
            actions.append("append")
        if do_update:
            actions.append("update")
        print(f"[SYNC] DB→钉钉 sheet={sheet} mode={','.join(actions)} {mode}{sku_hint}")

        if do_update:
            field_keys = resolve_up_fields(args.up_fields)
            sync_fields_to_sheet(
                wb,
                sheet,
                df,
                field_keys,
                dry_run=bool(args.dry_run),
                allow_empty=bool(args.allow_empty),
                skus=args.skus,
            )

        if do_append:
            stats = sync_new_rows_to_sheet(
                wb,
                sheet,
                df,
                dry_run=bool(args.dry_run),
                skus=args.skus,
                preview=int(args.preview),
            )
            if stats.missing_sku_col:
                print(
                    f"[FAIL] 表格缺少列: [{SKU_SHEET_COL}]；实际列={list(df.columns)}",
                    file=sys.stderr,
                )
                return 2
            verb = "would_append" if args.dry_run else "appended"
            print(
                f"[APPEND] {TABLE} → [{SKU_SHEET_COL}]  "
                f"db_rows={stats.db_rows} sheet_skus={stats.sheet_skus} "
                f"already_in_sheet={stats.already_in_sheet} "
                f"to_append={stats.to_append} "
                f"{verb}={stats.appended if not args.dry_run else stats.to_append}"
            )

        return 0
    except DingDiskError as exc:
        print(f"[FAIL] DingDisk: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:  # noqa: BLE001
        print(f"[FAIL] {type(exc).__name__}: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
