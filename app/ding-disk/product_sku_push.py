"""根据 ``product_sku`` 表内容，回写钉钉「产品信息」表格白名单字段。

与 ``productInfo.py`` 方向相反：DB → 钉钉表。
表格读写委托 ``api.ding_disk.workbook.Workbook``。

文档 ID（workbookId / nodeId）::
    Obva6QBXJwjBxoE2sM62MrzGVn4qY5Pr

用法（项目根目录）::

    # 默认：读 Sheet1，用 DB 更新白名单全部字段
    python app/ding-disk/product_sku_push.py
    python app/ding-disk/product_sku_push.py --dry-run

    # 只更新部分字段 / 指定 SKU
    python app/ding-disk/product_sku_push.py --up-field supplier_abbr
    python app/ding-disk/product_sku_push.py --sku XPM0213
    python app/ding-disk/product_sku_push.py --sku XPM0213 --up-field hs_code --dry-run

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
BATCH_SIZE = 500


@dataclass(frozen=True)
class FieldMap:
    """字段名 → 钉钉列 → product_sku 列。"""

    sheet_col: str
    db_col: str


# 白名单：DB → 钉钉；默认全部更新（与 productInfo 对齐）
UP_FIELD_MAP: Mapping[str, FieldMap] = {
    "supplier_abbr": FieldMap(sheet_col="供应商", db_col="supplier_abbr"),
    "supplier_name": FieldMap(sheet_col="供应商全称", db_col="supplier_name"),
    "ops_model": FieldMap(sheet_col="运营模式", db_col="ops_model"),
    "product_uid": FieldMap(sheet_col="商品ID", db_col="product_uid"),
    "category_lv1": FieldMap(sheet_col="一级分类", db_col="category_lv1"),
    "category_lv2": FieldMap(sheet_col="二级分类", db_col="category_lv2"),
    "category_lv3": FieldMap(sheet_col="三级分类", db_col="category_lv3"),
    "declare_price": FieldMap(sheet_col="申报价值USD", db_col="declare_price_usd"),
    "declare_name_cn": FieldMap(sheet_col="申报中文", db_col="declare_name_cn"),
    "declare_name_en": FieldMap(sheet_col="申报英文", db_col="declare_name_en"),
    "hs_code": FieldMap(sheet_col="海关编码", db_col="hs_code"),
    "amz_lifecycle": FieldMap(sheet_col="AMZ状态", db_col="amz_lifecycle"),
    "local_lifecycle": FieldMap(sheet_col="本土状态", db_col="local_lifecycle"),
    "accounting_class": FieldMap(sheet_col="核算分类", db_col="accounting_class"),
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


def fetch_product_map(
    skus: Sequence[str],
    db_cols: Sequence[str],
) -> Dict[str, Dict[str, str]]:
    """批量读取 ``product_sku`` → ``{sku: {db_col: value}}``。"""
    if not skus or not db_cols:
        return {}
    # db_cols 来自白名单
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


def _sheet_col_index(df: pd.DataFrame, sheet_col: str) -> int:
    """DataFrame 列位置 = 钉钉表 0-based 列下标（按 A1 已用区域读取时）。"""
    return list(df.columns).index(sheet_col)


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
        # 保留相对原表的行号：用原 df 索引定位
        excel_rows = [int(idx) + 2 for idx in work.index.tolist()]

    sku_series = work[SKU_SHEET_COL].map(clean_cell)
    wanted_skus = [s for s in sku_series.tolist() if s]
    db_cols = [UP_FIELD_MAP[k].db_col for k in field_keys]
    # 去重保持顺序
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
            # 空串写 None，避免钉钉写成字面 "None"
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


def build_parser() -> argparse.ArgumentParser:
    up_choices = list(UP_FIELD_MAP.keys())
    parser = argparse.ArgumentParser(
        description="用 product_sku 表回写钉钉「产品信息」白名单字段（DB → 钉钉）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "默认更新字段:\n  "
            + ", ".join(up_choices)
            + "\n\n示例:\n"
            "  python app/ding-disk/upProductInfo.py\n"
            "  python app/ding-disk/upProductInfo.py --dry-run\n"
            "  python app/ding-disk/upProductInfo.py --sku XPM0213\n"
            "  python app/ding-disk/upProductInfo.py --up-field supplier_abbr\n"
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
        help="仅更新指定 SKU（可重复）；默认更新表格中全部有匹配的 SKU",
    )
    parser.add_argument(
        "--up-field",
        action="append",
        choices=up_choices,
        dest="up_fields",
        default=None,
        help="只更新指定字段（可重复）；默认更新白名单全部字段",
    )
    parser.add_argument("--dry-run", action="store_true", help="统计将更新的单元格，不写钉钉")
    parser.add_argument(
        "--allow-empty",
        action="store_true",
        help="允许用 DB 空值覆盖钉钉单元格（默认跳过空值）",
    )
    parser.add_argument(
        "--preview",
        type=int,
        default=-1,
        help="同步前预览表格行数；0=全量；默认 -1 不打印",
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
        if args.preview >= 0:
            print(f"=== {sheet}  shape={df.shape} ===")
            print(
                df.head(args.preview).to_string(index=False)
                if args.preview > 0
                else df.to_string(index=False)
            )
            print()

        field_keys = resolve_up_fields(args.up_fields)
        mode = "dry-run" if args.dry_run else "write"
        sku_hint = f" skus={','.join(args.skus)}" if args.skus else ""
        print(f"[SYNC] DB→钉钉 sheet={sheet} fields={','.join(field_keys)}{sku_hint} mode={mode}")

        sync_fields_to_sheet(
            wb,
            sheet,
            df,
            field_keys,
            dry_run=bool(args.dry_run),
            allow_empty=bool(args.allow_empty),
            skus=args.skus,
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
