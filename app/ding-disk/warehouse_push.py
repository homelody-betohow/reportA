"""将 ``warehouse`` 新增行同步到钉钉「仓库信息」表格（DB → 钉钉追加）。

与 ``warehouse_pull.py`` 方向相反：只处理表格中尚不存在的 ``warehouse_id``，
按表头列顺序追加整行（含仓库ID、仓库编码与白名单字段）。

匹配键::
    warehouse_id  →  仓库ID

文档 ID（workbookId / nodeId）::
    dpYLaezmVNwrRXb2tPm77keqJrMqPxX6

用法（项目根目录）::

    # 默认：找出库中有、表中无的仓库，追加到 Sheet1
    python app/ding-disk/warehouse_push.py
    python app/ding-disk/warehouse_push.py --dry-run

    # 只处理指定仓库ID
    python app/ding-disk/warehouse_push.py --id 237
    python app/ding-disk/warehouse_push.py --id 237 --dry-run --preview 5

    python app/ding-disk/warehouse_push.py --list-sheets
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional, Sequence

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
from api.ding_disk.workbook import Workbook, clean_cell  # noqa: E402

# 与 warehouse_pull.py 保持一致
WORKBOOK_ID = "dpYLaezmVNwrRXb2tPm77keqJrMqPxX6"
DEFAULT_SHEET = "Sheet1"
TABLE = "warehouse"
FIX_SHEET_COL = "仓库ID"
FIX_DB_COL = "warehouse_id"
BATCH_SIZE = 500


@dataclass(frozen=True)
class FieldMap:
    """字段名 → 钉钉列 → warehouse 列。"""

    sheet_col: str
    db_col: str


# 白名单：DB → 钉钉（追加行时写入；不含匹配键，匹配键单独处理）
# warehouse_code 虽不在 pull 白名单，追加新行时需写入「仓库编码」列
UP_FIELD_MAP: Mapping[str, FieldMap] = {
    "warehouse_code": FieldMap(sheet_col="仓库编码", db_col="warehouse_code"),
    "warehouse_name": FieldMap(sheet_col="仓库名称", db_col="warehouse_name"),
    "country_code": FieldMap(sheet_col="国家", db_col="country_code"),
    "provider_code": FieldMap(sheet_col="服务商", db_col="provider_code"),
    "market_code": FieldMap(sheet_col="销售平台", db_col="market_code"),
    "market_region": FieldMap(sheet_col="销售站点", db_col="market_region"),
    "ops_owner": FieldMap(sheet_col="运营负责人", db_col="ops_owner"),
    "is_transfer": FieldMap(sheet_col="可调拨", db_col="is_transfer"),
    "snapshot_inventory": FieldMap(sheet_col="库存快照", db_col="snapshot_inventory"),
    "warehouse_status": FieldMap(sheet_col="状态", db_col="warehouse_status"),
    "remark": FieldMap(sheet_col="备注", db_col="remark"),
}

# 库内值 → 表格文案（与 pull 的 FIELD_VALUE_MAP 互逆）
FIELD_DISPLAY_MAP: Mapping[str, Mapping[str, str]] = {
    "is_transfer": {
        "1": "是",
        "0": "否",
    },
    "snapshot_inventory": {
        "1": "是",
        "0": "否",
    },
    "warehouse_status": {
        "1": "启用",
        "0": "停用",
        "-1": "废弃",
    },
}


@dataclass
class AppendStats:
    db_rows: int = 0
    sheet_ids: int = 0
    already_in_sheet: int = 0
    to_append: int = 0
    appended: int = 0
    missing_id_col: bool = False


def _sheet_to_db_col() -> Dict[str, str]:
    """钉钉列名 → DB 列名（含匹配键）。"""
    mapping = {FIX_SHEET_COL: FIX_DB_COL}
    for field in UP_FIELD_MAP.values():
        mapping[field.sheet_col] = field.db_col
    return mapping


def _db_col_to_field_key() -> Dict[str, str]:
    return {f.db_col: key for key, f in UP_FIELD_MAP.items()}


def display_cell(field_key: Optional[str], value: str) -> str:
    """DB 值转钉钉展示文案。"""
    if not field_key:
        return value
    value_map = FIELD_DISPLAY_MAP.get(field_key)
    if not value_map:
        return value
    return value_map.get(value, value)


def fetch_warehouse_rows(
    warehouse_ids: Optional[Sequence[str]] = None,
) -> List[Dict[str, Any]]:
    """读取 warehouse 行（按 id）；``warehouse_ids`` 非空时仅这些 warehouse_id。"""
    db_cols = [FIX_DB_COL] + [f.db_col for f in UP_FIELD_MAP.values()]
    seen: List[str] = []
    for c in db_cols:
        if c not in seen:
            seen.append(c)
    # remark 可能尚未建列：只选实际存在的列
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

            if warehouse_ids:
                wanted = [clean_cell(i) for i in warehouse_ids if clean_cell(i)]
                if not wanted:
                    return []
                for i in range(0, len(wanted), BATCH_SIZE):
                    chunk = wanted[i : i + BATCH_SIZE]
                    placeholders = ",".join(["%s"] * len(chunk))
                    cur.execute(
                        f"SELECT {col_sql} FROM `{TABLE}` "
                        f"WHERE `{FIX_DB_COL}` IN ({placeholders}) "
                        f"ORDER BY id ASC",
                        chunk,
                    )
                    rows.extend(cur.fetchall())
            else:
                cur.execute(
                    f"SELECT {col_sql} FROM `{TABLE}` "
                    f"WHERE `{FIX_DB_COL}` IS NOT NULL "
                    f"ORDER BY id ASC"
                )
                rows.extend(cur.fetchall())
    finally:
        conn.close()
    return rows


def sheet_id_set(df: pd.DataFrame) -> set[str]:
    """表格已有仓库ID集合。"""
    if FIX_SHEET_COL not in df.columns:
        return set()
    return {
        wid
        for wid in df[FIX_SHEET_COL].map(clean_cell).tolist()
        if wid
    }


def build_append_dataframe(
    sheet_columns: Sequence[str],
    db_rows: Sequence[Mapping[str, Any]],
) -> pd.DataFrame:
    """按钉钉表头列顺序构造待追加 DataFrame。"""
    sheet_to_db = _sheet_to_db_col()
    db_to_field = _db_col_to_field_key()
    records: List[Dict[str, str]] = []
    for row in db_rows:
        rec: Dict[str, str] = {}
        for col in sheet_columns:
            db_col = sheet_to_db.get(col)
            if not db_col:
                rec[col] = ""
                continue
            raw = clean_cell(row.get(db_col))
            field_key = db_to_field.get(db_col)
            rec[col] = display_cell(field_key, raw)
        records.append(rec)
    return pd.DataFrame(records, columns=list(sheet_columns))


def sync_new_rows_to_sheet(
    wb: Workbook,
    sheet: str,
    df: pd.DataFrame,
    *,
    dry_run: bool = False,
    warehouse_ids: Optional[Sequence[str]] = None,
    preview: int = -1,
) -> AppendStats:
    """把库中有、表格无的仓库追加到钉钉表。"""
    stats = AppendStats()
    if FIX_SHEET_COL not in df.columns:
        stats.missing_id_col = True
        return stats

    existing = sheet_id_set(df)
    stats.sheet_ids = len(existing)

    db_rows = fetch_warehouse_rows(warehouse_ids)
    stats.db_rows = len(db_rows)

    new_rows: List[Dict[str, Any]] = []
    for row in db_rows:
        wid = clean_cell(row.get(FIX_DB_COL))
        if not wid:
            continue
        if wid in existing:
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
    parser = argparse.ArgumentParser(
        description="将 warehouse 新增行追加到钉钉「仓库信息」表格（DB → 钉钉）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "匹配键: warehouse_id → 仓库ID\n"
            "追加列: 仓库ID + 仓库编码 + 白名单字段（按表头对齐，未知列留空）\n\n"
            "示例:\n"
            "  python app/ding-disk/warehouse_push.py\n"
            "  python app/ding-disk/warehouse_push.py --dry-run\n"
            "  python app/ding-disk/warehouse_push.py --id 237\n"
            "  python app/ding-disk/warehouse_push.py --dry-run --preview 10\n"
        ),
    )
    parser.add_argument("--workbook-id", default=WORKBOOK_ID, help="表格文档 ID")
    parser.add_argument("--list-sheets", action="store_true", help="仅列出工作表")
    parser.add_argument("--sheet", default=None, help=f"工作表名称；默认 {DEFAULT_SHEET}")
    parser.add_argument("--raw", action="store_true", help="--list-sheets 时缩进 JSON")
    parser.add_argument(
        "--id",
        action="append",
        dest="warehouse_ids",
        default=None,
        help="仅处理指定仓库ID（可重复）；默认对比库内全部",
    )
    parser.add_argument("--dry-run", action="store_true", help="统计待追加行，不写钉钉")
    parser.add_argument(
        "--preview",
        type=int,
        default=-1,
        help="预览待追加行数；0=全量；默认 -1 不打印",
    )
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)

    try:
        workbook_id = (args.workbook_id or WORKBOOK_ID).strip()
        if not workbook_id:
            print(
                "[FAIL] 未指定表格文档 ID：请设置 WORKBOOK_ID 或传入 --workbook-id",
                file=sys.stderr,
            )
            return 2

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

        mode = "dry-run" if args.dry_run else "write"
        id_hint = f" ids={','.join(args.warehouse_ids)}" if args.warehouse_ids else ""
        print(f"[SYNC] DB→钉钉追加 sheet={sheet}{id_hint} mode={mode}")

        stats = sync_new_rows_to_sheet(
            wb,
            sheet,
            df,
            dry_run=bool(args.dry_run),
            warehouse_ids=args.warehouse_ids,
            preview=int(args.preview),
        )

        if stats.missing_id_col:
            print(
                f"[FAIL] 表格缺少列: [{FIX_SHEET_COL}]；实际列={list(df.columns)}",
                file=sys.stderr,
            )
            return 2

        verb = "would_append" if args.dry_run else "appended"
        print(
            f"[APPEND] {TABLE} → [{FIX_SHEET_COL}]  "
            f"db_rows={stats.db_rows} sheet_ids={stats.sheet_ids} "
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
