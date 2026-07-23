"""将 ``platform_shop`` 新增行同步到钉钉「店铺信息」表格（DB → 钉钉追加）。

与 ``platform_shop_pull.py`` 方向相反：只处理表格中尚不存在的 ``shop_hash``，
按表头列顺序追加整行（含店铺编码与白名单字段）。

匹配键::
    shop_hash  →  店铺编码

文档 ID（workbookId / nodeId）::
    NZQYprEoWo75xoEDtBqGzKqPW1waOeDk

用法（项目根目录）::

    # 默认：找出库中有、表中无的店铺，追加到 Sheet1
    python app/ding-disk/platform_shop_push.py
    python app/ding-disk/platform_shop_push.py --dry-run

    # 只处理指定店铺编码
    python app/ding-disk/platform_shop_push.py --hash ecdfa4883185ebf5...
    python app/ding-disk/platform_shop_push.py --hash ecdfa488... --dry-run --preview 5

    python app/ding-disk/platform_shop_push.py --list-sheets
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
from api.ding_disk.workbook import Workbook, normalize_cell  # noqa: E402

# 与 platform_shop_pull.py 保持一致
WORKBOOK_ID = "NZQYprEoWo75xoEDtBqGzKqPW1waOeDk"
DEFAULT_SHEET = "Sheet1"
TABLE = "platform_shop"
FIX_SHEET_COL = "店铺编码"
FIX_DB_COL = "shop_hash"
BATCH_SIZE = 500


@dataclass(frozen=True)
class FieldMap:
    """字段名 → 钉钉列 → platform_shop 列。"""

    sheet_col: str
    db_col: str


# 白名单：DB → 钉钉（追加行时写入；不含匹配键，匹配键单独处理）
UP_FIELD_MAP: Mapping[str, FieldMap] = {
    "shop_name_en": FieldMap(sheet_col="店铺名称", db_col="shop_name_en"),
    "platform": FieldMap(sheet_col="平台", db_col="platform"),
    "platform_site": FieldMap(sheet_col="站点", db_col="platform_site"),
    "market_code": FieldMap(sheet_col="销售平台", db_col="market_code"),
    "market_region": FieldMap(sheet_col="销售站点", db_col="market_region"),
    "commission_rate": FieldMap(sheet_col="平台费率", db_col="commission_rate"),
    "vat_rate": FieldMap(sheet_col="VAT费率", db_col="vat_rate"),
    "store_fees": FieldMap(sheet_col="月租", db_col="store_fees"),
    "currency": FieldMap(sheet_col="币种", db_col="currency"),
    "ops_leader": FieldMap(sheet_col="运营经理", db_col="ops_leader"),
    "ops_owner": FieldMap(sheet_col="运营负责人", db_col="ops_owner"),
    "shop_status": FieldMap(sheet_col="状态", db_col="shop_status"),
    "remark": FieldMap(sheet_col="备注", db_col="remark"),
}

# 库内值 → 表格文案（与 pull 的 FIELD_VALUE_MAP 互逆）
FIELD_DISPLAY_MAP: Mapping[str, Mapping[str, str]] = {
    "shop_status": {
        "1": "正常",
        "0": "停用",
    },
}


@dataclass
class AppendStats:
    db_rows: int = 0
    sheet_hashes: int = 0
    already_in_sheet: int = 0
    to_append: int = 0
    appended: int = 0
    missing_hash_col: bool = False


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


def fetch_shop_rows(hashes: Optional[Sequence[str]] = None) -> List[Dict[str, Any]]:
    """读取 platform_shop 行（按 id）；``hashes`` 非空时仅这些 shop_hash。"""
    db_cols = [FIX_DB_COL] + [f.db_col for f in UP_FIELD_MAP.values()]
    # 去重保持顺序
    seen: List[str] = []
    for c in db_cols:
        if c not in seen:
            seen.append(c)
    col_sql = ", ".join(f"`{c}`" for c in seen)

    db = get_db_manager()
    conn = db.get_connection()
    rows: List[Dict[str, Any]] = []
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            if hashes:
                wanted = [normalize_cell(h) for h in hashes if normalize_cell(h)]
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
                    f"WHERE TRIM(`{FIX_DB_COL}`) <> '' "
                    f"ORDER BY id ASC"
                )
                rows.extend(cur.fetchall())
    finally:
        conn.close()
    return rows


def sheet_hash_set(df: pd.DataFrame) -> set[str]:
    """表格已有店铺编码集合。"""
    if FIX_SHEET_COL not in df.columns:
        return set()
    return {
        h
        for h in df[FIX_SHEET_COL].map(normalize_cell).tolist()
        if h
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
            raw = normalize_cell(row.get(db_col))
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
    hashes: Optional[Sequence[str]] = None,
    preview: int = -1,
) -> AppendStats:
    """把库中有、表格无的店铺追加到钉钉表。"""
    stats = AppendStats()
    if FIX_SHEET_COL not in df.columns:
        stats.missing_hash_col = True
        return stats

    existing = sheet_hash_set(df)
    stats.sheet_hashes = len(existing)

    db_rows = fetch_shop_rows(hashes)
    stats.db_rows = len(db_rows)

    new_rows: List[Dict[str, Any]] = []
    for row in db_rows:
        shop_hash = normalize_cell(row.get(FIX_DB_COL))
        if not shop_hash:
            continue
        if shop_hash in existing:
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
        description="将 platform_shop 新增行追加到钉钉「店铺信息」表格（DB → 钉钉）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "匹配键: shop_hash → 店铺编码\n"
            "追加列: 店铺编码 + 白名单字段（按表头对齐，未知列留空）\n\n"
            "示例:\n"
            "  python app/ding-disk/platform_shop_push.py\n"
            "  python app/ding-disk/platform_shop_push.py --dry-run\n"
            "  python app/ding-disk/platform_shop_push.py --hash <shop_hash>\n"
            "  python app/ding-disk/platform_shop_push.py --dry-run --preview 10\n"
        ),
    )
    parser.add_argument("--workbook-id", default=WORKBOOK_ID, help="表格文档 ID")
    parser.add_argument("--list-sheets", action="store_true", help="仅列出工作表")
    parser.add_argument("--sheet", default=None, help=f"工作表名称；默认 {DEFAULT_SHEET}")
    parser.add_argument("--raw", action="store_true", help="--list-sheets 时缩进 JSON")
    parser.add_argument(
        "--hash",
        action="append",
        dest="hashes",
        default=None,
        help="仅处理指定店铺编码 shop_hash（可重复）；默认对比库内全部",
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
        hash_hint = f" hashes={','.join(args.hashes)}" if args.hashes else ""
        print(f"[SYNC] DB→钉钉追加 sheet={sheet}{hash_hint} mode={mode}")

        stats = sync_new_rows_to_sheet(
            wb,
            sheet,
            df,
            dry_run=bool(args.dry_run),
            hashes=args.hashes,
            preview=int(args.preview),
        )

        if stats.missing_hash_col:
            print(
                f"[FAIL] 表格缺少列: [{FIX_SHEET_COL}]；实际列={list(df.columns)}",
                file=sys.stderr,
            )
            return 2

        verb = "would_append" if args.dry_run else "appended"
        print(
            f"[APPEND] {TABLE} → [{FIX_SHEET_COL}]  "
            f"db_rows={stats.db_rows} sheet_hashes={stats.sheet_hashes} "
            f"already_in_sheet={stats.already_in_sheet} "
            f"to_append={stats.to_append} {verb}={stats.appended if not args.dry_run else stats.to_append}"
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
