"""将 ``platform_shop`` 同步到钉钉「店铺信息」表格（DB → 钉钉）。

与 ``platform_shop_pull.py`` 方向相反，支持两种模式：

1. **追加**（默认）：表格中尚不存在的 ``shop_hash`` 按表头追加整行。
2. **更新列**：表格已有行按 ``shop_hash`` 对齐，仅回写白名单中有差异的单元格。

匹配键::
    shop_hash  →  店铺编码

文档 ID（workbookId / nodeId）::
    NZQYprEoWo75xoEDtBqGzKqPW1waOeDk

用法（项目根目录）::

    # 默认：找出库中有、表中无的店铺，追加到 Sheet1
    python app/ding-disk/platform_shop_push.py
    python app/ding-disk/platform_shop_push.py --dry-run

    # 更新已有行的白名单列（全部 / 指定列）
    python app/ding-disk/platform_shop_push.py --update
    python app/ding-disk/platform_shop_push.py --up-field store_fees
    python app/ding-disk/platform_shop_push.py --hash ecdfa488... --up-field vat_rate --dry-run

    # 追加 + 更新可同时进行
    python app/ding-disk/platform_shop_push.py --append --update

    python app/ding-disk/platform_shop_push.py --list-sheets
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


@dataclass
class UpdateStats:
    sheet_rows: int = 0
    unique_keys: int = 0
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


def fetch_shop_map(
    hashes: Sequence[str],
    db_cols: Sequence[str],
) -> Dict[str, Dict[str, str]]:
    """批量读取 ``platform_shop`` → ``{shop_hash: {db_col: value}}``。"""
    if not hashes or not db_cols:
        return {}
    seen: List[str] = []
    for c in db_cols:
        if c not in seen:
            seen.append(c)
    col_sql = ", ".join(f"`{c}`" for c in seen)
    result: Dict[str, Dict[str, str]] = {}
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            for i in range(0, len(hashes), BATCH_SIZE):
                chunk = list(hashes[i : i + BATCH_SIZE])
                placeholders = ",".join(["%s"] * len(chunk))
                cur.execute(
                    f"SELECT `{FIX_DB_COL}`, {col_sql} FROM `{TABLE}` "
                    f"WHERE `{FIX_DB_COL}` IN ({placeholders})",
                    chunk,
                )
                for row in cur.fetchall():
                    shop_hash = clean_cell(row[FIX_DB_COL])
                    if not shop_hash:
                        continue
                    result[shop_hash] = {c: clean_cell(row.get(c)) for c in seen}
    finally:
        conn.close()
    return result


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
                wanted = [clean_cell(h) for h in hashes if clean_cell(h)]
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


def _sheet_col_index(df: pd.DataFrame, sheet_col: str) -> int:
    """DataFrame 列位置 = 钉钉表 0-based 列下标。"""
    return list(df.columns).index(sheet_col)


def sheet_hash_set(df: pd.DataFrame) -> set[str]:
    """表格已有店铺编码集合。"""
    if FIX_SHEET_COL not in df.columns:
        return set()
    return {
        h
        for h in df[FIX_SHEET_COL].map(clean_cell).tolist()
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
            raw = clean_cell(row.get(db_col))
            field_key = db_to_field.get(db_col)
            rec[col] = display_cell(field_key, raw)
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
    hashes: Optional[Sequence[str]] = None,
) -> Dict[str, UpdateStats]:
    """用 DB 值回写钉钉表已有行；按 shop_hash 对齐，仅改有差异的单元格。"""
    if FIX_SHEET_COL not in df.columns:
        raise KeyError(f"表格缺少列: [{FIX_SHEET_COL}]；实际列={list(df.columns)}")

    work = df
    excel_rows: Optional[List[int]] = None

    if hashes:
        work, missing = filter_by_column(df, FIX_SHEET_COL, hashes)
        if missing:
            print(f"[WARN] 表格中未找到店铺编码: {', '.join(missing)}", file=sys.stderr)
        if work.empty:
            print("[WARN] 无匹配行，跳过更新", file=sys.stderr)
            return {}
        excel_rows = [int(idx) + 2 for idx in work.index.tolist()]

    hash_series = work[FIX_SHEET_COL].map(clean_cell)
    wanted_hashes = [h for h in hash_series.tolist() if h]
    db_cols = [UP_FIELD_MAP[k].db_col for k in field_keys]
    seen_cols: List[str] = []
    for c in db_cols:
        if c not in seen_cols:
            seen_cols.append(c)
    db_map = fetch_shop_map(wanted_hashes, seen_cols)

    results: Dict[str, UpdateStats] = {}
    for key in field_keys:
        field = UP_FIELD_MAP[key]
        if field.sheet_col not in work.columns:
            results[key] = UpdateStats(sheet_rows=len(work), skipped_missing_col=True)
            print(f"[SKIP] {key}: 表格无列[{field.sheet_col}]", file=sys.stderr)
            continue

        col_idx = _sheet_col_index(df, field.sheet_col)
        stats = UpdateStats(sheet_rows=len(work), unique_keys=len(set(wanted_hashes)))
        updates: List[Tuple[int, Any]] = []

        for i, shop_hash in enumerate(hash_series.tolist()):
            if not shop_hash:
                continue
            excel_row = excel_rows[i] if excel_rows is not None else (int(work.index[i]) + 2)
            sheet_val = clean_cell(work.iloc[i][field.sheet_col])
            db_row = db_map.get(shop_hash)
            if db_row is None:
                stats.missing_in_db += 1
                continue
            db_val = display_cell(key, db_row.get(field.db_col, ""))
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
            f"unique={stats.unique_keys} empty_skipped={stats.empty_value} "
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
        shop_hash = clean_cell(row.get(FIX_DB_COL))
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
    up_choices = list(UP_FIELD_MAP.keys())
    parser = argparse.ArgumentParser(
        description="将 platform_shop 同步到钉钉「店铺信息」表格（DB → 钉钉：追加 / 更新列）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "匹配键: shop_hash → 店铺编码\n"
            "默认：追加库中有、表中无的新行\n"
            "白名单字段:\n  "
            + ", ".join(up_choices)
            + "\n\n示例:\n"
            "  python app/ding-disk/platform_shop_push.py\n"
            "  python app/ding-disk/platform_shop_push.py --dry-run\n"
            "  python app/ding-disk/platform_shop_push.py --update\n"
            "  python app/ding-disk/platform_shop_push.py --up-field store_fees\n"
            "  python app/ding-disk/platform_shop_push.py --hash <shop_hash> --up-field vat_rate\n"
            "  python app/ding-disk/platform_shop_push.py --append --update --dry-run\n"
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
        help="仅处理指定店铺编码 shop_hash（可重复）；默认处理表格/库内全部",
    )
    parser.add_argument(
        "--append",
        action="store_true",
        default=None,
        help="追加库中有、表中无的新行（未指定 --update/--up-field 时为默认）",
    )
    parser.add_argument(
        "--update",
        action="store_true",
        help="更新表格已有行的白名单全部字段",
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

        do_update = bool(args.update) or bool(args.up_fields)
        if args.append is not None:
            do_append = bool(args.append)
        else:
            do_append = not do_update

        if not do_append and not do_update:
            print("[FAIL] 请指定 --append 和/或 --update（或 --up-field）", file=sys.stderr)
            return 2

        mode = "dry-run" if args.dry_run else "write"
        hash_hint = f" hashes={','.join(args.hashes)}" if args.hashes else ""
        actions = []
        if do_append:
            actions.append("append")
        if do_update:
            actions.append("update")
        print(f"[SYNC] DB→钉钉 sheet={sheet} mode={','.join(actions)} {mode}{hash_hint}")

        if do_update:
            field_keys = resolve_up_fields(args.up_fields)
            sync_fields_to_sheet(
                wb,
                sheet,
                df,
                field_keys,
                dry_run=bool(args.dry_run),
                allow_empty=bool(args.allow_empty),
                hashes=args.hashes,
            )

        if do_append:
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
