"""读取钉钉在线表格「仓库信息」，并回写 ``warehouse`` 白名单字段。

表格读写委托 ``api.ding_disk.workbook.Workbook``；本脚本只负责字段映射与 DB 更新。

匹配键::
    仓库ID  →  warehouse_id

文档 ID（workbookId / nodeId）请填下方 ``WORKBOOK_ID``，或运行时 ``--workbook-id``。

用法（项目根目录）::

    # 只更新部分字段 / 指定仓库ID
    python app/ding-disk/warehouse_pull.py --up-field ops_owner
    python app/ding-disk/warehouse_pull.py --id 237
    python app/ding-disk/warehouse_pull.py --id 237 --up-field warehouse_status --dry-run

    # 只读（不写库）
    python app/ding-disk/warehouse_pull.py --no-update --preview 20
    python app/ding-disk/warehouse_pull.py --list-sheets
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from dataclasses import dataclass, field
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
    clean_pairs,
    filter_by_column,
    kv_pairs_from_df,
)

# 钉钉表格文档 ID（知识库 nodeId / dentryUuid）；空则必须传 --workbook-id
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


# 白名单：仅允许更新这些字段；默认全部更新（不含匹配键 warehouse_id）
UP_FIELD_MAP: Mapping[str, FieldMap] = {
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

# 表格文案 → 库内值；亦接受已是数字的写法
# warehouse_status：1=可用 0=不可用 -1=已废弃
FIELD_VALUE_MAP: Mapping[str, Mapping[str, str]] = {
    "is_transfer": {
        "是": "1",
        "否": "0",
        "1": "1",
        "0": "0",
    },
    "snapshot_inventory": {
        "是": "1",
        "否": "0",
        "1": "1",
        "0": "0",
    },
    "warehouse_status": {
        "启用": "1",
        "可用": "1",
        "停用": "0",
        "不可用": "0",
        "废弃": "-1",
        "已废弃": "-1",
        "1": "1",
        "0": "0",
        "-1": "-1",
    },
}


# 自由文本：保留字面量 "None"/"none"（默认 clean_cell 会当成空而跳过）
PRESERVE_NULLISH_FIELDS = frozenset(
    {
        "warehouse_name",
        "country_code",
        "provider_code",
        "market_code",
        "market_region",
        "ops_owner",
        "remark",
    }
)


@dataclass
class UpdateStats:
    sheet_rows: int = 0
    unique_keys: int = 0
    empty_value: int = 0
    invalid_value: int = 0
    missing_in_db: int = 0
    unchanged: int = 0
    updated: int = 0
    skipped_missing_col: bool = False
    # (warehouse_id, old_value, new_value)
    changes: List[Tuple[str, str, str]] = field(default_factory=list)


def map_field_values(
    field_key: str,
    pairs: Mapping[str, str],
    *,
    nullish_as_empty: bool = True,
) -> Tuple[Dict[str, str], int]:
    """按 FIELD_VALUE_MAP 转换单元格值；未配置映射则原样返回。

    返回 ``(mapped_pairs, invalid_skipped)``。映射表存在但值不在表内时跳过。
    """
    value_map = FIELD_VALUE_MAP.get(field_key)
    if not value_map:
        return clean_pairs(pairs, nullish_as_empty=nullish_as_empty), 0
    mapped: Dict[str, str] = {}
    invalid = 0
    for key, raw in clean_pairs(pairs, nullish_as_empty=nullish_as_empty).items():
        if raw in value_map:
            mapped[key] = value_map[raw]
        else:
            invalid += 1
    return mapped, invalid


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


def load_frames(
    wb: Workbook,
    *,
    sheet: Optional[str],
    read_all: bool,
    header: bool,
) -> Dict[str, pd.DataFrame]:
    """通过 Workbook 读取工作表。"""
    if read_all:
        return wb.read_all_sheets(header=header)
    target = (sheet or "").strip() or DEFAULT_SHEET
    return {target: wb.read_sheet(target, header=header)}


def fetch_existing_values(
    ids: Sequence[str],
    db_col: str,
    *,
    nullish_as_empty: bool = True,
) -> Dict[str, str]:
    """批量读取 warehouse 当前字段值（按 warehouse_id）。"""
    if not ids:
        return {}
    result: Dict[str, str] = {}
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            for i in range(0, len(ids), BATCH_SIZE):
                chunk = list(ids[i : i + BATCH_SIZE])
                placeholders = ",".join(["%s"] * len(chunk))
                cur.execute(
                    f"SELECT `{FIX_DB_COL}`, `{db_col}` AS v FROM `{TABLE}` "
                    f"WHERE `{FIX_DB_COL}` IN ({placeholders})",
                    chunk,
                )
                for row in cur.fetchall():
                    result[clean_cell(row[FIX_DB_COL])] = clean_cell(
                        row.get("v"),
                        nullish_as_empty=nullish_as_empty,
                    )
    finally:
        conn.close()
    return result


def apply_field_updates(
    pairs: Mapping[str, str],
    *,
    db_col: str,
    dry_run: bool = False,
    nullish_as_empty: bool = True,
) -> UpdateStats:
    """按 warehouse_id 更新 ``warehouse.<db_col>``；仅写入有变化的行。

    写库前再次 ``clean_cell``，确保两端空格/不可见字符不会入库。
    """
    pairs = clean_pairs(pairs, nullish_as_empty=nullish_as_empty)
    stats = UpdateStats(unique_keys=len(pairs))
    if not pairs:
        return stats

    existing = fetch_existing_values(
        list(pairs.keys()),
        db_col,
        nullish_as_empty=nullish_as_empty,
    )
    to_update: List[Tuple[str, str]] = []  # (value, warehouse_id)
    for wid, value in pairs.items():
        if wid not in existing:
            stats.missing_in_db += 1
            continue
        old = existing[wid]
        if old == value:
            stats.unchanged += 1
            continue
        to_update.append((value, wid))
        stats.changes.append((wid, old, value))

    stats.updated = len(to_update)
    if dry_run or not to_update:
        return stats

    sql = f"UPDATE `{TABLE}` SET `{db_col}` = %s WHERE `{FIX_DB_COL}` = %s"
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor() as cur:
            for i in range(0, len(to_update), BATCH_SIZE):
                cur.executemany(sql, to_update[i : i + BATCH_SIZE])
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return stats


def update_field_from_df(
    df: pd.DataFrame,
    field_key: str,
    *,
    dry_run: bool = False,
    allow_empty: bool = False,
) -> UpdateStats:
    field = UP_FIELD_MAP.get(field_key)
    if field is None:
        known = ", ".join(sorted(UP_FIELD_MAP))
        raise ValueError(f"未知字段={field_key!r}；可选: {known}")

    if FIX_SHEET_COL not in df.columns or field.sheet_col not in df.columns:
        return UpdateStats(sheet_rows=len(df), skipped_missing_col=True)

    nullish_as_empty = field_key not in PRESERVE_NULLISH_FIELDS
    pairs, empty_skipped = kv_pairs_from_df(
        df,
        FIX_SHEET_COL,
        field.sheet_col,
        allow_empty=allow_empty,
        nullish_as_empty=nullish_as_empty,
    )
    pairs, invalid_skipped = map_field_values(
        field_key,
        pairs,
        nullish_as_empty=nullish_as_empty,
    )
    stats = apply_field_updates(
        pairs,
        db_col=field.db_col,
        dry_run=dry_run,
        nullish_as_empty=nullish_as_empty,
    )
    stats.sheet_rows = len(df)
    stats.empty_value = empty_skipped
    stats.invalid_value = invalid_skipped
    return stats


def update_fields_from_df(
    df: pd.DataFrame,
    field_keys: Sequence[str],
    *,
    dry_run: bool = False,
    allow_empty: bool = False,
    warehouse_ids: Optional[Sequence[str]] = None,
) -> Dict[str, UpdateStats]:
    """依次更新白名单字段；缺列跳过。``warehouse_ids`` 非空时仅更新这些仓库ID。"""
    if FIX_SHEET_COL not in df.columns:
        raise KeyError(f"表格缺少列: [{FIX_SHEET_COL}]；实际列={list(df.columns)}")

    work = df
    if warehouse_ids:
        work, missing = filter_by_column(df, FIX_SHEET_COL, warehouse_ids)
        if missing:
            print(f"[WARN] 表格中未找到仓库ID: {', '.join(missing)}", file=sys.stderr)
        if work.empty:
            print("[WARN] 无匹配行，跳过更新", file=sys.stderr)
            return {}

    results: Dict[str, UpdateStats] = {}
    for key in field_keys:
        field = UP_FIELD_MAP[key]
        stats = update_field_from_df(
            work,
            key,
            dry_run=dry_run,
            allow_empty=allow_empty,
        )
        results[key] = stats
        if stats.skipped_missing_col:
            print(f"[SKIP] {key}: 表格无列[{field.sheet_col}]", file=sys.stderr)
            continue
        if stats.invalid_value:
            print(
                f"[WARN] {key}: 无法识别的值跳过 {stats.invalid_value} 条"
                f"（见 FIELD_VALUE_MAP[{key!r}]）",
                file=sys.stderr,
            )
        verb = "would_update" if dry_run else "updated"
        print(
            f"[UPDATE] [{field.sheet_col}] → {TABLE}.{field.db_col}  "
            f"unique={stats.unique_keys} empty_skipped={stats.empty_value} "
            f"invalid_skipped={stats.invalid_value} "
            f"missing_in_db={stats.missing_in_db} unchanged={stats.unchanged} "
            f"{verb}={stats.updated}"
        )
        for wid, old, new in stats.changes:
            print(f"  [CHANGE] {FIX_SHEET_COL}={wid} {field.db_col}: {old!r} → {new!r}")
    return results


def export_frames(frames: Mapping[str, pd.DataFrame], out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    for name, df in frames.items():
        safe = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in name) or "sheet"
        path = out_dir / f"{safe}.csv"
        df.to_csv(path, index=False, encoding="utf-8-sig")
        print(f"[OK] wrote {path} rows={len(df)} cols={len(df.columns)}")


def preview_frames(frames: Mapping[str, pd.DataFrame], preview: int) -> None:
    for name, df in frames.items():
        print(f"=== {name}  shape={df.shape} ===")
        print(df.head(preview).to_string(index=False) if preview > 0 else df.to_string(index=False))
        print()


def build_parser() -> argparse.ArgumentParser:
    up_choices = list(UP_FIELD_MAP.keys())
    parser = argparse.ArgumentParser(
        description="读取钉钉「仓库信息」表格，默认回写 warehouse 白名单全部字段",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "匹配键: 仓库ID → warehouse_id\n"
            "默认更新字段:\n  "
            + ", ".join(up_choices)
            + "\n\n示例:\n"
            "  python app/ding-disk/warehouse_pull.py --workbook-id <ID>\n"
            "  python app/ding-disk/warehouse_pull.py --dry-run\n"
            "  python app/ding-disk/warehouse_pull.py --up-field ops_owner\n"
            "  python app/ding-disk/warehouse_pull.py --id 237\n"
            "  python app/ding-disk/warehouse_pull.py --no-update --preview 20\n"
        ),
    )
    parser.add_argument("--workbook-id", default=WORKBOOK_ID, help="表格文档 ID")
    parser.add_argument("--list-sheets", action="store_true", help="仅列出工作表")
    parser.add_argument("--sheet", default=None, help=f"工作表名称；默认 {DEFAULT_SHEET}")
    parser.add_argument("--all", action="store_true", help="读取全部工作表（需 --no-update）")
    parser.add_argument("--no-header", action="store_true", help="首行不当表头")
    parser.add_argument(
        "--preview",
        type=int,
        default=None,
        help="预览行数；0=全量；-1=不打印。默认：更新时不打印，--no-update 时 10 行",
    )
    parser.add_argument("--out-dir", default=None, help="导出 CSV（utf-8-sig）目录")
    parser.add_argument("--raw", action="store_true", help="--list-sheets 时缩进 JSON")
    parser.add_argument(
        "--id",
        action="append",
        dest="warehouse_ids",
        default=None,
        help="仅更新指定仓库ID（可重复）；默认更新表格中全部行",
    )
    parser.add_argument(
        "--up-field",
        action="append",
        choices=up_choices,
        dest="up_fields",
        default=None,
        help="只更新指定字段（可重复）；默认更新白名单全部字段",
    )
    parser.add_argument("--no-update", action="store_true", help="只读模式：不写库")
    parser.add_argument("--dry-run", action="store_true", help="统计将更新的行，不写库")
    parser.add_argument(
        "--allow-empty",
        action="store_true",
        help="允许用空字符串覆盖库内字段（默认跳过空值）",
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

        do_update = not args.no_update
        if args.all and do_update:
            print("[FAIL] --all 不能与默认更新同时使用，请加 --no-update", file=sys.stderr)
            return 2

        frames = load_frames(
            wb,
            sheet=args.sheet,
            read_all=bool(args.all),
            header=not args.no_header,
        )

        if do_update:
            field_keys = resolve_up_fields(args.up_fields)
            name, df = next(iter(frames.items()))
            mode = "dry-run" if args.dry_run else "write"
            id_hint = f" ids={','.join(args.warehouse_ids)}" if args.warehouse_ids else ""
            print(
                f"[SYNC] sheet={name} fields={','.join(field_keys)}"
                f"{id_hint} mode={mode}"
            )
            update_fields_from_df(
                df,
                field_keys,
                dry_run=bool(args.dry_run),
                allow_empty=bool(args.allow_empty),
                warehouse_ids=args.warehouse_ids,
            )

        if args.out_dir:
            export_frames(frames, Path(args.out_dir))

        preview = args.preview if args.preview is not None else (10 if args.no_update else -1)
        if preview >= 0:
            preview_frames(frames, preview)

        return 0
    except DingDiskError as exc:
        print(f"[FAIL] DingDisk: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:  # noqa: BLE001
        print(f"[FAIL] {type(exc).__name__}: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
