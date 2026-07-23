"""读取钉钉在线表格「产品信息」，并回写 ``product_sku`` 白名单字段。

表格读写委托 ``api.ding_disk.workbook.Workbook``；本脚本只负责字段映射与 DB 更新。

文档 ID（workbookId / nodeId）::
    Obva6QBXJwjBxoE2sM62MrzGVn4qY5Pr

用法（项目根目录）::

    # 默认：读 Sheet1，更新白名单全部字段
    python app/ding-disk/product_sku_pull.py
    python app/ding-disk/product_sku_pull.py --dry-run

    # 只更新部分字段 / 指定 SKU
    python app/ding-disk/product_sku_pull.py --up-field supplier_abbr
    python app/ding-disk/product_sku_pull.py --sku XPM0213
    python app/ding-disk/product_sku_pull.py --sku XPM0213 --up-field supplier_abbr --dry-run

    # 只读（不写库）
    python app/ding-disk/product_sku_pull.py --no-update --preview 20
    python app/ding-disk/product_sku_pull.py --list-sheets
    python app/ding-disk/product_sku_pull.py --all --no-update --out-dir ./runtime/local/ding_product
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
    filter_by_column,
    kv_pairs_from_df,
    normalize_cell,
)

# 钉钉表格文档 ID（知识库 nodeId / dentryUuid）
WORKBOOK_ID = "Obva6QBXJwjBxoE2sM62MrzGVn4qY5Pr"
DEFAULT_SHEET = "Sheet1"
TABLE = "product_sku"
FIX_SHEET_COL = "SKU"
BATCH_SIZE = 500


@dataclass(frozen=True)
class FieldMap:
    """字段名 → 钉钉列 → product_sku 列。"""

    sheet_col: str
    db_col: str


# 白名单：仅允许更新这些字段；默认全部更新
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
    "ops_model": FieldMap(sheet_col="运营模式", db_col="ops_model"),
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


def fetch_existing_values(skus: Sequence[str], db_col: str) -> Dict[str, str]:
    """批量读取 product_sku 当前字段值。"""
    if not skus:
        return {}
    result: Dict[str, str] = {}
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            for i in range(0, len(skus), BATCH_SIZE):
                chunk = list(skus[i : i + BATCH_SIZE])
                placeholders = ",".join(["%s"] * len(chunk))
                cur.execute(
                    f"SELECT product_sku, `{db_col}` AS v FROM `{TABLE}` "
                    f"WHERE product_sku IN ({placeholders}) AND is_deleted = 0",
                    chunk,
                )
                for row in cur.fetchall():
                    result[str(row["product_sku"])] = normalize_cell(row.get("v"))
    finally:
        conn.close()
    return result


def apply_field_updates(
    pairs: Mapping[str, str],
    *,
    db_col: str,
    dry_run: bool = False,
) -> UpdateStats:
    """按 SKU 更新 ``product_sku.<db_col>``；仅写入有变化的行。"""
    stats = UpdateStats(unique_skus=len(pairs))
    if not pairs:
        return stats

    existing = fetch_existing_values(list(pairs.keys()), db_col)
    to_update: List[Tuple[str, str]] = []  # (value, sku)
    for sku, value in pairs.items():
        if sku not in existing:
            stats.missing_in_db += 1
            continue
        if existing[sku] == value:
            stats.unchanged += 1
            continue
        to_update.append((value, sku))

    stats.updated = len(to_update)
    if dry_run or not to_update:
        return stats

    sql = f"UPDATE `{TABLE}` SET `{db_col}` = %s WHERE product_sku = %s AND is_deleted = 0"
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

    pairs, empty_skipped = kv_pairs_from_df(
        df,
        FIX_SHEET_COL,
        field.sheet_col,
        allow_empty=allow_empty,
    )
    stats = apply_field_updates(pairs, db_col=field.db_col, dry_run=dry_run)
    stats.sheet_rows = len(df)
    stats.empty_value = empty_skipped
    return stats


def update_fields_from_df(
    df: pd.DataFrame,
    field_keys: Sequence[str],
    *,
    dry_run: bool = False,
    allow_empty: bool = False,
    skus: Optional[Sequence[str]] = None,
) -> Dict[str, UpdateStats]:
    """依次更新白名单字段；缺列跳过。``skus`` 非空时仅更新这些 SKU。"""
    if FIX_SHEET_COL not in df.columns:
        raise KeyError(f"表格缺少列: [{FIX_SHEET_COL}]；实际列={list(df.columns)}")

    work = df
    if skus:
        work, missing = filter_by_column(df, FIX_SHEET_COL, skus)
        if missing:
            print(f"[WARN] 表格中未找到 SKU: {', '.join(missing)}", file=sys.stderr)
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
        verb = "would_update" if dry_run else "updated"
        print(
            f"[UPDATE] [{field.sheet_col}] → {TABLE}.{field.db_col}  "
            f"unique={stats.unique_skus} empty_skipped={stats.empty_value} "
            f"missing_in_db={stats.missing_in_db} unchanged={stats.unchanged} "
            f"{verb}={stats.updated}"
        )
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
        description="读取钉钉「产品信息」表格，默认回写 product_sku 白名单全部字段",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "默认更新字段:\n  "
            + ", ".join(up_choices)
            + "\n\n示例:\n"
            "  python app/ding-disk/productInfo.py\n"
            "  python app/ding-disk/productInfo.py --dry-run\n"
            "  python app/ding-disk/productInfo.py --up-field supplier_abbr\n"
            "  python app/ding-disk/productInfo.py --sku XPM0213\n"
            "  python app/ding-disk/productInfo.py --no-update --preview 20\n"
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
        "--sku",
        action="append",
        dest="skus",
        default=None,
        help="仅更新指定 SKU（可重复）；默认更新表格中全部 SKU",
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
            sku_hint = f" skus={','.join(args.skus)}" if args.skus else ""
            print(
                f"[SYNC] sheet={name} fields={','.join(field_keys)}"
                f"{sku_hint} mode={mode}"
            )
            update_fields_from_df(
                df,
                field_keys,
                dry_run=bool(args.dry_run),
                allow_empty=bool(args.allow_empty),
                skus=args.skus,
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
