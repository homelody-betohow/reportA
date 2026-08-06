"""增量同步易仓 WMS 产品列表到 ``product_sku``，并将新增 SKU 推至钉钉表格。

仅同步 ``product_status = 1``（可用）的产品。
默认只拉取近 7 天有更新的产品（``--update-from``）。

流程::
    1. 分页调用 ``getWmsProductList``（``product_status=1``，默认近 7 天更新）
    2. 按 ``product_sku`` 检查本地表；不存在则 INSERT（已存在跳过）
    3. 对本次新增的 SKU 调用 ``product_sku_push`` 追加到钉钉「产品信息」表

用法（项目根目录）::

    python app/eccang/syncProductList.py
    python app/eccang/syncProductList.py --dry-run
    python app/eccang/syncProductList.py --update-from "2026-07-01 00:00:00"
    python app/eccang/syncProductList.py --update-from "" --update-to ""
    python app/eccang/syncProductList.py --limit-pages 1
    python app/eccang/syncProductList.py --product-sku HL02001
    python app/eccang/syncProductList.py --skip-ding
"""

from __future__ import annotations

import argparse
import importlib.util
import sys
import time
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Optional, Sequence

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

from api.eccang.exceptions import EccangApiError, EccangConfigError  # noqa: E402
from api.eccang.request.getWmsProductList import get_wms_product_list  # noqa: E402
from database.db_connection import get_db_manager  # noqa: E402

TABLE = "product_sku"
DEFAULT_PAGE_SIZE = 500
BATCH_SIZE = 200
SOURCE_TYPE = "EccangAPI"
# 仅同步可用产品：0不可用 / 1可用 / 2开发产品
PRODUCT_STATUS_ACTIVE = 1
DEFAULT_UPDATE_FROM_DAYS = 7
DATETIME_FMT = "%Y-%m-%d %H:%M:%S"

# INSERT 列（不含 id / created_at / updated_at）
INSERT_COLS: tuple[str, ...] = (
    "product_sku",
    "product_uid",
    "product_name_cn",
    "product_name_en",
    "category_lv1",
    "category_lv2",
    "category_lv3",
    "category_code",
    "ean_code",
    "supplier_abbr",
    "supplier_name",
    "product_color",
    "product_img",
    "declare_price_usd",
    "declare_name_cn",
    "declare_name_en",
    "hs_code",
    "purchase_price",
    "unit_weight_g",
    "outer_box_l_cm",
    "outer_box_w_cm",
    "outer_box_h_cm",
    "source_type",
    "is_deleted",
)

INSERT_SQL = f"""
INSERT INTO `{TABLE}` (
    {", ".join(f"`{c}`" for c in INSERT_COLS)}
) VALUES (
    {", ".join(["%s"] * len(INSERT_COLS))}
)
"""


def _as_text(value: Any, *, max_len: int | None = None, default: str = "") -> str:
    if value is None:
        return default
    text = str(value).strip()
    if not text or text.lower() in {"none", "null"}:
        return default
    if max_len is not None and len(text) > max_len:
        return text[:max_len]
    return text


def _as_decimal(value: Any, *, quant: str | None = "0.0001") -> Decimal | None:
    if value is None or value == "":
        return None
    try:
        d = Decimal(str(value).strip().replace(",", ""))
    except (InvalidOperation, ValueError):
        return None
    if quant is None:
        return d
    return d.quantize(Decimal(quant), rounding=ROUND_HALF_UP)


def _pick(record: dict[str, Any], *keys: str) -> Any:
    for key in keys:
        if key in record and record[key] not in (None, ""):
            return record[key]
    return None


def _weight_to_grams(value: Any) -> Decimal | None:
    """易仓 ``product_weight`` 多为 kg；转为 ``unit_weight_g``。"""
    d = _as_decimal(value, quant=None)
    if d is None:
        return None
    # 常见：kg 小数（如 0.571）；已是克的大整数原样保留
    if d == 0:
        return Decimal("0.00")
    if abs(d) < 100 and d != d.to_integral_value():
        d = d * Decimal("1000")
    return d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def _product_img(record: dict[str, Any]) -> str:
    images = _as_text(_pick(record, "product_images"), max_len=255)
    if images.startswith("http"):
        return images
    main = _as_text(_pick(record, "main_img"), max_len=255)
    if main.startswith("http"):
        return main
    return images or main


def _category_code(record: dict[str, Any]) -> str | None:
    parts = [
        _as_text(record.get(k))
        for k in (
            "procut_category_code1",
            "procut_category_code2",
            "procut_category_code3",
        )
    ]
    joined = "/".join(p for p in parts if p)
    return joined[:16] if joined else None


def map_api_record(record: dict[str, Any]) -> dict[str, Any] | None:
    """API 单条 → ``product_sku`` 行；无 SKU 返回 None。"""
    sku = _as_text(_pick(record, "product_sku"), max_len=64)
    if not sku:
        return None

    name_cn = _as_text(_pick(record, "product_title"), max_len=255)
    name_en = _as_text(_pick(record, "product_title_en"), max_len=255)
    if not name_cn and name_en:
        name_cn = name_en

    return {
        "product_sku": sku,
        "product_uid": _as_text(_pick(record, "product_spu"), max_len=64) or None,
        "product_name_cn": name_cn,
        "product_name_en": name_en,
        "category_lv1": _as_text(record.get("procut_category_name1"), max_len=64) or None,
        "category_lv2": _as_text(record.get("procut_category_name2"), max_len=64) or None,
        "category_lv3": _as_text(record.get("procut_category_name3"), max_len=64) or None,
        "category_code": _category_code(record),
        "ean_code": _as_text(_pick(record, "ean_code", "warehouse_barcode"), max_len=100),
        "supplier_abbr": _as_text(record.get("default_supplier_code"), max_len=60),
        "supplier_name": _as_text(record.get("default_supplier_name"), max_len=128) or None,
        "product_color": _as_text(record.get("product_color_name"), max_len=32) or None,
        "product_img": _product_img(record),
        "declare_price_usd": _as_decimal(record.get("product_declared_value"), quant="0.01")
        or Decimal("0.00"),
        "declare_name_cn": _as_text(record.get("pd_oversea_type_cn"), max_len=200),
        "declare_name_en": _as_text(record.get("pd_oversea_type_en"), max_len=200),
        "hs_code": _as_text(record.get("hs_code"), max_len=32),
        "purchase_price": _as_decimal(record.get("sp_unit_price"), quant="0.0001"),
        "unit_weight_g": _weight_to_grams(record.get("product_weight")),
        "outer_box_l_cm": _as_decimal(record.get("product_length"), quant="0.01"),
        "outer_box_w_cm": _as_decimal(record.get("product_width"), quant="0.01"),
        "outer_box_h_cm": _as_decimal(record.get("product_height"), quant="0.01"),
        "source_type": SOURCE_TYPE,
        "is_deleted": 0,
    }


def extract_records(resp: dict[str, Any]) -> tuple[list[dict[str, Any]], int | None, int | None]:
    """解析 getWmsProductList 响应 → (records, total, pages)。"""
    data = resp.get("data")
    if not isinstance(data, dict):
        return [], None, None
    items = data.get("data")
    if isinstance(items, dict):
        items = [items]
    if not isinstance(items, list):
        items = []
    records = [r for r in items if isinstance(r, dict)]

    total: int | None = None
    raw_total = data.get("total")
    if raw_total is not None and str(raw_total).strip() != "":
        try:
            total = int(raw_total)
        except (TypeError, ValueError):
            total = None

    page_size = data.get("page_size")
    pages: int | None = None
    if total is not None and page_size:
        try:
            size = int(page_size)
            if size > 0:
                pages = (total + size - 1) // size
        except (TypeError, ValueError):
            pages = None
    return records, total, pages


def default_update_from(*, days: int = DEFAULT_UPDATE_FROM_DAYS) -> str:
    """默认 ``product_update_time_from``：N 天前 00:00:00。"""
    return (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d 00:00:00")


def normalize_datetime_arg(value: str | None) -> str | None:
    """空串视为未设置；非空则校验格式。"""
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        datetime.strptime(text, DATETIME_FMT)
    except ValueError as exc:
        raise ValueError(
            f"时间格式应为 {DATETIME_FMT}，收到：{value!r}"
        ) from exc
    return text


def _is_active_status(record: dict[str, Any]) -> bool:
    """``product_status == 1``（可用）。"""
    raw = record.get("product_status")
    if raw is None or raw == "":
        return False
    try:
        return int(raw) == PRODUCT_STATUS_ACTIVE
    except (TypeError, ValueError):
        return str(raw).strip() == str(PRODUCT_STATUS_ACTIVE)


def fetch_all_products(
    *,
    page_size: int = DEFAULT_PAGE_SIZE,
    product_sku: str | None = None,
    product_update_time_from: str | None = None,
    product_update_time_to: str | None = None,
    limit_pages: int | None = None,
    sleep_sec: float = 0.15,
) -> list[dict[str, Any]]:
    """分页拉取产品列表（仅 ``product_status=1``）。"""
    page = 1
    all_records: list[dict[str, Any]] = []
    sku = (product_sku or "").strip() or None
    status_filter = {"product_status": PRODUCT_STATUS_ACTIVE}
    # 精确查 SKU 时不套更新时间窗，避免误过滤
    update_from = None if sku else product_update_time_from
    update_to = None if sku else product_update_time_to

    while True:
        if limit_pages is not None and page > limit_pages:
            break
        resp = get_wms_product_list(
            page=page,
            page_size=page_size if not sku else min(page_size, 20),
            product_sku=sku,
            product_update_time_from=update_from,
            product_update_time_to=update_to,
            extra=status_filter,
        )
        batch, total, pages = extract_records(resp)
        # 精确查 SKU 时 API 可能仍返回非可用；本地再过滤一次
        active = [r for r in batch if _is_active_status(r)]
        skipped = len(batch) - len(active)
        all_records.extend(active)
        print(
            f"[API] page={page}/{pages or '?'} got={len(batch)} "
            f"active={len(active)} skip_status={skipped} "
            f"accum={len(all_records)} total={total}",
            file=sys.stderr,
        )
        if not batch:
            break
        if sku:
            break
        if pages is not None and page >= pages:
            break
        if total is not None and len(all_records) >= total:
            break
        if len(batch) < page_size and pages is None:
            break
        page += 1
        if sleep_sec > 0:
            time.sleep(sleep_sec)
    return all_records


def fetch_existing_skus(skus: Sequence[str]) -> set[str]:
    """已存在的 product_sku（含软删，避免撞唯一键）。"""
    if not skus:
        return set()
    result: set[str] = set()
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            for i in range(0, len(skus), BATCH_SIZE):
                chunk = list(skus[i : i + BATCH_SIZE])
                placeholders = ",".join(["%s"] * len(chunk))
                cur.execute(
                    f"SELECT product_sku FROM `{TABLE}` "
                    f"WHERE product_sku IN ({placeholders})",
                    chunk,
                )
                for row in cur.fetchall():
                    sku = _as_text(row.get("product_sku"))
                    if sku:
                        result.add(sku)
    finally:
        conn.close()
    return result


def _row_params(row: dict[str, Any]) -> tuple[Any, ...]:
    return tuple(row.get(c) for c in INSERT_COLS)


def insert_rows(rows: Sequence[dict[str, Any]], *, dry_run: bool = False) -> int:
    if not rows:
        return 0
    if dry_run:
        return len(rows)
    db = get_db_manager()
    conn = db.get_connection()
    written = 0
    try:
        with conn.cursor() as cur:
            for i in range(0, len(rows), BATCH_SIZE):
                chunk = rows[i : i + BATCH_SIZE]
                cur.executemany(INSERT_SQL, [_row_params(r) for r in chunk])
                written += len(chunk)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return written


def _load_product_sku_push():
    """动态加载 ``app/ding-disk/product_sku_push.py``。"""
    push_path = _PROJECT_ROOT / "app" / "ding-disk" / "product_sku_push.py"
    if not push_path.is_file():
        raise FileNotFoundError(f"未找到 product_sku_push: {push_path}")
    mod_name = "product_sku_push_dyn"
    if mod_name in sys.modules:
        return sys.modules[mod_name]
    spec = importlib.util.spec_from_file_location(mod_name, push_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载: {push_path}")
    mod = importlib.util.module_from_spec(spec)
    # dataclass 依赖 sys.modules[cls.__module__]
    sys.modules[mod_name] = mod
    spec.loader.exec_module(mod)
    return mod


def push_new_skus_to_ding(
    skus: Sequence[str],
    *,
    dry_run: bool = False,
    preview_rows: Sequence[dict[str, Any]] | None = None,
) -> int:
    """调用 ``product_sku_push.sync_new_rows_to_sheet`` 追加新增 SKU。

    ``dry_run`` 且尚未写库时，可传 ``preview_rows``（待插入映射行）做追加预览。
    """
    if not skus and not preview_rows:
        return 0

    mod = _load_product_sku_push()
    workbook_id = getattr(mod, "WORKBOOK_ID", "")
    sheet_name = getattr(mod, "DEFAULT_SHEET", "Sheet1")
    sku_col = getattr(mod, "SKU_SHEET_COL", "SKU")

    print(
        f"[DING] append {len(skus) or len(preview_rows or [])} SKU → "
        f"workbook={workbook_id} sheet={sheet_name}"
        f"{' (dry-run)' if dry_run else ''}",
        file=sys.stderr,
    )

    wb = mod.Workbook(workbook_id)
    df = wb.read_sheet(sheet_name, header=True)

    if dry_run and preview_rows:
        existing = mod.sheet_sku_set(df)
        new_rows = [
            r
            for r in preview_rows
            if _as_text(r.get("product_sku"))
            and _as_text(r.get("product_sku")) not in existing
        ]
        append_df = mod.build_append_dataframe(list(df.columns), new_rows)
        print(
            f"[APPEND] dry-run preview to_append={len(new_rows)} "
            f"sheet_skus={len(existing)} shape={append_df.shape}",
            file=sys.stderr,
        )
        if new_rows:
            print(append_df.head(min(5, len(append_df))).to_string(index=False))
        return 0

    stats = mod.sync_new_rows_to_sheet(
        wb,
        sheet_name,
        df,
        dry_run=dry_run,
        skus=list(skus),
        preview=-1,
    )
    if stats.missing_sku_col:
        print(
            f"[FAIL] 表格缺少列: [{sku_col}]；实际列={list(df.columns)}",
            file=sys.stderr,
        )
        return 2

    verb = "would_append" if dry_run else "appended"
    print(
        f"[APPEND] product_sku → [{sku_col}]  "
        f"db_rows={stats.db_rows} sheet_skus={stats.sheet_skus} "
        f"already_in_sheet={stats.already_in_sheet} "
        f"to_append={stats.to_append} "
        f"{verb}={stats.appended if not dry_run else stats.to_append}",
        file=sys.stderr,
    )
    return 0


def sync_product_list(
    *,
    page_size: int = DEFAULT_PAGE_SIZE,
    product_sku: str | None = None,
    product_update_time_from: str | None = None,
    product_update_time_to: str | None = None,
    limit_pages: int | None = None,
    dry_run: bool = False,
    skip_ding: bool = False,
    sleep_sec: float = 0.15,
) -> dict[str, int]:
    """同步入口。返回统计字典。"""
    window = []
    if product_sku:
        window.append(f"sku={product_sku}")
    else:
        window.append(
            f"update_from={product_update_time_from or '(none)'} "
            f"update_to={product_update_time_to or '(none)'}"
        )
    print(
        f"[SYNC] 拉取易仓 WMS 产品列表（product_status={PRODUCT_STATUS_ACTIVE}；"
        f"{'; '.join(window)}）…",
        file=sys.stderr,
    )
    raw_records = fetch_all_products(
        page_size=page_size,
        product_sku=product_sku,
        product_update_time_from=product_update_time_from,
        product_update_time_to=product_update_time_to,
        limit_pages=limit_pages,
        sleep_sec=sleep_sec,
    )

    mapped: list[dict[str, Any]] = []
    seen: set[str] = set()
    skip_empty = 0
    for rec in raw_records:
        row = map_api_record(rec)
        if row is None:
            skip_empty += 1
            continue
        sku = row["product_sku"]
        if sku in seen:
            # 同 SKU 保留最后一次
            mapped = [r for r in mapped if r["product_sku"] != sku]
        seen.add(sku)
        mapped.append(row)

    print(
        f"[MAP] api={len(raw_records)} mapped={len(mapped)} empty_sku={skip_empty}",
        file=sys.stderr,
    )

    existing = fetch_existing_skus([r["product_sku"] for r in mapped])
    to_insert = [r for r in mapped if r["product_sku"] not in existing]
    skipped = len(mapped) - len(to_insert)

    print(
        f"[DB] existing={len(existing)} to_insert={len(to_insert)} skipped={skipped}"
        f"{' (dry-run)' if dry_run else ''}",
        file=sys.stderr,
    )
    if to_insert:
        preview = ", ".join(r["product_sku"] for r in to_insert[:10])
        more = "" if len(to_insert) <= 10 else f" …(+{len(to_insert) - 10})"
        print(f"[NEW] {preview}{more}", file=sys.stderr)

    written = insert_rows(to_insert, dry_run=dry_run)
    print(f"[WRITE] inserted={written}", file=sys.stderr)

    ding_rc = 0
    new_skus = [r["product_sku"] for r in to_insert]
    if new_skus and not skip_ding:
        ding_rc = push_new_skus_to_ding(
            new_skus,
            dry_run=dry_run,
            preview_rows=to_insert if dry_run else None,
        )
    elif skip_ding and new_skus:
        print(f"[DING] skipped ({len(new_skus)} new SKUs)", file=sys.stderr)
    else:
        print("[DING] no new SKU, skip", file=sys.stderr)

    return {
        "api_rows": len(raw_records),
        "mapped": len(mapped),
        "existing": len(existing),
        "to_insert": len(to_insert),
        "written": written,
        "ding_rc": ding_rc,
    }


def build_parser() -> argparse.ArgumentParser:
    return argparse.ArgumentParser(
        description=(
            "同步易仓产品到 product_sku（默认近 7 天有更新），"
            "并将新增 SKU 推至钉钉表格"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例:\n"
            "  python app/eccang/syncProductList.py\n"
            "  python app/eccang/syncProductList.py --dry-run\n"
            "  python app/eccang/syncProductList.py "
            '--update-from "2026-07-01 00:00:00"\n'
            "  python app/eccang/syncProductList.py --update-from \"\"\n"
            "  python app/eccang/syncProductList.py --product-sku HL02001\n"
            "  python app/eccang/syncProductList.py --limit-pages 1 --skip-ding\n"
        ),
    )


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    parser.add_argument(
        "--page-size",
        type=int,
        default=DEFAULT_PAGE_SIZE,
        help=f"每页条数，最大 1000；默认 {DEFAULT_PAGE_SIZE}",
    )
    parser.add_argument(
        "--product-sku",
        dest="product_sku",
        default=None,
        help="仅同步指定 SKU（精确匹配；忽略更新时间窗）",
    )
    parser.add_argument(
        "--update-from",
        dest="product_update_time_from",
        default=default_update_from(),
        metavar="DATETIME",
        help=(
            f"产品更新时间起（{DATETIME_FMT}），对应 product_update_time_from；"
            f"默认 {DEFAULT_UPDATE_FROM_DAYS} 天前 00:00:00；传空串表示不限制"
        ),
    )
    parser.add_argument(
        "--update-to",
        dest="product_update_time_to",
        default=None,
        metavar="DATETIME",
        help=f"产品更新时间止（{DATETIME_FMT}），对应 product_update_time_to；默认不限制",
    )
    parser.add_argument(
        "--limit-pages",
        type=int,
        default=None,
        help="最多拉取页数（调试用）",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只统计，不写库、不写钉钉",
    )
    parser.add_argument(
        "--skip-ding",
        action="store_true",
        help="跳过钉钉推送（仅写库）",
    )
    parser.add_argument(
        "--sleep",
        type=float,
        default=0.15,
        help="分页请求间隔秒数，默认 0.15",
    )
    args = parser.parse_args(argv)

    try:
        update_from = normalize_datetime_arg(args.product_update_time_from)
        update_to = normalize_datetime_arg(args.product_update_time_to)
        stats = sync_product_list(
            page_size=min(max(1, int(args.page_size)), 1000),
            product_sku=args.product_sku,
            product_update_time_from=update_from,
            product_update_time_to=update_to,
            limit_pages=args.limit_pages,
            dry_run=bool(args.dry_run),
            skip_ding=bool(args.skip_ding),
            sleep_sec=float(args.sleep),
        )
    except ValueError as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 2
    except EccangConfigError as exc:
        print(f"[FAIL] 配置错误：{exc}", file=sys.stderr)
        return 2
    except EccangApiError as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # noqa: BLE001
        print(f"[FAIL] {type(exc).__name__}: {exc}", file=sys.stderr)
        return 2

    print(
        f"[DONE] api={stats['api_rows']} mapped={stats['mapped']} "
        f"insert={stats['to_insert']} written={stats['written']} "
        f"ding_rc={stats['ding_rc']}"
    )
    return 0 if stats["ding_rc"] == 0 else int(stats["ding_rc"])


if __name__ == "__main__":
    raise SystemExit(main())
