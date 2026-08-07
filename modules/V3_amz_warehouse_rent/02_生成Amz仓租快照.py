"""生成 AMZ 仓租快照（功能同 J2_合并_站点商品ID识别码）。

读取 01 产出的 ``(已完成-1)FBA仓租明细{fba_date}.xlsx``，按站点商品ID识别码汇总后
写入 ``amz_warehouse_rent_snapshot``；同时另存 ``(处理完成)`` Excel 供下游兼容。

用法::

    python modules/V3_amz_warehouse_rent/02_生成Amz仓租快照.py
    python modules/V3_amz_warehouse_rent/02_生成Amz仓租快照.py --month 2026-05
    python modules/V3_amz_warehouse_rent/02_生成Amz仓租快照.py --month 2026-05 --dry-run
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import warnings
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Sequence

import pandas as pd
import pymysql.cursors

# 须在 import config/common 之前：加载项目根到 sys.path
_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
assert _spec.loader is not None
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.style import Color  # noqa: E402
from config.A0_paths import DESKTOP_ROOT  # noqa: E402
from config.A0_set_date import folder_name, shared_date  # noqa: E402
from database.db_connection import get_db_manager  # noqa: E402

# 复用 01 的月份解析（snapshot_date = 月末）
_01_path = Path(__file__).resolve().parent / "01_计算Amz仓租.py"
_01_spec = importlib.util.spec_from_file_location("v3_amz_calc_rent", _01_path)
_01_mod = importlib.util.module_from_spec(_01_spec)
assert _01_spec.loader is not None
_01_spec.loader.exec_module(_01_mod)
resolve_fba_snapshot_date = _01_mod.resolve_fba_snapshot_date

PRODUCT_SKU_TABLE = "product_sku"
RENT_TABLE = "amz_warehouse_rent_snapshot"
_KEY_CHUNK = 200
BATCH_SIZE = 200
DEFAULT_CURRENCY = "EUR"

# 表注释：market_code=站点代码，market_region=区域名称；
# 与 J2 对齐：market_code←站点，market_region←平台。
# （注意：与 platform_shop 的 market_code=平台 / market_region=站点 命名相反）
INSERT_COLS = (
    "snapshot_id",
    "snapshot_date",
    "snapshot_month",
    "product_sku",
    "product_uid",
    "market_code",
    "market_region",
    "rent_fee",
    "currency",
    "remark",
    "is_deleted",
)

UPSERT_SQL = f"""
INSERT INTO `{RENT_TABLE}` (
    {", ".join(f"`{c}`" for c in INSERT_COLS)}
) VALUES (
    {", ".join(["%s"] * len(INSERT_COLS))}
)
ON DUPLICATE KEY UPDATE
    `snapshot_date` = VALUES(`snapshot_date`),
    `snapshot_month` = VALUES(`snapshot_month`),
    `product_uid` = VALUES(`product_uid`),
    `market_region` = VALUES(`market_region`),
    `rent_fee` = VALUES(`rent_fee`),
    `currency` = VALUES(`currency`),
    `remark` = VALUES(`remark`),
    `updated_at` = CURRENT_TIMESTAMP
"""

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


def build_snapshot_id(snapshot_month: str) -> str:
    """``char(32)``：按归属月 ``yyyy-mm`` 生成稳定 MD5。"""
    raw = f"amz_warehouse_rent|{snapshot_month}"
    return hashlib.md5(raw.encode("utf-8")).hexdigest()


def format_snapshot_month(d: date) -> str:
    """归属月保存格式：``yyyy-mm``。"""
    return f"{d.year:04d}-{d.month:02d}"


def input_excel_path(fba_label: str) -> Path:
    return (
        Path(DESKTOP_ROOT)
        / f"{folder_name}{shared_date}"
        / "仓租"
        / "FBA仓租"
        / f"(已完成-1)FBA仓租明细{fba_label}.xlsx"
    )


def _as_text(value: Any, *, max_len: int | None = None) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip()
    if s.lower() in ("nan", "none", "<na>"):
        return ""
    if max_len is not None and len(s) > max_len:
        return s[:max_len]
    return s


def _as_decimal(value: Any) -> Decimal:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return Decimal("0")
    try:
        return Decimal(str(value))
    except (InvalidOperation, ValueError):
        return Decimal("0")


def _fetch_first_sku_by_uid(uids: list[str]) -> dict[str, str]:
    """product_uid → 最新 product_sku（按 id 降序，取 id 最大的一条）。"""
    uids = sorted({str(x).strip() for x in uids if x and str(x).strip()})
    if not uids:
        return {}

    mapping: dict[str, str] = {}
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            for i in range(0, len(uids), _KEY_CHUNK):
                chunk = uids[i : i + _KEY_CHUNK]
                placeholders = ", ".join(["%s"] * len(chunk))
                sql = f"""
                    SELECT product_uid, product_sku
                    FROM `{PRODUCT_SKU_TABLE}`
                    WHERE product_uid IN ({placeholders})
                      AND is_deleted = 0
                      AND product_sku IS NOT NULL
                      AND TRIM(product_sku) <> ''
                    ORDER BY id DESC
                """
                cur.execute(sql, chunk)
                for row in cur.fetchall():
                    uid = str(row.get("product_uid") or "").strip()
                    sku = str(row.get("product_sku") or "").strip()
                    if uid and sku and uid not in mapping:
                        mapping[uid] = sku
    finally:
        conn.close()
    return mapping


def map_product_uid_to_sku(main_df: pd.DataFrame, main_uid: str = "商品ID") -> pd.DataFrame:
    """
    商品ID（product_uid）→ 主产品编码（product_sku）。
    未命中保留原 SKU；原 商品ID 带 -NW 时，映射 SKU 缀回 -NW。
    """
    out = main_df.copy()
    if main_uid not in out.columns:
        raise KeyError(f"主表缺少列 {main_uid!r}，当前列: {list(out.columns)}")
    if "SKU" not in out.columns:
        raise KeyError(f"主表缺少列 'SKU'，当前列: {list(out.columns)}")

    series = out[main_uid].astype(str).str.strip()
    invalid = series.isin(("", "nan", "None", "NaN")) | out[main_uid].isna()
    nw_mask = series.str.endswith("-NW", na=False) & ~invalid
    series_no_nw = series.mask(nw_mask, series.str.replace(r"-NW$", "", regex=True))

    uid_sku_map = _fetch_first_sku_by_uid(series_no_nw[~invalid].tolist())
    print(f"[DB] product_sku 命中 {len(uid_sku_map)} 条 product_uid → 最新 product_sku")

    mapped = series_no_nw.map(uid_sku_map)
    miss = (~invalid) & mapped.isna()
    mapped = mapped.mask(nw_mask & mapped.notna(), mapped.astype(str) + "-NW")

    orig_sku = out["SKU"]
    mapped = mapped.where(mapped.notna(), orig_sku)
    mapped = mapped.mask(invalid, orig_sku)

    out = out.rename(columns={"SKU": "原-SKU"})
    insert_pos = out.columns.get_loc(main_uid) + 1
    out.insert(insert_pos, "SKU", mapped)

    n_miss = int(miss.sum())
    if n_miss:
        preview_cols = [c for c in ("原-SKU", "SKU", "商品ID", "站点", "FBA仓租费") if c in out.columns]
        preview = out.loc[miss, preview_cols].head(10)
        print(
            f"{Color.YELLOW}[检查] 商品ID 有 {n_miss} 行未命中 product_sku"
            f"（已保留原 SKU），请核对：{Color.RESET}"
        )
        print(preview.to_string(index=False))
    return out


def aggregate_rent(df: pd.DataFrame) -> pd.DataFrame:
    """同 J2：按站点商品ID识别码汇总 FBA仓租费。"""
    out = df.rename(columns={"映射站点": "站点", "映射平台": "平台"}).copy()
    if "站点" not in out.columns or "平台" not in out.columns:
        raise KeyError(
            f"缺少 站点/平台（或 映射站点/映射平台），当前列: {list(out.columns)}"
        )

    out = out[out["FBA仓租费"] != 0]
    result = (
        out.groupby("站点商品ID识别码", dropna=False)
        .agg(
            {
                "SKU": "first",
                "商品ID": "first",
                "站点": "first",
                "平台": "first",
                "平台商品ID识别码": "first",
                "FBA仓租费": "sum",
            }
        )
        .reset_index()
    )
    result = map_product_uid_to_sku(result, main_uid="商品ID")
    return result[
        ["SKU", "商品ID", "站点", "平台", "站点商品ID识别码", "平台商品ID识别码", "FBA仓租费"]
    ]


def rows_for_db(
    result_df: pd.DataFrame,
    *,
    snapshot_id: str,
    snapshot_date: date,
    snapshot_month: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for rec in result_df.to_dict("records"):
        sku = _as_text(rec.get("SKU"), max_len=128)
        uid = _as_text(rec.get("商品ID"), max_len=64)
        site = _as_text(rec.get("站点"), max_len=64)
        platform = _as_text(rec.get("平台"), max_len=128)
        fee = _as_decimal(rec.get("FBA仓租费"))
        if not sku or fee == 0:
            continue
        if not site:
            print(f"{Color.YELLOW}[跳过] SKU={sku} 站点为空{Color.RESET}")
            continue
        rows.append(
            {
                "snapshot_id": snapshot_id,
                "snapshot_date": snapshot_date,
                "snapshot_month": snapshot_month,
                "product_sku": sku,
                "product_uid": uid,
                "market_code": site,
                "market_region": platform,
                "rent_fee": fee,
                "currency": DEFAULT_CURRENCY,
                "remark": _as_text(rec.get("站点商品ID识别码"), max_len=255),
                "is_deleted": 0,
            }
        )
    return rows


def delete_snapshot(*, snapshot_id: str, dry_run: bool = False) -> int:
    if dry_run:
        return 0
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(
                f"DELETE FROM `{RENT_TABLE}` WHERE `snapshot_id` = %s AND `is_deleted` = 0",
                (snapshot_id,),
            )
            deleted = int(cur.rowcount or 0)
        conn.commit()
        return deleted
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def upsert_rows(rows: Sequence[dict[str, Any]], *, dry_run: bool = False) -> int:
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
                params = [tuple(r[c] for c in INSERT_COLS) for r in chunk]
                cur.executemany(UPSERT_SQL, params)
                written += len(chunk)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return written


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="按站点商品ID汇总 AMZ 仓租并写入 amz_warehouse_rent_snapshot（同 J2）"
    )
    parser.add_argument(
        "--month",
        default=None,
        help="快照归属月 YYYY-MM（默认按 A0_set_date 的 fba_date；snapshot_date 取月末）",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只计算不写库、不写 Excel",
    )
    parser.add_argument(
        "--no-replace",
        action="store_true",
        help="不清空同 snapshot_id 旧数据，仅 UPSERT",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    snapshot_date, fba_label = resolve_fba_snapshot_date(args.month)
    snapshot_month = format_snapshot_month(snapshot_date)
    snapshot_id = build_snapshot_id(snapshot_month)
    print(
        f"fba_date={fba_label}, snapshot_date={snapshot_date}, "
        f"snapshot_month={snapshot_month}, snapshot_id={snapshot_id}"
    )

    excel_path = input_excel_path(fba_label)
    if not excel_path.is_file():
        raise FileNotFoundError(
            f"未找到 01 产出文件：{excel_path}\n"
            f"请先运行：python modules/V3_amz_warehouse_rent/01_计算Amz仓租.py"
            + (f" --month {args.month}" if args.month else "")
        )

    amazon_df = pd.read_excel(excel_path)
    print(f"[Excel] 读取 {len(amazon_df)} 行：{excel_path}")

    result_df = aggregate_rent(amazon_df)
    print(f"[汇总] 站点商品ID识别码 {len(result_df)} 行")

    db_rows = rows_for_db(
        result_df,
        snapshot_id=snapshot_id,
        snapshot_date=snapshot_date,
        snapshot_month=snapshot_month,
    )
    print(f"[DB] 待写入 {len(db_rows)} 行 → `{RENT_TABLE}`")

    deleted = 0
    if not args.no_replace:
        deleted = delete_snapshot(snapshot_id=snapshot_id, dry_run=args.dry_run)
        print(f"[DB] 删除旧快照 {deleted} 行" + (" (dry-run)" if args.dry_run else ""))

    written = upsert_rows(db_rows, dry_run=args.dry_run)
    print(
        f"[DB] 写入 {written} 行"
        + (" (dry-run)" if args.dry_run else f" → `{RENT_TABLE}`")
    )

    if not args.dry_run:
        output_path = excel_path.with_name(
            excel_path.name.replace("已完成-1", "处理完成")
        )
        result_df.to_excel(output_path, index=False)
        print(f"处理完成，文件另存为：{output_path}")
    else:
        print("[dry-run] 跳过 Excel 写出")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
