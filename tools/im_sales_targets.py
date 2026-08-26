"""
将 OKR 月目标拆解 Excel 导入 snapshot_sales_targets。

读取 tools/OKR月目标拆分.py 生成的：
  {EXCEL_DIR}\\{yyyy-mm}月目标拆解及跟进.xlsx
默认 sheet：ALL（含「来源Sheet」）。

幂等：按 uk_sst_unique (account_code, product_sku, market_code, target_month) UPSERT；
导入前将该月未删行软删除，写入时清空 deleted_at，使 Excel 中已移除的行保留软删标记。

用法：
  python tools/im_sales_targets.py
  python tools/im_sales_targets.py --month 2026-08
  python tools/im_sales_targets.py --file "F:\\月目标拆解及跟进\\2026-08月目标拆解及跟进.xlsx"
  python tools/im_sales_targets.py --month 2026-08 --dry-run --preview 5
"""

from __future__ import annotations

import argparse
import importlib.util
import re
import sys
from datetime import date
from pathlib import Path
from typing import Any, Iterable, Sequence

import pandas as pd

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

from database.db_connection import get_db_manager  # noqa: E402

# 复用 OKR 脚本的目录 / 月份约定
_okr_path = Path(__file__).resolve().parent / "OKR月目标拆分.py"
_okr_spec = importlib.util.spec_from_file_location("okr_month_split", _okr_path)
_okr = importlib.util.module_from_spec(_okr_spec)
assert _okr_spec.loader is not None
_okr_spec.loader.exec_module(_okr)

TABLE = "snapshot_sales_targets"
DEFAULT_SHEET = "ALL"
BATCH_SIZE = 500

# Excel 列名（或候选）→ 库字段；历史均值列可能带滚动区间前缀，见 HIST_SUFFIX_MAP
COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    "market_code": ("平台",),
    "account_code": ("账号",),
    "product_sku": ("SKU",),
    "product_uid": ("商品ID",),
    "accounting_sku": ("核算SKU",),
    "identify_sku": ("识别SKU",),
    "identify_sku_uid": ("识别商品ID",),
    "category": ("品类",),
    "product_status": ("产品状态",),
    "ops_owner": ("负责人",),
    "dispatch_warehouse": ("发货仓库",),
    "is_transfer": ("是否调拨",),
    "hist_avg_aov": ("平均客单价",),
    "target_rma_rate": ("RMA占比目标",),
    "target_ad_rate": ("广告占比目标",),
    "target_review_rate": ("测评占比目标",),
    "target_aov": ("客单价目标",),
    "target_sales_qty": ("预估销量",),
    "target_platform_sales_amount": ("平台销售额目标",),
    "target_sales_amount": ("销售额目标",),
    "target_gross_profit": ("毛利额目标",),
    "target_operating_gross_profit": ("经营毛利额目标",),
    "target_gross_margin_rate": ("总目标毛利率",),
    "stock_on_hand_qty": ("在库库存",),
    "stock_in_transit_qty": ("在途库存",),
    "stock_unfulfilled_qty": ("未交库存",),
    "stock_total_qty": ("总库存",),
    "remark": ("备注",),
    "target_platform_fee": ("平台费目标",),
    "target_sales_tax": ("销售税目标",),
    "target_withdrawal_fee": ("提现费", "提现费目标"),
    "target_purchase_cost": ("采购成本目标",),
    "target_purchase_cost_ops": ("采购成本目标（经营））", "采购成本目标（经营）"),
    "target_first_leg_tariff": ("头程关税目标",),
    "target_last_mile_fee": ("尾程费目标",),
    "target_warehouse_rent": ("仓租目标",),
    "target_other_allocated_fee": ("其他分摊目标", "其他分摊费用目标"),
    "target_seckill_fee": ("秒杀花费目标",),
    "target_ad_fee": ("广告花费目标",),
    "target_review_fee": ("测评目标", "测评花费目标"),
    "target_return_qty": ("退货量目标",),
    "source_sheet": ("来源Sheet",),
}

# 带区间前缀的历史列：按「列名后缀」匹配（取第一个命中）
HIST_SUFFIX_MAP: dict[str, tuple[str, ...]] = {
    "hist_avg_sales_qty": ("平均销量",),
    "hist_avg_gross_margin_rate": ("平均毛利率",),
    "hist_avg_rma_rate": ("平均RMA",),
    "hist_avg_ad_rate": ("平均广告占比",),
    "hist_avg_review_rate": ("平均测评占比",),
}

RATE_COLS = (
    "hist_avg_gross_margin_rate",
    "hist_avg_rma_rate",
    "hist_avg_ad_rate",
    "hist_avg_review_rate",
    "target_rma_rate",
    "target_ad_rate",
    "target_review_rate",
    "target_gross_margin_rate",
)
DECIMAL4_COLS = (
    "hist_avg_aov",
    "hist_avg_sales_qty",
    "target_aov",
    "target_sales_qty",
)
DECIMAL6_COLS = (
    "target_platform_sales_amount",
    "target_sales_amount",
    "target_gross_profit",
    "target_operating_gross_profit",
    "target_platform_fee",
    "target_sales_tax",
    "target_withdrawal_fee",
    "target_purchase_cost",
    "target_purchase_cost_ops",
    "target_first_leg_tariff",
    "target_last_mile_fee",
    "target_warehouse_rent",
    "target_other_allocated_fee",
    "target_seckill_fee",
    "target_ad_fee",
    "target_review_fee",
)
INT_COLS = (
    "stock_on_hand_qty",
    "stock_in_transit_qty",
    "stock_unfulfilled_qty",
    "stock_total_qty",
    "target_return_qty",
)

UPSERT_COLS = (
    "target_month",
    "market_code",
    "account_code",
    "product_sku",
    "product_uid",
    "accounting_sku",
    "identify_sku",
    "identify_sku_uid",
    "category",
    "product_status",
    "ops_owner",
    "dispatch_warehouse",
    "is_transfer",
    "hist_avg_aov",
    "hist_avg_sales_qty",
    "hist_avg_gross_margin_rate",
    "hist_avg_rma_rate",
    "hist_avg_ad_rate",
    "hist_avg_review_rate",
    "target_rma_rate",
    "target_ad_rate",
    "target_review_rate",
    "target_aov",
    "target_sales_qty",
    "target_platform_sales_amount",
    "target_sales_amount",
    "target_gross_profit",
    "target_operating_gross_profit",
    "target_gross_margin_rate",
    "stock_on_hand_qty",
    "stock_in_transit_qty",
    "stock_unfulfilled_qty",
    "stock_total_qty",
    "target_platform_fee",
    "target_sales_tax",
    "target_withdrawal_fee",
    "target_purchase_cost",
    "target_purchase_cost_ops",
    "target_first_leg_tariff",
    "target_last_mile_fee",
    "target_warehouse_rent",
    "target_other_allocated_fee",
    "target_seckill_fee",
    "target_ad_fee",
    "target_review_fee",
    "target_return_qty",
    "remark",
    "source_sheet",
    "deleted_at",
)

# 唯一键列不在 UPDATE 中改写
_UK_COLS = frozenset({"account_code", "product_sku", "market_code", "target_month"})

UPSERT_SQL = f"""
INSERT INTO `{TABLE}` (
    {", ".join(f"`{c}`" for c in UPSERT_COLS)}
) VALUES (
    {", ".join(["%s"] * len(UPSERT_COLS))}
)
ON DUPLICATE KEY UPDATE
    {", ".join(f"`{c}` = VALUES(`{c}`)" for c in UPSERT_COLS if c not in _UK_COLS)},
    `updated_at` = CURRENT_TIMESTAMP
"""

SOFT_DELETE_SQL = f"""
UPDATE `{TABLE}`
SET `deleted_at` = CURRENT_TIMESTAMP
WHERE `target_month` = %s
  AND `deleted_at` IS NULL
"""


def _norm_header(name: Any) -> str:
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return ""
    return str(name).strip()


def _as_text(value: Any, *, max_len: int | None = None) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip()
    if s.lower() in ("nan", "none", "<na>"):
        return ""
    if max_len is not None and len(s) > max_len:
        return s[:max_len]
    return s


def _as_float(value: Any, *, ndigits: int) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, str):
        s = value.strip().replace(",", "").replace("%", "")
        if not s or s.lower() in ("nan", "none", "-"):
            return 0.0
        try:
            return round(float(s), ndigits)
        except ValueError:
            return 0.0
    try:
        return round(float(value), ndigits)
    except (TypeError, ValueError):
        return 0.0


def _as_int(value: Any) -> int:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0
    if isinstance(value, str):
        s = value.strip().replace(",", "")
        if not s or s.lower() in ("nan", "none", "-"):
            return 0
        try:
            return int(round(float(s)))
        except ValueError:
            return 0
    try:
        return int(round(float(value)))
    except (TypeError, ValueError):
        return 0


def _as_transfer(value: Any) -> int:
    s = _as_text(value).lower()
    if s in ("1", "是", "y", "yes", "true", "调拨"):
        return 1
    return 0


def parse_target_month(raw: str) -> date:
    """解析为目标月第一天。支持 2026-08 / 2026.8 / 2026/8月 等。"""
    s = str(raw).strip().replace("月", "").replace("年", "-").replace(".", "-").replace("/", "-")
    s = re.sub(r"\s+", "", s)
    if re.fullmatch(r"\d{2}-\d{1,2}", s):
        y, m = s.split("-")
        year, month = 2000 + int(y), int(m)
    elif re.fullmatch(r"\d{4}-\d{1,2}", s):
        y, m = s.split("-")
        year, month = int(y), int(m)
    else:
        raise ValueError(f"无法识别的月份格式: {raw}")
    if not 1 <= month <= 12:
        raise ValueError(f"月份须在 1-12 之间: {month}")
    return date(year, month, 1)


def resolve_column_map(columns: Iterable[Any]) -> dict[str, str]:
    """db_col → 实际 Excel 列名。"""
    headers = [_norm_header(c) for c in columns]
    header_set = {h: h for h in headers if h}
    resolved: dict[str, str] = {}

    for db_col, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in header_set:
                resolved[db_col] = alias
                break

    for db_col, suffixes in HIST_SUFFIX_MAP.items():
        if db_col in resolved:
            continue
        for h in headers:
            if not h:
                continue
            for suf in suffixes:
                if h == suf or h.endswith(suf):
                    # 避免「平均客单价」误命中「平均销量」类后缀
                    if db_col == "hist_avg_sales_qty" and "客单" in h:
                        continue
                    resolved[db_col] = h
                    break
            if db_col in resolved:
                break

    return resolved


def load_okr_excel(
    path: Path,
    *,
    sheet: str = DEFAULT_SHEET,
) -> pd.DataFrame:
    if not path.is_file():
        raise FileNotFoundError(f"Excel 不存在: {path}")
    df = pd.read_excel(path, sheet_name=sheet, dtype=object)
    df.columns = [_norm_header(c) for c in df.columns]
    # 去掉整行空
    if not df.empty:
        nonempty = df.apply(
            lambda row: any(_as_text(v) != "" for v in row.tolist()),
            axis=1,
        )
        df = df.loc[nonempty].reset_index(drop=True)
    return df


def build_rows(
    df: pd.DataFrame,
    *,
    target_month: date,
    col_map: dict[str, str],
) -> tuple[list[dict[str, Any]], int]:
    """返回 (可写入行, 跳过行数)。"""
    rows: list[dict[str, Any]] = []
    skipped = 0
    seen_keys: set[tuple[str, str, str]] = set()

    for rec in df.to_dict(orient="records"):
        def cell(db_col: str) -> Any:
            excel_col = col_map.get(db_col)
            if not excel_col:
                return None
            return rec.get(excel_col)

        market = _as_text(cell("market_code"), max_len=32)
        account = _as_text(cell("account_code"), max_len=64)
        sku = _as_text(cell("product_sku"), max_len=64)
        if not market or not account or not sku:
            skipped += 1
            continue

        uk = (account, sku, market)
        if uk in seen_keys:
            skipped += 1
            continue
        seen_keys.add(uk)

        row: dict[str, Any] = {
            "target_month": target_month,
            "market_code": market,
            "account_code": account,
            "product_sku": sku,
            "product_uid": _as_text(cell("product_uid"), max_len=64),
            "accounting_sku": _as_text(cell("accounting_sku"), max_len=64),
            "identify_sku": _as_text(cell("identify_sku"), max_len=128),
            "identify_sku_uid": _as_text(cell("identify_sku_uid"), max_len=128),
            "category": _as_text(cell("category"), max_len=64),
            "product_status": _as_text(cell("product_status"), max_len=32),
            "ops_owner": _as_text(cell("ops_owner"), max_len=64),
            "dispatch_warehouse": _as_text(cell("dispatch_warehouse"), max_len=32),
            "is_transfer": _as_transfer(cell("is_transfer")),
            "remark": _as_text(cell("remark"), max_len=512),
            "source_sheet": _as_text(cell("source_sheet"), max_len=64),
            "deleted_at": None,
        }

        for col in RATE_COLS + DECIMAL4_COLS:
            row[col] = _as_float(cell(col), ndigits=4)
        for col in DECIMAL6_COLS:
            row[col] = _as_float(cell(col), ndigits=6)
        for col in INT_COLS:
            row[col] = _as_int(cell(col))

        rows.append(row)

    return rows, skipped


def soft_delete_month(target_month: date, *, dry_run: bool = False) -> int:
    if dry_run:
        db = get_db_manager()
        conn = db.get_connection()
        try:
            with conn.cursor() as cur:
                cur.execute(
                    f"""
                    SELECT COUNT(*) FROM `{TABLE}`
                    WHERE `target_month` = %s AND `deleted_at` IS NULL
                    """,
                    (target_month,),
                )
                return int(cur.fetchone()[0] or 0)
        finally:
            conn.close()

    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(SOFT_DELETE_SQL, (target_month,))
            n = int(cur.rowcount or 0)
        conn.commit()
        return n
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
                params = [tuple(r[c] for c in UPSERT_COLS) for r in chunk]
                cur.executemany(UPSERT_SQL, params)
                written += len(chunk)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return written


def resolve_excel_path(
    *,
    month: date | None,
    file: str | None,
    excel_dir: str | None,
) -> tuple[Path, date]:
    if file:
        path = Path(file)
        if month is None:
            # 从文件名推断：2026-08月目标拆解及跟进.xlsx
            m = re.search(r"(\d{4})[-.](\d{1,2})", path.name)
            if not m:
                raise ValueError(
                    f"无法从文件名推断月份，请加 --month：{path.name}"
                )
            month = date(int(m.group(1)), int(m.group(2)), 1)
        return path, month

    if month is None:
        raise ValueError("请指定 --month 或 --file")

    path = Path(
        _okr.get_monthly_path(month.year, month.month, excel_dir=excel_dir)
    )
    return path, month


def import_sales_targets(
    *,
    excel_path: Path,
    target_month: date,
    sheet: str = DEFAULT_SHEET,
    dry_run: bool = False,
    preview: int = 0,
) -> int:
    print("=" * 60)
    print("导入 snapshot_sales_targets")
    print("=" * 60)
    print(f"文件: {excel_path}")
    print(f"Sheet: {sheet}")
    print(f"目标月: {target_month:%Y-%m}")
    print(f"模式: {'DRY-RUN' if dry_run else '写入'}")

    df = load_okr_excel(excel_path, sheet=sheet)
    print(f"读取行数: {len(df)}")

    col_map = resolve_column_map(df.columns)
    required = ("market_code", "account_code", "product_sku")
    missing_req = [c for c in required if c not in col_map]
    if missing_req:
        raise KeyError(
            f"Excel 缺少必填列映射 {missing_req}；"
            f"现有列: {list(df.columns)}"
        )

    expected_from_excel = set(COLUMN_ALIASES) | set(HIST_SUFFIX_MAP)
    unresolved = sorted(
        c
        for c in expected_from_excel
        if c not in col_map and c != "is_transfer"
    )
    # is_transfer 有别名，但上面已在 COLUMN_ALIASES；未匹配时默认 0
    if "is_transfer" not in col_map:
        unresolved.append("is_transfer")
    if unresolved:
        print(f"警告：未匹配到 Excel 列（将按 0/空写入）: {', '.join(unresolved)}")

    print("列映射:")
    for db_col in sorted(col_map):
        print(f"  {db_col} ← {col_map[db_col]}")

    rows, skipped = build_rows(df, target_month=target_month, col_map=col_map)
    print(f"可导入: {len(rows)} 行，跳过: {skipped} 行")

    if preview > 0 and rows:
        print(f"\n预览前 {min(preview, len(rows))} 行:")
        for r in rows[:preview]:
            print(
                f"  {r['market_code']}/{r['account_code']}/{r['product_sku']} "
                f"qty={r['target_sales_qty']} owner={r['ops_owner']!r}"
            )

    soft_n = soft_delete_month(target_month, dry_run=dry_run)
    print(f"软删除当月旧行: {soft_n}" + (" (dry-run 计数)" if dry_run else ""))

    written = upsert_rows(rows, dry_run=dry_run)
    print(f"{'将写入' if dry_run else '已写入'}: {written} 行 → `{TABLE}`")
    print("=" * 60)
    return written


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="导入 OKR 月目标拆解 Excel → snapshot_sales_targets"
    )
    parser.add_argument(
        "-m",
        "--month",
        default=None,
        help="目标月 yyyy-mm（默认取 OKR月目标拆分.OKR_MONTH / 环境变量 OKR_MONTH）",
    )
    parser.add_argument(
        "-f",
        "--file",
        default=None,
        help="Excel 完整路径（默认按月份拼到 EXCEL_DIR）",
    )
    parser.add_argument(
        "--dir",
        default=None,
        help=f"Excel 目录，默认 {_okr.EXCEL_DIR}",
    )
    parser.add_argument(
        "--sheet",
        default=DEFAULT_SHEET,
        help=f"工作表名，默认 {DEFAULT_SHEET}",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只解析不写库",
    )
    parser.add_argument(
        "--preview",
        type=int,
        default=0,
        help="打印前 N 行预览",
    )
    args = parser.parse_args(argv)

    month: date | None = None
    if args.month:
        month = parse_target_month(args.month)
    elif not args.file:
        month_str = _okr.get_okr_month()
        if month_str:
            month = parse_target_month(month_str)
        else:
            raw = input("请输入月份（如 2026-08）: ").strip()
            month = parse_target_month(raw)

    try:
        excel_path, target_month = resolve_excel_path(
            month=month,
            file=args.file,
            excel_dir=args.dir,
        )
        import_sales_targets(
            excel_path=excel_path,
            target_month=target_month,
            sheet=args.sheet,
            dry_run=args.dry_run,
            preview=args.preview,
        )
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        import traceback

        traceback.print_exc()
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
