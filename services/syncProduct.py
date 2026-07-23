"""从 BTH全部SKU明细（基础数据维护）同步更新 product_sku 表。

数据源：``config.A0_paths.BTH_ALL_SKU_DETAIL_PATH``
目标表：``product_sku``（按 ``product_sku`` 唯一键 UPSERT）

仅写入 Excel 可提供的字段；不写 ``product_name_cn``。
ERP 字段（product_uid / EAN / 报关 / 图片等）在 UPDATE 时保留。
``line_hash`` 用于变更检测：内容未变则跳过写入。

用法（项目根目录）::

    python -m services.syncProduct
    python -m services.syncProduct --dry-run
    python -m services.syncProduct --limit 100
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import re
import sys
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any

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
_epr_mod.bootstrap(__file__)

from common.style import Color  # noqa: E402
from config.A0_paths import BTH_ALL_SKU_DETAIL_PATH  # noqa: E402
from database.db_connection import get_db_manager  # noqa: E402

SHEET_NAME = "基础数据维护"
TABLE = "product_sku"
BATCH_SIZE = 500

# Excel 列字母（与「基础数据维护」表头一致，打开文件即可对照）
# A=序号旁列 … B=SKU … F~H=内箱长宽高 … I~K=外箱 … L=箱规 … M=箱毛重
# N~R=头程 EU/AU|US|CA|JP|UK … S~W=关税含税 EU|US|CA/AU|JP|UK
# AC=原始采购价 AE=供应商简称 AF=品类 AG=运营模式 AL=供应商全称
_COL = {
    "product_sku": "B",
    "cost_price_cny": "C",          # 成本价
    "unit_weight_g": "D",           # 重量（g)
    "inner_box_l_cm": "F",          # 内箱-长
    "inner_box_w_cm": "G",          # 内箱-宽
    "inner_box_h_cm": "H",          # 内箱-高
    "outer_box_l_cm": "I",          # 外箱-长
    "outer_box_w_cm": "J",          # 外箱-宽
    "outer_box_h_cm": "K",          # 外箱-高
    "carton_qty": "L",              # 每箱数量
    "carton_gross_g": "M",          # 箱规毛重
    "first_leg_eu_au_cny": "N",     # 头程 EU/AU
    "first_leg_us_cny": "O",        # 头程 US
    "first_leg_uk_cny": "R",        # 头程 UK
    "duty_eu_cny": "S",             # 关税含税 EU
    "duty_us_cny": "T",             # 关税含税 US
    "duty_uk_cny": "W",             # 关税含税 UK
    "purchase_price": "AC",         # 原始采购价
    "supplier_abbr": "AE",          # 供应商简称
    "category_lv1": "AF",           # 品类
    "ops_model": "AG",              # 运营模式
    "supplier_full": "AL",          # 供应商全称 F003[公司名]
}


def _excel_col_to_idx(col: str) -> int:
    """Excel 列字母 → 0-based 列下标（A→0, B→1, …, Z→25, AA→26）。"""
    col = col.strip().upper()
    if not col.isalpha():
        raise ValueError(f"非法 Excel 列字母：{col!r}")
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


_COL_IDX = {key: _excel_col_to_idx(letter) for key, letter in _COL.items()}
_MIN_COLS = max(_COL_IDX.values()) + 1

# UPSERT 写入的业务字段（顺序固定，参与 line_hash）
_SYNC_FIELDS = (
    "product_sku",
    "category_lv1",
    "supplier_abbr",
    "supplier_name",
    "purchase_price",
    "cost_price_cny",
    "unit_weight_g",
    "carton_qty",
    "carton_gross_g",
    "inner_box_l_cm",
    "inner_box_w_cm",
    "inner_box_h_cm",
    "outer_box_l_cm",
    "outer_box_w_cm",
    "outer_box_h_cm",
    "first_leg_eu_au_cny",
    "first_leg_us_cny",
    "first_leg_uk_cny",
    "duty_eu_cny",
    "duty_us_cny",
    "duty_uk_cny",
    "ops_model",
)

_SUPPLIER_FULL_RE = re.compile(r"^[^\[\]]+\[(.+)\]\s*$")

_UPSERT_SQL = f"""
INSERT INTO `{TABLE}` (
    line_hash,
    product_sku,
    category_lv1,
    supplier_abbr,
    supplier_name,
    purchase_price,
    cost_price_cny,
    unit_weight_g,
    carton_qty,
    carton_gross_g,
    inner_box_l_cm,
    inner_box_w_cm,
    inner_box_h_cm,
    outer_box_l_cm,
    outer_box_w_cm,
    outer_box_h_cm,
    first_leg_eu_au_cny,
    first_leg_us_cny,
    first_leg_uk_cny,
    duty_eu_cny,
    duty_us_cny,
    duty_uk_cny,
    ops_model,
    source_type,
    is_deleted
) VALUES (
    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
    'Excel', 0
)
ON DUPLICATE KEY UPDATE
    line_hash = VALUES(line_hash),
    category_lv1 = VALUES(category_lv1),
    supplier_abbr = IF(VALUES(supplier_abbr) <> '', VALUES(supplier_abbr), supplier_abbr),
    supplier_name = IF(VALUES(supplier_name) IS NOT NULL AND VALUES(supplier_name) <> '',
                       VALUES(supplier_name), supplier_name),
    purchase_price = VALUES(purchase_price),
    cost_price_cny = VALUES(cost_price_cny),
    unit_weight_g = VALUES(unit_weight_g),
    carton_qty = VALUES(carton_qty),
    carton_gross_g = VALUES(carton_gross_g),
    inner_box_l_cm = VALUES(inner_box_l_cm),
    inner_box_w_cm = VALUES(inner_box_w_cm),
    inner_box_h_cm = VALUES(inner_box_h_cm),
    outer_box_l_cm = VALUES(outer_box_l_cm),
    outer_box_w_cm = VALUES(outer_box_w_cm),
    outer_box_h_cm = VALUES(outer_box_h_cm),
    first_leg_eu_au_cny = VALUES(first_leg_eu_au_cny),
    first_leg_us_cny = VALUES(first_leg_us_cny),
    first_leg_uk_cny = VALUES(first_leg_uk_cny),
    duty_eu_cny = VALUES(duty_eu_cny),
    duty_us_cny = VALUES(duty_us_cny),
    duty_uk_cny = VALUES(duty_uk_cny),
    ops_model = VALUES(ops_model),
    is_deleted = 0
"""


def _cell(row: pd.Series, key: str) -> Any:
    idx = _COL_IDX[key]
    if idx >= len(row):
        return None
    return row.iloc[idx]


def _is_empty(val: Any) -> bool:
    if val is None:
        return True
    if isinstance(val, float) and pd.isna(val):
        return True
    if isinstance(val, str) and not val.strip():
        return True
    return False


def _as_str(val: Any, *, default: str = "") -> str:
    if _is_empty(val):
        return default
    return str(val).strip()


def _as_decimal(val: Any, *, quant: str | None = "0.0001") -> Decimal | None:
    if _is_empty(val):
        return None
    try:
        d = Decimal(str(val).strip().replace(",", ""))
    except (InvalidOperation, ValueError):
        return None
    if quant is None:
        return d
    return d.quantize(Decimal(quant), rounding=ROUND_HALF_UP)


def _as_int(val: Any) -> int | None:
    d = _as_decimal(val, quant="1")
    if d is None:
        return None
    return int(d)


def _parse_supplier_name(full: Any) -> str | None:
    """从「供应商.1」解析全称 ``F003[公司名]`` → ``公司名``；无则返回 None（不覆盖库内已有全称）。"""
    raw = _as_str(full)
    if not raw:
        return None
    m = _SUPPLIER_FULL_RE.match(raw)
    name = (m.group(1).strip() if m else raw).strip()
    return name or None


def _hash_norm(val: Any) -> str:
    if val is None:
        return ""
    if isinstance(val, Decimal):
        return format(val, "f")
    if isinstance(val, int):
        return str(val)
    return str(val).strip()


def compute_line_hash(record: dict[str, Any]) -> str:
    """对 Excel 同步字段做稳定 SHA-256（字段顺序固定）。"""
    parts = [_hash_norm(record.get(k)) for k in _SYNC_FIELDS]
    payload = "\x1f".join(parts)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def load_excel_rows(path: str | Path) -> list[dict[str, Any]]:
    """读取「基础数据维护」，跳过双行表头中的副标题行与空 SKU。"""
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(f"BTH SKU 明细不存在：{path}")

    # header=None：按列字母定位，避免 Unnamed 列名随 pandas 变化
    df = pd.read_excel(path, sheet_name=SHEET_NAME, header=None)
    if df.shape[1] < _MIN_COLS:
        raise ValueError(
            f"工作表「{SHEET_NAME}」列数过少（{df.shape[1]} < {_MIN_COLS}），"
            "请确认文件版本"
        )

    rows: list[dict[str, Any]] = []
    seen: set[str] = set()
    for i in range(len(df)):
        raw = df.iloc[i]
        sku = _as_str(_cell(raw, "product_sku"))
        if not sku or sku.upper() == "SKU":
            continue
        if sku in seen:
            # 同 SKU 多次出现时保留最后一次（与常见映射表习惯一致）
            rows = [r for r in rows if r["product_sku"] != sku]
        seen.add(sku)

        record = {
            "product_sku": sku,
            "category_lv1": _as_str(_cell(raw, "category_lv1")) or None,
            "supplier_abbr": _as_str(_cell(raw, "supplier_abbr")),
            "supplier_name": _parse_supplier_name(_cell(raw, "supplier_full")),
            "purchase_price": _as_decimal(_cell(raw, "purchase_price")),
            "cost_price_cny": _as_decimal(_cell(raw, "cost_price_cny")),
            "unit_weight_g": _as_decimal(_cell(raw, "unit_weight_g"), quant="0.01"),
            "carton_qty": _as_int(_cell(raw, "carton_qty")),
            "carton_gross_g": _as_decimal(_cell(raw, "carton_gross_g"), quant="0.01"),
            "inner_box_l_cm": _as_decimal(_cell(raw, "inner_box_l_cm"), quant="0.01"),
            "inner_box_w_cm": _as_decimal(_cell(raw, "inner_box_w_cm"), quant="0.01"),
            "inner_box_h_cm": _as_decimal(_cell(raw, "inner_box_h_cm"), quant="0.01"),
            "outer_box_l_cm": _as_decimal(_cell(raw, "outer_box_l_cm"), quant="0.01"),
            "outer_box_w_cm": _as_decimal(_cell(raw, "outer_box_w_cm"), quant="0.01"),
            "outer_box_h_cm": _as_decimal(_cell(raw, "outer_box_h_cm"), quant="0.01"),
            "first_leg_eu_au_cny": _as_decimal(_cell(raw, "first_leg_eu_au_cny")),
            "first_leg_us_cny": _as_decimal(_cell(raw, "first_leg_us_cny")),
            "first_leg_uk_cny": _as_decimal(_cell(raw, "first_leg_uk_cny")),
            "duty_eu_cny": _as_decimal(_cell(raw, "duty_eu_cny")),
            "duty_us_cny": _as_decimal(_cell(raw, "duty_us_cny")),
            "duty_uk_cny": _as_decimal(_cell(raw, "duty_uk_cny")),
            "ops_model": _as_str(_cell(raw, "ops_model")),
        }
        record["line_hash"] = compute_line_hash(record)
        rows.append(record)
    return rows


def fetch_existing_hashes(skus: list[str]) -> dict[str, str | None]:
    """批量读取已有 product_sku → line_hash。"""
    if not skus:
        return {}
    db = get_db_manager()
    conn = db.get_connection()
    result: dict[str, str | None] = {}
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            for i in range(0, len(skus), BATCH_SIZE):
                chunk = skus[i : i + BATCH_SIZE]
                placeholders = ",".join(["%s"] * len(chunk))
                cur.execute(
                    f"SELECT product_sku, line_hash FROM `{TABLE}` "
                    f"WHERE product_sku IN ({placeholders})",
                    tuple(chunk),
                )
                for row in cur.fetchall():
                    result[str(row["product_sku"])] = row.get("line_hash")
    finally:
        conn.close()
    return result


def _record_to_params(rec: dict[str, Any]) -> tuple[Any, ...]:
    return (
        rec["line_hash"],
        rec["product_sku"],
        rec["category_lv1"],
        rec["supplier_abbr"],
        rec["supplier_name"],
        rec["purchase_price"],
        rec["cost_price_cny"],
        rec["unit_weight_g"],
        rec["carton_qty"],
        rec["carton_gross_g"],
        rec["inner_box_l_cm"],
        rec["inner_box_w_cm"],
        rec["inner_box_h_cm"],
        rec["outer_box_l_cm"],
        rec["outer_box_w_cm"],
        rec["outer_box_h_cm"],
        rec["first_leg_eu_au_cny"],
        rec["first_leg_us_cny"],
        rec["first_leg_uk_cny"],
        rec["duty_eu_cny"],
        rec["duty_us_cny"],
        rec["duty_uk_cny"],
        rec["ops_model"],
    )


def upsert_rows(rows: list[dict[str, Any]]) -> int:
    """批量 UPSERT，返回 executemany 影响行数累计。"""
    if not rows:
        return 0
    db = get_db_manager()
    conn = db.get_connection()
    affected = 0
    try:
        with conn.cursor() as cur:
            for i in range(0, len(rows), BATCH_SIZE):
                chunk = rows[i : i + BATCH_SIZE]
                params = [_record_to_params(r) for r in chunk]
                n = cur.executemany(_UPSERT_SQL, params)
                affected += int(n or 0)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return affected


def sync_product_sku(
    *,
    excel_path: str | Path | None = None,
    dry_run: bool = False,
    limit: int | None = None,
) -> dict[str, int]:
    """
    同步入口。

    Returns:
        统计：excel_rows / unchanged / to_insert / to_update / written
    """
    path = excel_path or BTH_ALL_SKU_DETAIL_PATH
    print(f"{Color.CYAN}[读取] {path}{Color.RESET}")
    print(f"{Color.CYAN}[工作表] {SHEET_NAME}{Color.RESET}")

    records = load_excel_rows(path)
    if limit is not None:
        records = records[: max(0, limit)]
    print(f"{Color.GREEN}[Excel] 有效 SKU {len(records)} 条{Color.RESET}")

    existing = fetch_existing_hashes([r["product_sku"] for r in records])
    to_write: list[dict[str, Any]] = []
    unchanged = 0
    to_insert = 0
    to_update = 0
    for rec in records:
        sku = rec["product_sku"]
        old_hash = existing.get(sku)
        if sku not in existing:
            to_insert += 1
            to_write.append(rec)
        elif old_hash == rec["line_hash"]:
            unchanged += 1
        else:
            to_update += 1
            to_write.append(rec)

    print(
        f"{Color.CYAN}[变更] 新增 {to_insert}，更新 {to_update}，"
        f"未变 {unchanged}{Color.RESET}"
    )

    written = 0
    if dry_run:
        print(f"{Color.YELLOW}[dry-run] 跳过数据库写入{Color.RESET}")
    else:
        written = upsert_rows(to_write)
        print(f"{Color.GREEN}[写入] executemany 影响行数 {written}{Color.RESET}")

    return {
        "excel_rows": len(records),
        "unchanged": unchanged,
        "to_insert": to_insert,
        "to_update": to_update,
        "written": written,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="根据 BTH全部SKU明细 同步更新 product_sku 表",
    )
    parser.add_argument(
        "--path",
        default=None,
        help="覆盖默认 BTH_ALL_SKU_DETAIL_PATH",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只统计变更，不写库",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="仅处理前 N 条（调试用）",
    )
    args = parser.parse_args(argv)

    try:
        sync_product_sku(
            excel_path=args.path,
            dry_run=args.dry_run,
            limit=args.limit,
        )
    except FileNotFoundError as exc:
        print(f"{Color.RED}[ERROR] {exc}{Color.RESET}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"{Color.RED}[ERROR] {exc}{Color.RESET}", file=sys.stderr)
        raise
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
