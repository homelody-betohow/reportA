"""
upProductMapping.py — 回填 product_sku_mapping（HY / warehouse）

流程：
  1. 主流程：读取 product_sku 为空的行
     - warehouse_sku（已无 900008-）按尾缀规则清洗
     - 去 product_sku 表匹配；命中回写 product_sku / seller_ean / is_active=1 等
  2. 兜底：读取 product_sku 非空且 seller_ean 为空的行
     - 用已有 product_sku 去 product_sku 表匹配 product_uid
     - 命中则回写 seller_ean，并设 is_active=1

尾缀规则（与 K1_HY_仓租.py 一致，另含仓租常见 BC/BTL）：
  -KA/-JI/-CH/-DA/-FB/-AT/-BC/-C1/-C2/-C3/-ECO/-REAL/-ES/-4PX/-UMI/-BTL/-MF/
  -KL/-YES/-ML/-ZSJ/-ZJF

用法：
  python modules/K_storage_fee_product_info/upProductMapping.py
"""

from __future__ import annotations

import hashlib
import importlib.util
import re
from pathlib import Path

import pymysql.cursors

_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.style import Color  # noqa: E402
from database.db_connection import get_db_manager  # noqa: E402

PRODUCT_SKU_TABLE = "product_sku"
PSM_TABLE = "product_sku_mapping"
PARTNER_CODE_HY = "HY"
PARTNER_TYPE_WH = "warehouse"
_KEY_CHUNK = 500

SUFFIXES_TO_REMOVE = (
    "KA", "JI", "CH", "DA", "FB", "AT", "BC", "NW",
    "C1", "C2", "C3", "ECO", "REAL", "ES", "4PX",
    "UMI", "BTL", "MF", "KL", "YES", "ML", "ZSJ", "ZJF",
)
_SUFFIX_RE = re.compile(
    r"-(?:" + "|".join(re.escape(s) for s in SUFFIXES_TO_REMOVE) + r")$",
    re.IGNORECASE,
)

_PSM_HASH_FIELDS = (
    "partner_code",
    "partner_type",
    "shop_hash",
    "seller_sku",
    "warehouse_sku",
    "mapping_type",
    "product_sku",
)


def clean_warehouse_sku(raw: str) -> str:
    """仅去尾缀（warehouse_sku 已无 900008- 前缀）。"""
    s = str(raw or "").strip()
    if not s:
        return ""
    return _SUFFIX_RE.sub("", s).strip()


def _psm_line_hash(*, warehouse_sku: str, product_sku: str) -> str:
    record = {
        "partner_code": PARTNER_CODE_HY,
        "partner_type": PARTNER_TYPE_WH,
        "shop_hash": "",
        "seller_sku": "",
        "warehouse_sku": warehouse_sku,
        "mapping_type": "single",
        "product_sku": product_sku,
    }
    parts = [str(record.get(k) or "").strip() for k in _PSM_HASH_FIELDS]
    return hashlib.sha256("\x1f".join(parts).encode("utf-8")).hexdigest()


def _query_rows(sql: str) -> list[dict]:
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            cur.execute(sql, (PARTNER_CODE_HY, PARTNER_TYPE_WH))
            return list(cur.fetchall())
    finally:
        conn.close()


def fetch_empty_product_sku_rows() -> list[dict]:
    """product_sku 为空。"""
    return _query_rows(
        f"""
        SELECT id, warehouse_sku
        FROM `{PSM_TABLE}`
        WHERE partner_code = %s
          AND partner_type = %s
          AND (product_sku IS NULL OR TRIM(product_sku) = '')
        ORDER BY id
        """
    )


def fetch_empty_seller_ean_rows() -> list[dict]:
    """product_sku 非空且 seller_ean 为空（兜底）。"""
    return _query_rows(
        f"""
        SELECT id, warehouse_sku, product_sku
        FROM `{PSM_TABLE}`
        WHERE partner_code = %s
          AND partner_type = %s
          AND product_sku IS NOT NULL
          AND TRIM(product_sku) <> ''
          AND (seller_ean IS NULL OR TRIM(seller_ean) = '')
        ORDER BY id
        """
    )


def fetch_product_sku_uid_map(skus: list[str]) -> dict[str, str]:
    """product_sku → product_uid（is_deleted=0）；无 uid 时值为 ''。"""
    skus = sorted({str(x).strip() for x in skus if x and str(x).strip()})
    if not skus:
        return {}

    found: dict[str, str] = {}
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            for i in range(0, len(skus), _KEY_CHUNK):
                chunk = skus[i : i + _KEY_CHUNK]
                placeholders = ", ".join(["%s"] * len(chunk))
                sql = f"""
                    SELECT product_sku, product_uid
                    FROM `{PRODUCT_SKU_TABLE}`
                    WHERE product_sku IN ({placeholders})
                      AND is_deleted = 0
                """
                cur.execute(sql, chunk)
                for row in cur.fetchall():
                    sku = str(row.get("product_sku") or "").strip()
                    if sku:
                        found[sku] = str(row.get("product_uid") or "").strip()
    finally:
        conn.close()
    return found


def backfill_product_sku_rows(updates: list[tuple[int, str, str, str]]) -> int:
    """主流程回写：updates = (id, product_sku, seller_ean, line_hash)"""
    if not updates:
        return 0

    sql = f"""
        UPDATE `{PSM_TABLE}`
        SET product_sku = %s,
            seller_ean = %s,
            mapping_type = 'single',
            component_info = NULL,
            is_active = 1,
            line_hash = %s,
            source_type = 'Auto'
        WHERE id = %s
          AND partner_code = %s
          AND partner_type = %s
          AND (product_sku IS NULL OR TRIM(product_sku) = '')
    """
    return _exec_updates(
        sql,
        [
            (product_sku, seller_ean, line_hash, row_id, PARTNER_CODE_HY, PARTNER_TYPE_WH)
            for row_id, product_sku, seller_ean, line_hash in updates
        ],
    )


def backfill_seller_ean_rows(updates: list[tuple[int, str]]) -> int:
    """兜底回写：updates = (id, seller_ean)；仅补 seller_ean。"""
    if not updates:
        return 0

    sql = f"""
        UPDATE `{PSM_TABLE}`
        SET seller_ean = %s,
            is_active = 1,
            source_type = 'Auto'
        WHERE id = %s
          AND partner_code = %s
          AND partner_type = %s
          AND product_sku IS NOT NULL
          AND TRIM(product_sku) <> ''
          AND (seller_ean IS NULL OR TRIM(seller_ean) = '')
    """
    return _exec_updates(
        sql,
        [
            (seller_ean, row_id, PARTNER_CODE_HY, PARTNER_TYPE_WH)
            for row_id, seller_ean in updates
        ],
    )


def _exec_updates(sql: str, params_list: list[tuple]) -> int:
    updated = 0
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor() as cur:
            for params in params_list:
                try:
                    cur.execute(sql, params)
                    updated += int(cur.rowcount or 0)
                except pymysql.err.IntegrityError as exc:
                    print(
                        f"{Color.YELLOW}[跳过] params={params} 唯一键冲突：{exc}{Color.RESET}"
                    )
        conn.commit()
    finally:
        conn.close()
    return updated


def _print_preview(preview: list[str], total: int) -> None:
    for line in preview:
        print(f"  {line}")
    if total > len(preview):
        print(f"  ... 另有 {total - len(preview)} 条")


def run_fill_empty_product_sku() -> tuple[int, int]:
    """
    主流程：product_sku 为空 → 清洗 warehouse_sku → 匹配 → 回写。
    返回 (待处理行数, 成功回写数)。
    """
    rows = fetch_empty_product_sku_rows()
    if not rows:
        print(f"{Color.GREEN}[主流程] 无 product_sku 为空的行{Color.RESET}")
        return 0, 0

    print(f"[主流程] product_sku 为空：{len(rows)} 行")

    candidates: dict[int, tuple[str, str]] = {}
    for row in rows:
        row_id = int(row["id"])
        wh = str(row.get("warehouse_sku") or "").strip()
        if not wh:
            continue
        cleaned = clean_warehouse_sku(wh)
        if cleaned:
            candidates[row_id] = (wh, cleaned)

    cleaned_skus = [c for _, c in candidates.values()]
    sku_uid_map = fetch_product_sku_uid_map(cleaned_skus)
    print(
        f"[主流程] 清洗后候选 {len(set(cleaned_skus))} 个，"
        f"product_sku 命中 {len(sku_uid_map)} 个"
    )

    updates: list[tuple[int, str, str, str]] = []
    preview: list[str] = []
    for row_id, (wh, cleaned) in candidates.items():
        if cleaned not in sku_uid_map:
            continue
        uid = sku_uid_map[cleaned]
        line_hash = _psm_line_hash(warehouse_sku=wh, product_sku=cleaned)
        updates.append((row_id, cleaned, uid, line_hash))
        if len(preview) < 15:
            preview.append(f"{wh} → {cleaned}" + (f"  uid={uid}" if uid else ""))

    if not updates:
        print(f"{Color.YELLOW}[主流程] 尾缀清洗后仍无命中{Color.RESET}")
        return len(rows), 0

    n = backfill_product_sku_rows(updates)
    print(f"{Color.GREEN}[主流程] 回填成功 {n} / {len(updates)} 行{Color.RESET}")
    _print_preview(preview, len(updates))
    return len(rows), n


def run_fill_empty_seller_ean() -> tuple[int, int]:
    """
    兜底：product_sku 非空且 seller_ean 为空 → 用 product_sku 匹配 product_uid → 回写 seller_ean。
    返回 (待处理行数, 成功回写数)。
    """
    rows = fetch_empty_seller_ean_rows()
    if not rows:
        print(f"{Color.GREEN}[兜底] 无「product_sku 有值且 seller_ean 为空」的行{Color.RESET}")
        return 0, 0

    print(f"[兜底] product_sku 有值、seller_ean 为空：{len(rows)} 行")

    skus = [
        str(r.get("product_sku") or "").strip()
        for r in rows
        if str(r.get("product_sku") or "").strip()
    ]
    sku_uid_map = fetch_product_sku_uid_map(skus)
    # 仅统计有非空 product_uid 的命中（空 uid 回写无意义）
    hit_with_uid = {k: v for k, v in sku_uid_map.items() if v}
    print(
        f"[兜底] product_sku 表命中 {len(sku_uid_map)} 个，"
        f"其中有 product_uid {len(hit_with_uid)} 个"
    )

    updates: list[tuple[int, str]] = []
    preview: list[str] = []
    for row in rows:
        row_id = int(row["id"])
        sku = str(row.get("product_sku") or "").strip()
        uid = hit_with_uid.get(sku)
        if not uid:
            continue
        updates.append((row_id, uid))
        if len(preview) < 15:
            preview.append(f"{sku} → seller_ean={uid}")

    if not updates:
        print(f"{Color.YELLOW}[兜底] 无一可回写的 product_uid{Color.RESET}")
        return len(rows), 0

    n = backfill_seller_ean_rows(updates)
    print(f"{Color.GREEN}[兜底] 回填 seller_ean 成功 {n} / {len(updates)} 行{Color.RESET}")
    _print_preview(preview, len(updates))
    return len(rows), n


def main() -> int:
    empty_cnt, filled_sku = run_fill_empty_product_sku()
    ean_cnt, filled_ean = run_fill_empty_seller_ean()

    print(
        f"[汇总] 主流程回填 {filled_sku}/{empty_cnt}；"
        f"兜底回填 seller_ean {filled_ean}/{ean_cnt}"
    )
    if empty_cnt == 0 and ean_cnt == 0:
        return 0
    if filled_sku == 0 and filled_ean == 0:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
