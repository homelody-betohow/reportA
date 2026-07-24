"""行哈希与 INSERT IGNORE 辅助，供 snapshot 等脚本共用。"""

from __future__ import annotations

import hashlib
import json
import math
from datetime import date, datetime
from decimal import Decimal
from typing import Any, Iterable, Sequence


def norm_for_hash(v: Any) -> Any:
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.isoformat(sep=" ", timespec="seconds")
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, Decimal):
        return format(v.normalize(), "f")
    if isinstance(v, bool):
        return v
    if isinstance(v, int):
        return v
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return None
        return format(Decimal(str(v)).normalize(), "f")
    if isinstance(v, str):
        t = v.strip()
        return t if t else None
    return str(v).strip() or None


def row_subset_for_line_hash(row: dict[str, Any], keys: Sequence[str]) -> dict[str, Any]:
    return {k: row.get(k) for k in keys}


def stable_line_hash(field_values: dict[str, Any]) -> str:
    payload = {k: norm_for_hash(field_values[k]) for k in sorted(field_values.keys())}
    raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":"), sort_keys=True)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def insert_ignore_rows(
    conn,
    *,
    table: str,
    columns: Sequence[str],
    rows: Iterable[tuple[Any, ...]],
    chunk_size: int = 300,
) -> int:
    """INSERT IGNORE：遇唯一键冲突则跳过。返回累计尝试写入行数。"""
    if not columns:
        return 0
    cols_sql = ", ".join(f"`{c}`" for c in columns)
    placeholders = ", ".join(["%s"] * len(columns))
    sql = f"INSERT IGNORE INTO `{table}` ({cols_sql}) VALUES ({placeholders})"
    cur = conn.cursor()
    n = 0
    buf: list[tuple[Any, ...]] = []
    try:
        for row in rows:
            buf.append(row)
            if len(buf) >= chunk_size:
                cur.executemany(sql, buf)
                n += cur.rowcount if cur.rowcount and cur.rowcount > 0 else 0
                buf.clear()
        if buf:
            cur.executemany(sql, buf)
            if cur.rowcount and cur.rowcount > 0:
                n += cur.rowcount
    finally:
        cur.close()
    return n
