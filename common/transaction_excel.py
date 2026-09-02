"""Amazon transaction 交易明细 Excel 读取（兼容新旧 ERP 导出格式）。"""

from __future__ import annotations

import re
import warnings
from pathlib import Path

import pandas as pd

# 新格式：前 3 行为 Date Time / 币种 / 交易状态，第 4 行为列名；旧格式：第 1 行即列名
_TRANSACTION_HEADER_KEYS = frozenset({"order id", "seller sku", "fba fees"})
_REQUIRED_COLUMNS = ("order id", "seller sku", "fba fees")
# 仅订单级 FBA 派送费（排除 fba_transaction_fees 等无 order id 的汇总行）
_ORDER_FBA_FEE_TYPES = frozenset({"payment_order", "refund_order"})
_INVALID_ORDER_IDS = frozenset({"null", "null[null]", "nan", "none", ""})


def find_transaction_header_row(xlsx_path, *, max_scan: int = 12) -> int:
    """定位 transaction 表头行（兼容 header=0 或 header=3）。"""
    preview = pd.read_excel(xlsx_path, header=None, nrows=max_scan, engine="openpyxl")
    for i in range(len(preview)):
        row_vals = {
            str(v).replace("\n", " ").strip().lower()
            for v in preview.iloc[i]
            if pd.notna(v) and str(v).strip()
        }
        if _TRANSACTION_HEADER_KEYS.issubset(row_vals):
            return i
    raise ValueError(
        f"未在前 {max_scan} 行找到 transaction 表头（需含 order id / seller sku / fba fees）：{xlsx_path}"
    )


def _normalize_column_names(columns) -> list[str]:
    normalized = []
    for col in columns:
        name = ("" if col is None else str(col)).replace("\n", " ").strip()
        normalized.append(name)
    return normalized


def _canonicalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """英文列名统一为小写，便于新旧导出混用。"""
    rename = {}
    for col in df.columns:
        key = str(col).replace("\n", " ").strip().lower()
        if key in _TRANSACTION_HEADER_KEYS and col != key:
            rename[col] = key
    if rename:
        df = df.rename(columns=rename)
    return df


def read_transaction_excel(xlsx_path) -> pd.DataFrame:
    """读取 transaction 源表，自动跳过元数据行并校验必要列。"""
    path = Path(xlsx_path)
    if not path.is_file():
        raise FileNotFoundError(f"未找到 transaction 文件：{path}")

    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
        header_row = find_transaction_header_row(path)
        df = pd.read_excel(path, header=header_row, engine="openpyxl")

    df.columns = _normalize_column_names(df.columns)
    df = _canonicalize_columns(df)
    df = df.dropna(how="all")

    missing = [c for c in _REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise KeyError(
            f"transaction 文件缺少列 {missing}（header={header_row}），"
            f"当前列: {list(df.columns)[:15]}..."
        )
    return df


def extract_platform_sku(value) -> str | None:
    """从 seller sku 提取平台 SKU（与订单统计 SKU 对齐）。"""
    if pd.isna(value):
        return None
    s = str(value).strip()
    if not s:
        return None
    if "amzn.gr." in s:
        return s.split("amzn.gr.")[-1].split("-")[0].split("_")[0]
    return s.split("#")[0].split("BCFBAFL")[0].split("FBFBAFL")[0]


def _is_valid_order_id(value) -> bool:
    if pd.isna(value):
        return False
    s = str(value).strip()
    if not s or s.lower() in _INVALID_ORDER_IDS:
        return False
    return bool(re.match(r"^(\d{3}-\d{7}-\d{7}|S\d{2}-\d{7}-\d{7})$", s))


def filter_fba_fee_rows(df: pd.DataFrame) -> pd.DataFrame:
    """筛选可用于 FBA 派送费映射的订单行。"""
    out = df.copy()
    out["fba fees"] = pd.to_numeric(out["fba fees"], errors="coerce").fillna(0)
    out = out[out["fba fees"] != 0]

    if "费用类型" in out.columns:
        fee_type = out["费用类型"].astype(str).str.strip().str.lower()
        out = out[fee_type.isin(_ORDER_FBA_FEE_TYPES)]

    out = out[out["order id"].map(_is_valid_order_id)]
    out = out[out["seller sku"].notna() & (out["seller sku"].astype(str).str.strip() != "")]
    return out


def resolve_transaction_source_file(
    tx_dir: Path,
    *,
    kind: str,
    transaction_date: str,
) -> Path:
    """解析已发放/已推迟 transaction 源文件（精确名优先，否则 glob）。"""
    labels = {"released": "已发放", "deferred": "已推迟"}
    if kind not in labels:
        raise ValueError(f"kind 只能是 {tuple(labels)}，当前: {kind!r}")

    label = labels[kind]
    exact = tx_dir / f"transaction交易明细-{label}订单{transaction_date}.xlsx"
    if exact.is_file():
        return exact

    pattern = f"transaction交易明细-{label}订单*.xlsx"
    candidates = sorted(tx_dir.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)
    if not candidates:
        raise FileNotFoundError(f"未找到 {pattern}，目录：{tx_dir}")

    for path in candidates:
        if transaction_date in path.stem:
            return path
    return candidates[0]


def resolve_transaction_processed_file(tx_dir: Path, transaction_date: str) -> Path:
    """解析 B4_1 产出的 (处理完成)transaction…xlsx。"""
    exact = tx_dir / f"(处理完成)transaction交易明细_已发放-推迟订单{transaction_date}.xlsx"
    if exact.is_file():
        return exact

    candidates = sorted(
        tx_dir.glob("(处理完成)transaction交易明细_已发放-推迟订单*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not candidates:
        raise FileNotFoundError(
            f"未找到 (处理完成)transaction交易明细_已发放-推迟订单*.xlsx，请先运行 B4_1，目录：{tx_dir}"
        )

    for path in candidates:
        if transaction_date in path.stem:
            return path
    return candidates[0]
