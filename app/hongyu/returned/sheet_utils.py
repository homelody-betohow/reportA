"""退件表格共用：列名清洗、SheetRow、DataFrame 转换、钉钉写入值、任务信息截断。"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List, Mapping, Optional, Sequence

import pandas as pd

from api.ding_disk.workbook import clean_cell

TASK_MSG_MAX = 500


@dataclass
class SheetRow:
    excel_row: int
    values: Dict[str, str]


def normalize_col_name(name: Any) -> str:
    """清洗列名：去空白，去掉表头必填星号（如 ``ERP订单号 *`` → ``ERP订单号``）。"""
    text = clean_cell(name)
    if text.endswith("*"):
        text = text[:-1].rstrip()
    return text


def has_col(df: pd.DataFrame, col_name: str) -> bool:
    return normalize_col_name(col_name) in {
        normalize_col_name(c) for c in df.columns
    }


def sheet_col_index(df: pd.DataFrame, col_name: str) -> int:
    cols = [normalize_col_name(c) for c in df.columns]
    try:
        return cols.index(normalize_col_name(col_name))
    except ValueError as exc:
        raise KeyError(
            f"表格缺少列: [{col_name}]；实际列={list(df.columns)}"
        ) from exc


def task_msg(text: str, *, limit: int = TASK_MSG_MAX) -> str:
    return text[:limit]


def cell_write_value(val: Any) -> str:
    """钉钉 ranges 写入要求字符串（否则报 String is mandatory）。"""
    if val is None:
        return ""
    if isinstance(val, float) and pd.isna(val):
        return ""
    if isinstance(val, bool):
        return "1" if val else "0"
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    if isinstance(val, (int, float)):
        return str(val)
    return str(val)


def dataframe_to_rows(
    df: pd.DataFrame,
    *,
    bool_as_true_false: bool = False,
) -> List[SheetRow]:
    """DataFrame → SheetRow；excel_row = index + 2（表头第 1 行）。"""
    rows: List[SheetRow] = []
    for idx, series in df.iterrows():
        values: Dict[str, str] = {}
        for c in df.columns:
            key = normalize_col_name(c)
            raw = series[c]
            if bool_as_true_false and isinstance(raw, bool):
                values[key] = "true" if raw else "false"
            else:
                values[key] = clean_cell(raw)
        rows.append(SheetRow(excel_row=int(idx) + 2, values=values))
    return rows


def format_api_error(exc: BaseException) -> str:
    """从 HyOmsError.raw 等结构提取可读错误文案。"""
    raw = getattr(exc, "raw", None)
    if isinstance(raw, dict):
        err = raw.get("Error")
        if isinstance(err, dict):
            for key in ("errMessage", "message", "errMsg"):
                text = clean_cell(err.get(key))
                if text:
                    return text
        for key in ("message", "errMessage", "ask"):
            text = clean_cell(raw.get(key))
            if text:
                return text
    return str(exc)


def wanted_set(items: Optional[Sequence[str]]) -> set[str]:
    return {clean_cell(x) for x in (items or []) if clean_cell(x)}
