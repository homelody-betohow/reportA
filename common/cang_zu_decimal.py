"""
海外仓仓租金额精度：计算与写出统一保留 4 位小数。
"""

from __future__ import annotations

from typing import Iterable

import pandas as pd

RENT_DECIMALS = 4


def round_rent(value: object) -> float | object:
    """标量仓租金额 → 4 位小数；空值原样返回。"""
    if value is None:
        return value
    try:
        if pd.isna(value):
            return value
    except (TypeError, ValueError):
        pass
    return round(float(value), RENT_DECIMALS)


def round_rent_series(series: pd.Series) -> pd.Series:
    """Series 仓租金额 → 4 位小数（非数值→NaN 再 round）。"""
    return pd.to_numeric(series, errors="coerce").round(RENT_DECIMALS)


def round_rent_columns(df: pd.DataFrame, cols: Iterable[str]) -> pd.DataFrame:
    """就地式拷贝：对存在的金额列 round 到 4 位。"""
    out = df.copy()
    for col in cols:
        if col in out.columns:
            out[col] = round_rent_series(out[col])
    return out
