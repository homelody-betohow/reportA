"""将 OMS 接口结果落成 Excel，便于接入现有报表流水线。"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Iterable, List, Mapping, Optional, Sequence, Union

PathLike = Union[str, Path]


def rows_to_dataframe(rows: Sequence[Mapping[str, Any]]):
    import pandas as pd

    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(list(rows))


def export_rows_to_excel(
    rows: Sequence[Mapping[str, Any]],
    path: PathLike,
    *,
    sheet_name: str = "Sheet1",
    columns: Optional[Sequence[str]] = None,
) -> Path:
    """把 dict 列表写成 xlsx；目录不存在时自动创建。"""
    import pandas as pd

    out = Path(path)
    out.parent.mkdir(parents=True, exist_ok=True)
    df = rows_to_dataframe(rows)
    if columns:
        for col in columns:
            if col not in df.columns:
                df[col] = None
        df = df.loc[:, list(columns)]
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return out


def export_inventory_storage_like_rent_sheet(
    rows: Iterable[Mapping[str, Any]],
    path: PathLike,
    *,
    sheet_name: str = "bizWarehouseRentByMonthDetail",
) -> Path:
    """将 ``getWhInventoryStorage`` 结果导出为接近仓租明细的工作表。

    注意：快照接口无「产品金额」字段，金额列留空，供后续手工或费用接口补齐。
    """
    mapped: List[dict] = []
    for row in rows:
        mapped.append(
            {
                "产品代码（SKU）": row.get("sku"),
                "产品名称": row.get("productName"),
                "产品金额（Product amount）": row.get("amount") or row.get("产品金额（Product amount）"),
                "数量": row.get("quantity"),
                "总体积": row.get("totalVolume"),
                "库龄": row.get("libraryOfAge"),
                "计费日期": row.get("chargeDate"),
                "入库单号": row.get("receivingOrder"),
                "仓租类型": row.get("storageType"),
                "计费类型": row.get("billingType"),
            }
        )
    return export_rows_to_excel(mapped, path, sheet_name=sheet_name)
