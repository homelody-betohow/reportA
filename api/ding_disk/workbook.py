"""钉钉在线表格常用封装：按工作表读写 DataFrame、分块读取大表等。

底层 HTTP 见 ``DingDiskClient``；本模块提供面向业务脚本的高层 API。

示例::

    from api.ding_disk.workbook import Workbook

    wb = Workbook("Obva6QBXJwjBxoE2sM62MrzGVn4qY5Pr")
    sheets = wb.list_sheets()
    rows, cols = wb.used_size("Sheet1")   # 有效行/列（有值最大）
    df = wb.read_sheet("Sheet1")
    wb.write_dataframe("Sheet2", df)
"""

from __future__ import annotations

from typing import Any, Dict, List, Mapping, Optional, Sequence, Union

import pandas as pd

from .client import DingDiskClient, _col_to_a1, build_a1_range, parse_a1_cell
from .exceptions import DingDiskError

# 钉钉 ranges API：单次最多读写约 30000 个单元格
MAX_RANGE_CELLS = 30_000
# 钉钉 values 数组：单次最多 1000 行（items）
MAX_RANGE_ROWS = 1_000

JsonDict = Dict[str, Any]
CellMatrix = Sequence[Sequence[Any]]
SheetRef = Union[str, Dict[str, Any]]


def col_to_a1(col_index: int) -> str:
    """0-based 列索引 → A1 列标（0→A, 25→Z, 26→AA）。"""
    return _col_to_a1(col_index)


def used_range_bounds(sheet_meta: Dict[str, Any]) -> Optional[tuple[int, int]]:
    """从工作表元数据取已用区域 ``(last_row, last_col)``（0-based）；空表返回 None。"""
    try:
        last_row = int(sheet_meta.get("lastNonEmptyRow", -1))
        last_col = int(sheet_meta.get("lastNonEmptyColumn", -1))
    except (TypeError, ValueError):
        return None
    if last_row < 0 or last_col < 0:
        return None
    return last_row, last_col


def max_used_row(sheet_meta: Dict[str, Any]) -> int:
    """有值的最大行号（1-based，与 Excel 行号一致）；空表返回 0。

    对应钉钉 ``lastNonEmptyRow``（0-based）+ 1。
    """
    try:
        last_row = int(sheet_meta.get("lastNonEmptyRow", -1))
    except (TypeError, ValueError):
        return 0
    return last_row + 1 if last_row >= 0 else 0


def max_used_col(sheet_meta: Dict[str, Any]) -> int:
    """有值的最大列号（1-based，1=A、2=B…）；空表返回 0。

    对应钉钉 ``lastNonEmptyColumn``（0-based）+ 1。
    """
    try:
        last_col = int(sheet_meta.get("lastNonEmptyColumn", -1))
    except (TypeError, ValueError):
        return 0
    return last_col + 1 if last_col >= 0 else 0


def max_used_col_letter(sheet_meta: Dict[str, Any]) -> str:
    """有值的最大列的 A1 列标（如 ``C``）；空表返回空字符串。"""
    n = max_used_col(sheet_meta)
    if n <= 0:
        return ""
    return col_to_a1(n - 1)


def used_range_address(sheet_meta: Dict[str, Any]) -> Optional[str]:
    """根据工作表元数据拼出已用区域 ``A1:…``；空表返回 None。"""
    bounds = used_range_bounds(sheet_meta)
    if bounds is None:
        return None
    last_row, last_col = bounds
    return f"A1:{col_to_a1(last_col)}{last_row + 1}"

def sheet_name(item: SheetRef) -> str:
    """从 list_sheets 条目或字符串取出工作表名称/ID。"""
    if isinstance(item, str):
        return item.strip()
    name = str(item.get("name") or item.get("id") or "").strip()
    return name


def normalize_cell(val: Any) -> str:
    """单元格值规范为字符串：去空白；空/NaN → ``""``；浮点整型去 ``.0``。"""
    if val is None:
        return ""
    if isinstance(val, float) and pd.isna(val):
        return ""
    text = str(val).strip()
    if text.lower() in {"nan", "none", "null"}:
        return ""
    if isinstance(val, float) and text.endswith(".0"):
        text = text[:-2]
    return text


# strip() 默认不含的常见不可见字符（表格粘贴常带入）
_EDGE_INVISIBLE = (
    "\u200b"  # ZERO WIDTH SPACE
    "\u200c"  # ZERO WIDTH NON-JOINER
    "\u200d"  # ZERO WIDTH JOINER
    "\ufeff"  # BOM / ZERO WIDTH NO-BREAK SPACE
)


def clean_cell(val: Any) -> str:
    """写库/比对前清洗单元格：去两端空白与常见不可见字符。

    在 ``normalize_cell`` 基础上再剥零宽空格等，避免肉眼看不见的误输入入库。
    """
    text = normalize_cell(val)
    if not text:
        return ""
    while text and (text[0] in _EDGE_INVISIBLE or text[0].isspace()):
        text = text[1:]
    while text and (text[-1] in _EDGE_INVISIBLE or text[-1].isspace()):
        text = text[:-1]
    return text


def clean_pairs(pairs: Mapping[str, str]) -> Dict[str, str]:
    """清洗键值两端空白；空 key 丢弃，同 key 后者覆盖。"""
    cleaned: Dict[str, str] = {}
    for key, value in pairs.items():
        k = clean_cell(key)
        if not k:
            continue
        cleaned[k] = clean_cell(value)
    return cleaned


def require_columns(df: pd.DataFrame, cols: Sequence[str]) -> None:
    """确认 DataFrame 含指定列，否则抛 ``KeyError``。"""
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise KeyError(f"表格缺少列: {missing}；实际列={list(df.columns)}")


def filter_by_column(
    df: pd.DataFrame,
    column: str,
    values: Sequence[Any],
) -> tuple[pd.DataFrame, List[str]]:
    """按列值过滤；返回 ``(filtered_df, missing_values)``。

    比较前对单元格与目标值均做 ``clean_cell``。
    """
    require_columns(df, [column])
    wanted = {clean_cell(v) for v in values if clean_cell(v)}
    if not wanted:
        return df.iloc[0:0].copy(), []
    series = df[column].map(clean_cell)
    filtered = df.loc[series.isin(wanted)].copy()
    found = set(series[series.isin(wanted)].tolist())
    missing = sorted(wanted - found)
    return filtered, missing


def kv_pairs_from_df(
    df: pd.DataFrame,
    key_col: str,
    value_col: str,
    *,
    allow_empty: bool = False,
) -> tuple[Dict[str, str], int]:
    """从两列提取 ``{key: value}``（同 key 后者覆盖）。

    键值经 ``clean_cell`` 清洗。返回 ``(pairs, empty_value_skipped)``。
    空 key 始终跳过；``allow_empty=False`` 时跳过空 value。
    """
    require_columns(df, [key_col, value_col])
    keys = df[key_col].map(clean_cell)
    vals = df[value_col].map(clean_cell)
    mask_key = keys.ne("")
    empty_skipped = int((mask_key & vals.eq("")).sum()) if not allow_empty else 0
    keep = mask_key if allow_empty else (mask_key & vals.ne(""))
    pairs: Dict[str, str] = {}
    for key, value in zip(keys[keep].tolist(), vals[keep].tolist()):
        pairs[key] = value
    return pairs, empty_skipped


def matrix_to_dataframe(
    values: Sequence[Sequence[Any]],
    *,
    header: bool = True,
) -> pd.DataFrame:
    """二维单元格 → DataFrame。

    ``header=True``：首行作表头（空名补 ``col_N``，重名加后缀）。
    ``header=False``：全部为数据，列名为 ``col_1…``。
    """
    if not values:
        return pd.DataFrame()

    if not header:
        width = max((len(r) for r in values), default=0)
        cols = [f"col_{i + 1}" for i in range(width)]
        rows: List[List[Any]] = []
        for row in values:
            cells = list(row)[:width]
            if len(cells) < width:
                cells.extend([None] * (width - len(cells)))
            rows.append(cells)
        return pd.DataFrame(rows, columns=cols)

    header_row = [str(c).strip() if c is not None else "" for c in values[0]]
    seen: Dict[str, int] = {}
    columns: List[str] = []
    for i, name in enumerate(header_row):
        base = name or f"col_{i + 1}"
        if base in seen:
            seen[base] += 1
            columns.append(f"{base}_{seen[base]}")
        else:
            seen[base] = 0
            columns.append(base)

    width = len(columns)
    normalized: List[List[Any]] = []
    for row in values[1:]:
        cells = list(row)[:width]
        if len(cells) < width:
            cells.extend([None] * (width - len(cells)))
        normalized.append(cells)
    return pd.DataFrame(normalized, columns=columns)


def dataframe_to_matrix(
    df: pd.DataFrame,
    *,
    include_header: bool = True,
) -> List[List[Any]]:
    """DataFrame → 二维单元格（None 表示空单元格）。"""
    frame = df.where(pd.notnull(df), None)
    body = frame.values.tolist()
    if not include_header:
        return body
    return [list(frame.columns.astype(str))] + body


def get_values_chunked(
    client: DingDiskClient,
    sheet: str,
    workbook_id: str,
    *,
    last_row: int,
    last_col: int,
    max_cells: int = MAX_RANGE_CELLS,
) -> List[List[Any]]:
    """按行分块读取 ``A1:…``，避免单次超过钉钉单元格上限。"""
    nrows = last_row + 1
    ncols = last_col + 1
    if ncols > max_cells:
        raise DingDiskError(
            f"列数 {ncols} 超过单次上限 {max_cells}，无法按行分块读取"
        )
    rows_per_chunk = max(1, max_cells // ncols)
    col_end = col_to_a1(last_col)
    if nrows * ncols <= max_cells:
        return client.get_values(sheet, f"A1:{col_end}{nrows}", workbook_id)

    values: List[List[Any]] = []
    start = 0
    while start < nrows:
        end = min(start + rows_per_chunk, nrows)
        addr = f"A{start + 1}:{col_end}{end}"
        values.extend(client.get_values(sheet, addr, workbook_id))
        start = end
    return values


def _rows_per_write_chunk(ncols: int, *, max_cells: int = MAX_RANGE_CELLS) -> int:
    """单次写入行数：同时受单元格上限与 values 数组 1000 行限制。"""
    if ncols <= 0:
        raise ValueError(f"列数必须为正: {ncols}")
    if ncols > max_cells:
        raise DingDiskError(
            f"列数 {ncols} 超过单次上限 {max_cells}，无法按行分块写入"
        )
    by_cells = max(1, max_cells // ncols)
    return max(1, min(by_cells, MAX_RANGE_ROWS))


def write_column_updates(
    client: DingDiskClient,
    sheet: str,
    col_index: int,
    updates: Sequence[tuple[int, Any]],
    workbook_id: str,
    *,
    max_cells: int = MAX_RANGE_CELLS,
) -> int:
    """按列写入零散单元格更新。

    ``updates`` 为 ``(excel_row_1based, value)`` 列表；会合并连续行后分块写入。
    返回实际写入的单元格数。
    """
    if not updates:
        return 0
    if col_index < 0:
        raise ValueError(f"列索引不能为负: {col_index}")

    ordered = sorted(((int(r), v) for r, v in updates), key=lambda x: x[0])
    col_letter = col_to_a1(col_index)
    rows_limit = _rows_per_write_chunk(1, max_cells=max_cells)
    written = 0
    i = 0
    while i < len(ordered):
        start_row, val = ordered[i]
        chunk: List[List[Any]] = [[val]]
        j = i + 1
        while (
            j < len(ordered)
            and ordered[j][0] == ordered[j - 1][0] + 1
            and len(chunk) < rows_limit
        ):
            chunk.append([ordered[j][1]])
            j += 1
        write_values_chunked(
            client,
            sheet,
            chunk,
            workbook_id,
            start_cell=f"{col_letter}{start_row}",
            max_cells=max_cells,
        )
        written += len(chunk)
        i = j
    return written


def write_values_chunked(
    client: DingDiskClient,
    sheet: str,
    values: CellMatrix,
    workbook_id: str,
    *,
    start_cell: str = "A1",
    max_cells: int = MAX_RANGE_CELLS,
) -> None:
    """按行分块写入，避免单次超过钉钉单元格/行数上限。"""
    matrix = [list(row) for row in values]
    if not matrix:
        return
    ncols = max(len(r) for r in matrix)
    if ncols <= 0:
        return
    rows_per_chunk = _rows_per_write_chunk(ncols, max_cells=max_cells)
    start_row, start_col = parse_a1_cell(start_cell)
    offset = 0
    while offset < len(matrix):
        chunk = matrix[offset : offset + rows_per_chunk]
        # 对齐列宽，避免 range 计算偏差
        padded = [row + [None] * (ncols - len(row)) for row in chunk]
        cell = f"{col_to_a1(start_col)}{start_row + offset + 1}"
        client.write_values(sheet, padded, workbook_id, start_cell=cell)
        offset += rows_per_chunk


def append_rows_chunked(
    client: DingDiskClient,
    sheet: str,
    values: CellMatrix,
    workbook_id: str,
    *,
    max_cells: int = MAX_RANGE_CELLS,
) -> None:
    """按行分块追加，避免单次超过钉钉单元格/行数上限。"""
    matrix = [list(row) for row in values]
    if not matrix:
        return
    ncols = max((len(r) for r in matrix), default=0) or 1
    rows_per_chunk = _rows_per_write_chunk(ncols, max_cells=max_cells)
    for i in range(0, len(matrix), rows_per_chunk):
        client.append_rows(sheet, matrix[i : i + rows_per_chunk], workbook_id)


def list_sheets(
    workbook_id: str,
    *,
    client: Optional[DingDiskClient] = None,
) -> List[Dict[str, Any]]:
    """列出表格内全部工作表（id / name 等）。"""
    client = client or DingDiskClient.from_config()
    data = client.list_sheets(workbook_id)
    sheets = data.get("value") or []
    if not isinstance(sheets, list):
        raise DingDiskError(f"list_sheets 返回异常: {data}")
    return [s for s in sheets if isinstance(s, dict)]


def read_sheet(
    sheet: str,
    workbook_id: str,
    *,
    client: Optional[DingDiskClient] = None,
    header: bool = True,
) -> pd.DataFrame:
    """读取指定工作表（ID 或名称）为 DataFrame。"""
    client = client or DingDiskClient.from_config()
    meta = client.get_sheet(sheet, workbook_id)
    bounds = used_range_bounds(meta)
    if not bounds:
        return pd.DataFrame()
    last_row, last_col = bounds
    values = get_values_chunked(
        client,
        sheet,
        workbook_id,
        last_row=last_row,
        last_col=last_col,
    )
    return matrix_to_dataframe(values, header=header)


def read_all_sheets(
    workbook_id: str,
    *,
    client: Optional[DingDiskClient] = None,
    header: bool = True,
) -> Dict[str, pd.DataFrame]:
    """读取全部工作表，返回 ``{sheet_name: DataFrame}``。"""
    client = client or DingDiskClient.from_config()
    result: Dict[str, pd.DataFrame] = {}
    for item in list_sheets(workbook_id, client=client):
        name = sheet_name(item)
        if not name:
            continue
        result[name] = read_sheet(
            name,
            workbook_id,
            client=client,
            header=header,
        )
    return result


def write_dataframe(
    sheet: str,
    df: pd.DataFrame,
    workbook_id: str,
    *,
    client: Optional[DingDiskClient] = None,
    start_cell: str = "A1",
    include_header: bool = True,
) -> None:
    """将 DataFrame 写入工作表（大表自动分块）。"""
    client = client or DingDiskClient.from_config()
    values = dataframe_to_matrix(df, include_header=include_header)
    write_values_chunked(
        client,
        sheet,
        values,
        workbook_id,
        start_cell=start_cell,
    )


def append_dataframe(
    sheet: str,
    df: pd.DataFrame,
    workbook_id: str,
    *,
    client: Optional[DingDiskClient] = None,
    include_header: bool = False,
) -> None:
    """将 DataFrame 追加到工作表末尾（大表自动分块）。"""
    client = client or DingDiskClient.from_config()
    values = dataframe_to_matrix(df, include_header=include_header)
    append_rows_chunked(client, sheet, values, workbook_id)


class Workbook:
    """绑定 ``workbook_id`` 的便捷封装。"""

    def __init__(
        self,
        workbook_id: str,
        *,
        client: Optional[DingDiskClient] = None,
    ) -> None:
        wid = (workbook_id or "").strip()
        if not wid:
            raise ValueError("workbook_id 不能为空")
        self.workbook_id = wid
        self.client = client or DingDiskClient.from_config()

    @classmethod
    def from_client(
        cls,
        workbook_id: str,
        client: DingDiskClient,
    ) -> "Workbook":
        return cls(workbook_id, client=client)

    def list_sheets(self) -> List[Dict[str, Any]]:
        return list_sheets(self.workbook_id, client=self.client)

    def get_sheet(self, sheet: str) -> JsonDict:
        return self.client.get_sheet(sheet, self.workbook_id)

    def used_range(self, sheet: str) -> Optional[str]:
        return used_range_address(self.get_sheet(sheet))

    def max_used_row(self, sheet: str) -> int:
        """读取有效行：有值的最大行号（1-based）；空表返回 0。"""
        return max_used_row(self.get_sheet(sheet))

    def max_used_col(self, sheet: str) -> int:
        """读取有效列：有值的最大列号（1-based，1=A）；空表返回 0。"""
        return max_used_col(self.get_sheet(sheet))

    def max_used_col_letter(self, sheet: str) -> str:
        """读取有效列的 A1 列标（如 ``C``）；空表返回空字符串。"""
        return max_used_col_letter(self.get_sheet(sheet))

    def used_size(self, sheet: str) -> tuple[int, int]:
        """``(有效行数, 有效列数)``，均为有值最大的 1-based 序号；空表 ``(0, 0)``。"""
        meta = self.get_sheet(sheet)
        return max_used_row(meta), max_used_col(meta)

    def get_values(
        self,
        sheet: str,
        range_address: Optional[str] = None,
    ) -> List[List[Any]]:
        """读取区域 values；``range_address`` 为空时读已用区域（自动分块）。"""
        if range_address:
            # 调用方自管范围大小；超限时钉钉会报错
            return self.client.get_values(sheet, range_address, self.workbook_id)
        meta = self.get_sheet(sheet)
        bounds = used_range_bounds(meta)
        if not bounds:
            return []
        last_row, last_col = bounds
        return get_values_chunked(
            self.client,
            sheet,
            self.workbook_id,
            last_row=last_row,
            last_col=last_col,
        )

    def read_sheet(self, sheet: str, *, header: bool = True) -> pd.DataFrame:
        return read_sheet(
            sheet,
            self.workbook_id,
            client=self.client,
            header=header,
        )

    def read_all_sheets(self, *, header: bool = True) -> Dict[str, pd.DataFrame]:
        return read_all_sheets(
            self.workbook_id,
            client=self.client,
            header=header,
        )

    def write_values(
        self,
        sheet: str,
        values: CellMatrix,
        *,
        start_cell: str = "A1",
    ) -> None:
        write_values_chunked(
            self.client,
            sheet,
            values,
            self.workbook_id,
            start_cell=start_cell,
        )

    def write_column_updates(
        self,
        sheet: str,
        col_index: int,
        updates: Sequence[tuple[int, Any]],
    ) -> int:
        """按列写入 ``(excel_row_1based, value)`` 更新，自动合并连续行。"""
        return write_column_updates(
            self.client,
            sheet,
            col_index,
            updates,
            self.workbook_id,
        )

    def write_dataframe(
        self,
        sheet: str,
        df: pd.DataFrame,
        *,
        start_cell: str = "A1",
        include_header: bool = True,
    ) -> None:
        write_dataframe(
            sheet,
            df,
            self.workbook_id,
            client=self.client,
            start_cell=start_cell,
            include_header=include_header,
        )

    def append_rows(self, sheet: str, values: CellMatrix) -> None:
        append_rows_chunked(self.client, sheet, values, self.workbook_id)

    def append_dataframe(
        self,
        sheet: str,
        df: pd.DataFrame,
        *,
        include_header: bool = False,
    ) -> None:
        append_dataframe(
            sheet,
            df,
            self.workbook_id,
            client=self.client,
            include_header=include_header,
        )

    def clear_data(self, sheet: str, range_address: str) -> JsonDict:
        return self.client.clear_data(sheet, range_address, self.workbook_id)

    def clear_all(self, sheet: str, range_address: str) -> JsonDict:
        return self.client.clear_all(sheet, range_address, self.workbook_id)


__all__ = [
    "MAX_RANGE_CELLS",
    "MAX_RANGE_ROWS",
    "Workbook",
    "append_dataframe",
    "append_rows_chunked",
    "col_to_a1",
    "dataframe_to_matrix",
    "filter_by_column",
    "get_values_chunked",
    "kv_pairs_from_df",
    "list_sheets",
    "matrix_to_dataframe",
    "max_used_col",
    "max_used_col_letter",
    "max_used_row",
    "clean_cell",
    "clean_pairs",
    "normalize_cell",
    "read_all_sheets",
    "read_sheet",
    "require_columns",
    "sheet_name",
    "used_range_address",
    "used_range_bounds",
    "write_column_updates",
    "write_dataframe",
    "write_values_chunked",
    "build_a1_range",
    "parse_a1_cell",
]
