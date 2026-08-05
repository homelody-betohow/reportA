import re
import pandas as pd


def _cell_to_str(val):
    """Excel/CSV 可能把纯数字 SKU 读成 int/float，后续 re.split 与字符串拼接需要 str。"""
    if pd.isna(val):
        return val
    if isinstance(val, float) and val == int(val):
        return str(int(val))
    return str(val).strip()


def strip_hash_suffix(val):
    """去掉 SKU 中 # 及之后的字母尾缀，如 E54047001#FBDE → E54047001。"""
    if pd.isna(val):
        return val
    return str(val).strip().split('#')[0]


def extract_internal_sku(s):
    """从原始 sku 字符串中提取内部产品编码。

    含 ``amzn.gr.`` / ``AMZN.GR.``（不区分大小写）时：取前缀后第一段，再按 ``-`` / ``_`` 截断。
    例：``AMZN.GR.U56033002-3NTXWVKLD7LRM8RCCY3-LN`` → ``U56033002``

    否则返回去首尾空格后的原字符串；空值原样返回。
    """
    if pd.isna(s):
        return s
    text = str(s).strip()
    lower = text.lower()
    marker = "amzn.gr."
    if marker in lower:
        rest = text[lower.index(marker) + len(marker) :]
        code = rest.split("-")[0].split("_")[0].strip()
        if code:
            return code
    return text


def split_one_rows_data(input_df, data_column, value_column, sync_columns=None):
    """
    使用 Pandas 处理 Excel 文件，将指定列中包含 '+' 或 ',' 的数据拆分为多行，
    并将金额列平均分配到这几行，其余列的数据进行复制。

    :param input_df: 输入的 DataFrame
    :param data_column: 包含 '+' 或 ',' 的数据所在的列名
    :param value_column: 需要平均分配的值所在的列名（str 或列名列表）
    :param sync_columns: 拆分后同步为子 SKU 的列名列表（如 仓库sku）
    :return: 处理后的 DataFrame
    """
    input_df = input_df.copy()
    input_df[data_column] = input_df[data_column].apply(_cell_to_str)
    value_columns = [value_column] if isinstance(value_column, str) else list(value_column)
    sync_columns = list(sync_columns or [])
    output_df = pd.DataFrame(columns=input_df.columns)

    for _, row in input_df.iterrows():
        cell_data = row[data_column]

        if pd.notna(cell_data):
            # 使用正则拆分，支持 '+' 或 ','
            split_data = re.split(r'[+,]', cell_data)
            split_data = [part.strip() for part in split_data if part.strip()]
            num_parts = len(split_data)

            if num_parts > 1:
                for part in split_data:
                    new_row = row.copy()
                    new_row[data_column] = part
                    for col in value_columns:
                        cell_value = row[col]
                        if pd.notna(cell_value):
                            try:
                                new_row[col] = float(cell_value) / num_parts
                            except (TypeError, ValueError):
                                new_row[col] = cell_value
                        else:
                            new_row[col] = None
                    for col in sync_columns:
                        synced = part
                        if col in ('仓库sku', '仓库SKU'):
                            synced = strip_hash_suffix(part)
                        new_row[col] = synced
                    output_df = pd.concat([output_df, new_row.to_frame().T], ignore_index=True)
                continue

        # 无需拆分，直接复制
        output_df = pd.concat([output_df, row.to_frame().T], ignore_index=True)

    return output_df
