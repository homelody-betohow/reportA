import re
import pandas as pd


def _cell_to_str(val):
    """Excel/CSV 可能把纯数字 SKU 读成 int/float，后续 re.split 与字符串拼接需要 str。"""
    if pd.isna(val):
        return val
    if isinstance(val, float) and val == int(val):
        return str(int(val))
    return str(val).strip()


def split_one_rows_data(input_df, data_column, value_column):
    """
    使用 Pandas 处理 Excel 文件，将指定列中包含 '+' 或 ',' 的数据拆分为多行，
    并将另一列的数据平均分配到这几行，其余列的数据进行复制。
    :param input_df: 输入的 Excel 文件的pd
    :param data_column: 包含 '+' 或 ',' 的数据所在的列名
    :param value_column: 需要平均分配的值所在的列名
    :return: 处理后的 DataFrame
    """
    input_df = input_df.copy()
    input_df[data_column] = input_df[data_column].apply(_cell_to_str)
    output_df = pd.DataFrame(columns=input_df.columns)

    for _, row in input_df.iterrows():
        cell_data = row[data_column]
        cell_value = row[value_column]

        if pd.notna(cell_data):
            # 使用正则拆分，支持 '+' 或 ','
            split_data = re.split(r'[+,]', cell_data)
            split_data = [part.strip() for part in split_data if part.strip()]
            num_parts = len(split_data)

            if num_parts > 1:
                avg_value = float(cell_value) / num_parts if pd.notna(cell_value) else None

                for part in split_data:
                    new_row = row.copy()
                    new_row[data_column] = part
                    new_row[value_column] = avg_value
                    output_df = pd.concat([output_df, new_row.to_frame().T], ignore_index=True)
                continue

        # 无需拆分，直接复制
        output_df = pd.concat([output_df, row.to_frame().T], ignore_index=True)

    return output_df
