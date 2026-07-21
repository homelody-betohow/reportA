import numpy as np
import pandas as pd
import importlib.util
from pathlib import Path
# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from config.A0_set_date import shared_date, folder_name
from config.A0_paths import DESKTOP_ROOT


def merge_excel_files(file_paths):
    """
    合并多个Excel文件中的指定列。
    :param file_paths: 包含Excel文件路径的列表
    """
    # 初始化一个空的列表，用于存储每个文件的DataFrame
    all_data = []

    # 遍历文件路径列表
    for file_path in file_paths:
        # 读取Excel文件，只提取指定的几列
        if 'DLZ' in file_path:
            df = pd.read_excel(file_path, sheet_name='广告花费-欧元',
                               usecols=['SKU', '站点', '映射平台', 'SKU-站点识别码', 'SKU-平台识别码',
                                        '广告费(非AMZ)'])
        else:
            df = pd.read_excel(file_path,
                               usecols=['SKU', '站点', '映射平台', 'SKU-站点识别码', 'SKU-平台识别码',
                                        '广告费(非AMZ)'])
        # 将读取的数据添加到列表中
        all_data.append(df)
        print(f"成功读取文件：{file_path}")

    # 合并所有数据
    merged_data_df = pd.concat(all_data, ignore_index=True)
    return merged_data_df


# 4个平台处理好的广告费用文件路径
# TODO 文件路径！！！
file_paths = [
    fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\广告\OTTO\(处理完成)OTTO广告.xlsx',
    fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\广告\REAL\(处理完成)REAL广告.xlsx',
    fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\广告\MANO\(处理完成)MANO广告.xlsx',
    fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\广告\DLZ\(处理完成)DLZ-总的-广告数据.xlsx'
]
merged_df = merge_excel_files(file_paths)

# 去除 整张表 的前后空格
for col in merged_df.columns:
    merged_df[col] = merged_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

# SKU、SKU-站点识别码、SKU-平台识别码 去掉尾缀，去掉前后空格
cols = ['SKU', 'SKU-站点识别码', 'SKU-平台识别码']
# 正则规则：匹配末尾 -数字 或 -AT 或 --5
suffix_pattern = r'(-\d+|-AT|--5)$'

for col in cols:
    # strip去除前后空格 + 正则剔除指定尾部后缀
    merged_df[col] = merged_df[col].astype(str).str.strip().str.replace(suffix_pattern, '', regex=True)


merged_df = merged_df.rename(columns={'映射平台': '平台'})
# 按照 'SKU-站点识别码' 列进行分组，并对 '广告费(非AMZ)' 列进行汇总
grouped_df = merged_df.groupby('SKU-站点识别码').agg({
    '广告费(非AMZ)': 'sum',  # 汇总 广告费(非AMZ)
    'SKU': 'first',  # 保留每组的第一行数据
    '站点': 'first',  # 保留每组的第一行数据
    '平台': 'first',  # 保留每组的第一行数据
    'SKU-平台识别码': 'first'  # 保留每组的第一行数据
}).reset_index()
# 保留 2为小数
grouped_df['广告费(非AMZ)'] = np.round(grouped_df['广告费(非AMZ)'], 2)
# 筛选出 广告费(非AMZ) 列 中值不等于 0 的行（相对应删掉=0的行）
grouped_df = grouped_df[grouped_df['广告费(非AMZ)'] != 0]

grouped_df = grouped_df[
    ['SKU', '站点', '平台', 'SKU-站点识别码', 'SKU-平台识别码', '广告费(非AMZ)']]
# 将所有结果写入
output_path = file_paths[0].rsplit('\\', 2)[0] + '\\(处理完成)所有平台广告费用.xlsx'  # 合并后的文件名
grouped_df.to_excel(output_path, index=False)
print(f'所有平台广告费用文件合并成功，路径：{output_path}')
