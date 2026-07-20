import numpy as np
import pandas as pd
import importlib.util
from pathlib import Path
# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.Z_method.platform_shop import map_region_to_platform
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name, USD_to_EUR
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！
file_path_1 = fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\广告\DLZ\(处理完成)DLZ-总的-广告数据.xlsx'
df1 = pd.read_excel(file_path_1, sheet_name='分摊明细-美元')

file_path_2 = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\广告\DLZ\(处理完成)DLZ广告-shopping-美元.xlsx"
df2 = pd.read_excel(file_path_2)

# 确保列名正确
df1.rename(columns={'SKU-站点识别码': 'SKU-站点识别码', '费用': '费用', 'SKU': 'SKU'}, inplace=True)
df2.rename(columns={'SKU-站点识别码': 'SKU-站点识别码', '费用': '费用', 'SKU': 'SKU'}, inplace=True)

# 合并数据
merged_df = pd.merge(df1, df2, on='SKU-站点识别码', how='left', suffixes=('_df1', '_df2'))

# 处理合并后的SKU，优先使用df1的SKU，如果为空则使用df2的
merged_df['SKU'] = merged_df['SKU_df1'].fillna(merged_df['SKU_df2'])

# 计算总的广告花费
merged_df['广告花费-美元'] = merged_df['费用_df1'].fillna(0) + merged_df['费用_df2'].fillna(0)

# 创建结果 DataFrame
result_df = merged_df[['SKU-站点识别码', '站点_df1', 'SKU', '广告花费-美元']].copy()  # 显式创建副本

# 重命名列
result_df = result_df.rename(columns={'站点_df1': '站点'})

# 检查未匹配的行
unmatched_df = df2[~df2['SKU-站点识别码'].isin(df1['SKU-站点识别码'])]

# 如果有未匹配的行，将它们添加到结果中
if not unmatched_df.empty:
    unmatched_df = unmatched_df.copy().rename(columns={'费用': '广告花费-美元'})
    unmatched_df = unmatched_df[['SKU-站点识别码', '站点', 'SKU', '广告花费-美元']]
    result_df = pd.concat([result_df, unmatched_df], ignore_index=True)

# 添加广告费(非AMZ)列                  美元 转 欧元
result_df['广告费(非AMZ)'] = np.round(result_df['广告花费-美元'] * USD_to_EUR, 2)

# 映射 平台（数据源：platform_shop）
result_df_1 = map_region_to_platform(result_df, site_col='站点')

# 去除 整张表 的前后空格
for col in result_df_1.columns:
    result_df_1[col] = result_df_1[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

# 在 SKU-站点识别码 后插入 SKU-平台识别码
new_column_name = "SKU-平台识别码"  # 新列名
new_column_data = result_df_1["映射平台"] + result_df_1["SKU"]  # 新列数据
target_column = "SKU-站点识别码"  # 目标列名（在其后插入）
insert_position = result_df_1.columns.get_loc(target_column) + 1  # 计算插入位置
result_df_1.insert(insert_position, new_column_name, new_column_data)  # 插入新列

# 保存目标列
result_df_1 = result_df_1[
    ['SKU', '站点', '映射平台', 'SKU-站点识别码', 'SKU-平台识别码', '广告花费-美元', '广告费(非AMZ)']]

# 筛选“广告费(非AMZ)”列不等于 0 的行
filtered_df = result_df_1[result_df_1['广告费(非AMZ)'] != 0]

# 检查 df1 是否有指定的列
if '无销量的站点' in df1.columns and '需要摊分花费（美元）' in df1.columns:
    # 筛选出需要摊分的行
    allocation_df = df1[df1['无销量的站点'].notna() & df1['需要摊分花费（美元）'].notna()]

    # 按站点分组处理
    for site, group in allocation_df.groupby('无销量的站点'):
        if site == 'DLZ-DE':
            # 计算需要摊分的总金额
            total_allocation = group['需要摊分花费（美元）'].sum()
            # 找到 DLZ-DE 的行
            dlz_de_rows = result_df_1[result_df_1['站点'] == 'DLZ-DE']
            # 如果有匹配的行，均匀分摊
            if not dlz_de_rows.empty:
                allocation_per_row = total_allocation / len(dlz_de_rows)
                result_df_1.loc[result_df_1['站点'] == 'DLZ-DE', '广告花费-美元'] += allocation_per_row
                result_df_1.loc[result_df_1['站点'] == 'DLZ-DE', '广告费(非AMZ)'] = np.round(
                    result_df_1.loc[result_df_1['站点'] == 'DLZ-DE', '广告花费-美元'] * USD_to_EUR, 2)

        elif site == 'DLZ-ES':
            # 计算需要摊分的总金额
            total_allocation = group['需要摊分花费（美元）'].sum()
            # 找到 DLZ-ES 的行
            dlz_es_rows = result_df_1[result_df_1['站点'] == 'DLZ-ES']
            # 如果有匹配的行，均匀分摊
            if not dlz_es_rows.empty:
                allocation_per_row = total_allocation / len(dlz_es_rows)
                result_df_1.loc[result_df_1['站点'] == 'DLZ-ES', '广告花费-美元'] += allocation_per_row
                result_df_1.loc[result_df_1['站点'] == 'DLZ-ES', '广告费(非AMZ)'] = np.round(
                    result_df_1.loc[result_df_1['站点'] == 'DLZ-ES', '广告花费-美元'] * USD_to_EUR, 2)




        elif site == 'DLZ-FR':
            # 订单统计中，无DLZ-FR销量时，则分摊到指定的SKU

            sku_file_path = fr'{DESKTOP_ROOT}\DLZ-FR_广告分摊sku.xlsx'

            sku_df = pd.read_excel(sku_file_path)

            # 确保SKU列没有重复值
            sku_df.drop_duplicates(subset=['SKU'], inplace=True)

            # 计算需要摊分的总金额
            total_allocation = group['需要摊分花费（美元）'].sum()

            # 检查sku_df中是否有必要的列
            required_columns = ['SKU', '站点', '映射平台', 'SKU-站点识别码', 'SKU-平台识别码']
            if not all(col in sku_df.columns for col in required_columns):
                raise ValueError(f"文件 {sku_file_path} 中缺少必要的列，请检查文件格式。")

            # 从sku_df中提取需要的列，并确保列的顺序
            sku_df = sku_df[required_columns].copy()

            # 将sku_df的数据直接加在result_df_1的下面
            result_df_1 = pd.concat([result_df_1, sku_df], ignore_index=True)

            # 找到 DLZ-FR 的行
            dlz_fr_rows = result_df_1[result_df_1['站点'] == 'DLZ-FR']

            # 如果有匹配的行，均匀分摊
            if not dlz_fr_rows.empty:
                allocation_per_row = total_allocation / len(dlz_fr_rows)

                # 更新DLZ-FR站点的广告花费
                result_df_1.loc[result_df_1['站点'] == 'DLZ-FR', '广告花费-美元'] = allocation_per_row

                # 更新DLZ-FR站点的广告费(非AMZ)
                result_df_1.loc[result_df_1['站点'] == 'DLZ-FR', '广告费(非AMZ)'] = np.round(
                    allocation_per_row * USD_to_EUR, 2)
# 删除 广告花费-美元 列值为 0 的行
result_df_1 = result_df_1[result_df_1['广告花费-美元'] != 0]
# 将结果写入文件
sheet_name = '广告花费-欧元'
with pd.ExcelWriter(file_path_1, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    result_df_1.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"操作完成，file_path：{file_path_1}，sheet_name：{sheet_name}")
