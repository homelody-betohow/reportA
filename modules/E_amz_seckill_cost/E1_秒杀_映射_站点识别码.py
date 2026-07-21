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

from common.platform_shop import map_region_to_platform
from config.A0_set_date import shared_date, folder_name
from common.split_rows_data_SKU import split_one_rows_data
from config.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\秒杀\秒杀数据-{shared_date}.xlsx"
# 读取秒杀费用文件的指定sheet
main_df = pd.read_excel(main_file_path)
# 重命名
main_df = main_df.rename(columns={'sku': 'SKU'})
#  拆分有“+”的  映射仓库sku
main_df_1 = split_one_rows_data(
    input_df=main_df,
    data_column='SKU',
    value_column='秒杀费'
)
# 使用字符串操作合并两列
main_df_1['SKU-站点识别码'] = main_df_1['站点'] + main_df_1['SKU']
# 映射 平台（数据源：platform_shop）
main_df_2 = map_region_to_platform(main_df_1, site_col='站点')
# 在 SKU-站点识别码 后插入 SKU-平台识别码
new_column_name = "SKU-平台识别码"  # 新列名
new_column_data = main_df_2["映射平台"] + main_df_2["SKU"]  # 新列数据
target_column = "SKU-站点识别码"  # 目标列名（在其后插入）
insert_position = main_df_2.columns.get_loc(target_column) + 1  # 计算插入位置
main_df_2.insert(insert_position, new_column_name, new_column_data)  # 插入新列

main_df_2 = main_df_2.rename(columns={'映射平台': '平台'})

# 按照 'SKU-站点识别码' 列进行分组，并对 '秒杀费' 列进行汇总
grouped_main_df = main_df_2.groupby('SKU-站点识别码').agg({
    '秒杀费': 'sum',  # 汇总 秒杀费
    'SKU': 'first',  # 保留每组的第一行数据
    '站点': 'first',  # 保留每组的第一行数据
    '平台': 'first',  # 保留每组的第一行数据
    'SKU-平台识别码': 'first'  # 保留每组的第一行数据
}).reset_index()

# 强制转换为数值类型，无法转换的变为 NaN
grouped_main_df['秒杀费'] = pd.to_numeric(grouped_main_df['秒杀费'], errors='coerce')
grouped_main_df['秒杀费'] = np.round(grouped_main_df['秒杀费'], 2)
# 自定义保存列
grouped_main_df = grouped_main_df[['SKU', '站点', '平台', 'SKU-站点识别码', 'SKU-平台识别码', '秒杀费']]

# 将结果保存到新的Excel文件
output_file_path = main_file_path.rsplit('\\', 1)[0] + '\\(处理完成)' + main_file_path.rsplit('\\', 1)[1]
grouped_main_df.to_excel(output_file_path, index=False)
print(f'处理完成，输出文件路径：{output_file_path}')
