import pandas as pd
import importlib.util
from pathlib import Path
# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.Z_method.sku_映射 import sku_mappings
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

# 2个平台处理好的仓租用文件路径
# TODO 文件路径！！！
file_paths = [
    fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\4PX\(处理完成)4PX-仓租明细.xlsx',
    fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\鸿羽\(处理完成)hy-仓租明细.xlsx'
]
all_data = []
# 遍历文件路径列表
for file_path in file_paths:
    df = pd.read_excel(file_path)
    # 将读取的数据添加到列表中
    all_data.append(df)
    print(f"成功读取文件：{file_path}")

# 合并所有数据
merged_df = pd.concat(all_data, ignore_index=True)
# 重命名列
merged_df = merged_df.rename(columns={'SKU': '原-SKU'})
# 计算 LM-BC 的海外仓仓租费总值
lm_bc_sum = merged_df[merged_df['原-平台'] == 'LM-BC']['海外仓仓租费'].sum()
merged_df['LM-BC的仓租'] = float('nan')  # 新建一列，默认填充 NaN
merged_df.loc[1, 'LM-BC的仓租'] = lm_bc_sum  # 在第二行（索引为1）填入总值

# 计算-一共需要分摊的仓租
all_fen_tan_cang_zu = merged_df['无平台-需要分摊的费用'].sum()  # 所有仓库 需要分摊的仓租

# 按照 '站点商品ID识别码' 列进行分组，并对 '仓租' 列进行汇总
# 空的'站点商品ID识别码'的数据会丢失，原-平台 == LM-BC 的数据会丢失
result_df = merged_df.groupby('站点商品ID识别码').agg({
    '海外仓仓租费': 'sum',  # 求和
    '无平台-需要分摊的费用': 'sum',  # 求和
    '原-SKU': 'first',  # 保留每组的第一行数据
    '商品ID': 'first',  # 保留每组的第一行数据
    '站点': 'first',  # 保留每组的第一行数据
    '平台': 'first',  # 保留每组的第一行数据
    '平台商品ID识别码': 'first'  # 保留每组的第一行数据
}).reset_index()

#  商品ID 去映射 产品信息库 的 第一个 产品编码（SKU）
product_map_sku_path = r"\\Betohow\数据报表\数据库\产品信息库2025.xlsx"  # 改成对应的映射表
result_df_1 = sku_mappings(
    main_df=result_df,
    main_sku='商品ID',
    map_sku_path=product_map_sku_path,
    map_old_sku="商品ID",
    map_new_sku="产品编码",
    map_sku_sheet='产品信息表'
)
# 重命名列
result_df_1 = result_df_1.rename(columns={'映射产品编码': 'SKU'})

result_df_1 = result_df_1[
    ['SKU', '商品ID', '站点', '平台', '站点商品ID识别码', '平台商品ID识别码', '海外仓仓租费']]

# 筛选“站点”列不为空、不为空字符串、不等于“无”、不等于“其它”，海外仓仓租费 不等于 0
filtered_df = result_df_1[result_df_1['站点'].notna() & (result_df_1['站点'] != '') & (result_df_1['站点'] != '无') & (
        result_df_1['站点'] != '其他') & (result_df_1['海外仓仓租费'] != 0)]

filtered_df = filtered_df.copy()
# 新增空列：所有仓库-无平台-需要分摊的费用
filtered_df['所有仓库-无平台-需要分摊的费用'] = pd.NA
# 将 求和结果，放在新增列的第一个单元格
filtered_df.iloc[0, filtered_df.columns.get_loc('所有仓库-无平台-需要分摊的费用')] = all_fen_tan_cang_zu

# 新增空列：所有仓库-无平台-需要分摊的费用
filtered_df['LM-BC的仓租'] = pd.NA
# 将 求和结果，放在新增列的第一个单元格
filtered_df.iloc[0, filtered_df.columns.get_loc('LM-BC的仓租')] = lm_bc_sum

# 将所有结果写入
output_path = fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租' + '\\(处理完成)所有-海外仓-仓租明细.xlsx'
filtered_df.to_excel(output_path, index=False)
print(f'所有，平台仓租费用，文件合并成功，path：{output_path}')
