import os
import glob
import warnings
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
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name, ku_cun_date
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

# 忽略特定的警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")

# TODO 文件夹路径！！！
folder_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\鸿羽"
# 获取文件夹中的所有 ‘当前日期.xlsx’ 文件
file_paths = glob.glob(os.path.join(folder_path, f'*{shared_date}.xlsx'))
# 合并-仓租文件表格
hy_df = pd.concat([pd.read_excel(path, sheet_name='bizWarehouseRentByMonthDetail') for path in file_paths],
                  ignore_index=True)

hy_df['产品代码（SKU）'] = hy_df['产品代码（SKU）'].str.replace(r'^900008-', '', regex=True)
# # 指定要移除的尾缀内容
suffixes_to_remove = ['KA', 'JI', 'CH', 'DA', 'FB', 'AT', 'C1', 'C2', 'C3', 'ECO', 'REAL', 'ES', '4PX', 'UMI', 'MF',
                      'KL', 'YES', 'ML', 'ZSJ', 'ZJF']
regex_pattern = r'-(?:' + '|'.join(suffixes_to_remove) + r')$'
hy_df['产品代码（SKU）'] = hy_df['产品代码（SKU）'].str.replace(regex_pattern, '', regex=True)

# 按照 '产品代码（SKU）' 分组并计算 '产品金额（Product amount）' 总和
hy_df = hy_df.groupby('产品代码（SKU）', as_index=False)['产品金额（Product amount）'].sum()
hy_df = hy_df.rename(columns={'产品金额（Product amount）': '总仓租'})
hy_df['总仓租'] = hy_df['总仓租']
hy_all_cang_zu = hy_df['总仓租'].sum()  # HY 总的仓租费用

# 映射 HY仓租-sku（儿子）
product_map_sku_path = fr"{DESKTOP_ROOT}\仓租-SKU映射.xlsx"  # 改成对应的映射表
hy_df_1 = sku_mappings(
    main_df=hy_df,
    main_sku='产品代码（SKU）',
    map_sku_path=product_map_sku_path,
    map_old_sku="SKU",
    map_new_sku="映射SKU",
    map_sku_sheet='HY仓租-sku'
)
hy_df_1 = hy_df_1.rename(columns={'映射映射SKU': 'SKU'})

# 映射 商品ID
product_map_sku_path = r"\\Betohow\数据报表\数据库\产品信息库2025.xlsx"  # 改成对应的映射表
hy_df_2 = sku_mappings(
    main_df=hy_df_1,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="产品编码",
    map_new_sku="商品ID",
    map_sku_sheet='产品信息表'
)

hy_df_2 = hy_df_2.rename(columns={'映射商品ID': '商品ID'})

# 保存文件
output_file_path = file_paths[0].rsplit('\\', 1)[0] + '\\(已完成-1)hy-仓租明细.xlsx'
hy_df_2.to_excel(output_file_path, index=False)
print(f"处理完成，结果已保存到{output_file_path}")

# TODO 文件路径！！！
ku_cun_path = fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\{ku_cun_date}库存周转明细.xlsx'
ku_cun_df = pd.read_excel(ku_cun_path, sheet_name='各平台SKU库存周转明细')
# 计算每个 商品ID 的总数量
sku_total_quantity = ku_cun_df.groupby('商品ID')['在库（可调拨）'].transform('sum')
# 计算每个 商品ID 在每个平台的数量占比
ku_cun_df['数量占比'] = ku_cun_df['在库（可调拨）'] / sku_total_quantity
# 合并hy_df和ku_cun_df，根据 商品ID 进行合并
DF = pd.merge(hy_df_2, ku_cun_df, on='商品ID', how='left')
# 计算每个 商品ID 在每个平台的仓租
DF['仓租'] = DF['总仓租'] * DF['数量占比']
# ’平台‘列中值为“无”或“其它”的行，并将对应的仓租列的值改为0
DF.loc[DF['平台'].isin(['无', '其他']), '仓租'] = 0
# 选择需要的列生成新的DataFrame
result_DF = DF[['SKU', '商品ID', '平台', '仓租']].copy()

hy_have_site_cang_zu = result_DF['仓租'].sum()  # HY 有平台（站点）的仓租
# 需要分摊没有平台（站点）的仓租
hy_no_site_fen_tan = hy_all_cang_zu - hy_have_site_cang_zu

result_DF.loc[:, '无平台-需要分摊的费用'] = None  # 使用 .loc 来确保修改原始 DataFrame
# 在第一行的 新建一列写入数据
result_DF.at[0, '无平台-需要分摊的费用'] = hy_no_site_fen_tan
# 映射 站点
product_map_sku_path = fr'{DESKTOP_ROOT}\仓租-站点映射.xlsx'  # 改成对应的映射表
result_DF_1 = sku_mappings(
    main_df=result_DF,
    main_sku='平台',
    map_sku_path=product_map_sku_path,
    map_old_sku="平台",
    map_new_sku="站点",
    map_sku_sheet='Sheet1'
)
# 在 映射站点 后插入新列 站点商品ID识别码
new_column_name = "站点商品ID识别码"  # 新列名
new_column_data = result_DF_1["映射站点"] + result_DF_1["商品ID"]  # 新列数据
target_column = "映射站点"  # 目标列名（在其后插入）
insert_position = result_DF_1.columns.get_loc(target_column) + 1  # 计算插入位置
result_DF_1.insert(insert_position, new_column_name, new_column_data)  # 插入新列

# 映射 平台 （原表的 平台列 里面有站点！）
product_map_sku_path = fr'{DESKTOP_ROOT}\站点-匹配表.xlsx'  # 改成对应的映射表
result_DF_1 = sku_mappings(
    main_df=result_DF_1,
    main_sku='平台',
    map_sku_path=product_map_sku_path,
    map_old_sku="站点",
    map_new_sku="平台",
    map_sku_sheet='站点匹配'
)
# 重命名
result_DF_1 = result_DF_1.rename(columns={'平台': '原-平台'})
result_DF_1 = result_DF_1.rename(columns={'映射平台': '平台'})
# 在 站点商品ID识别码 后插入 平台商品ID识别码
new_column_name = "平台商品ID识别码"  # 新列名
new_column_data = result_DF_1["平台"] + result_DF_1["商品ID"]  # 新列数据
target_column = "站点商品ID识别码"  # 目标列名（在其后插入）
insert_position = result_DF_1.columns.get_loc(target_column) + 1  # 计算插入位置
result_DF_1.insert(insert_position, new_column_name, new_column_data)  # 插入新列

result_DF_1 = result_DF_1.rename(columns={'映射站点': '站点'})
result_DF_1 = result_DF_1.rename(columns={'仓租': '海外仓仓租费'})

# 保存文件
output_file_path = file_paths[0].rsplit('\\', 1)[0] + '\\(处理完成)hy-仓租明细.xlsx'
result_DF_1.to_excel(output_file_path, index=False)
print(f"处理完成，结果已保存到{output_file_path}")
