import os
import glob
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
from A_报表.Z_method.style import Color
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！
main_file_path = fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\二次上架\鸿羽仓-二次上架明细-{shared_date}.xls'
main_df = pd.read_excel(main_file_path)
# 不要“实收数量” 为 0 的
main_df = main_df[main_df['实收数量'] != 0]
# SKU、儿子-站点识别码、儿子-平台识别码 去掉尾缀 -1、-2、-3、-4、-5、-6、-7、-8、-AT、--5、-BC、-FB；去掉前后空格
main_df['SKU'] = main_df['SKU'].astype(str).str.strip().str.replace(r'((-[1-8]|-AT|--5|-BC|-FB|-BTL)+)$', '', regex=True)
# 合并 订单管理-查询结果的表
folder_path = r'\\Betohow\数据报表\RPA\二次上架-数据查询\订单管理'  # 替换为你的文件夹路径
# 获取文件夹中的所有 .csv 文件
file_list = glob.glob(os.path.join(folder_path, '*.csv'))
columns_to_keep = ['店铺账号', '销售参考号', 'SKU', '产品数量', '产品SKU']  # 合并表格，要保留的列名
# 初始化一个空的 DataFrame，用于存储合并后的数据
merged_df = pd.DataFrame(columns=columns_to_keep)
# 遍历每个文件，只保留指定的列，并追加到合并后的 DataFrame 中
for file in file_list:
    temp_df = pd.read_csv(file, low_memory=False)
    # 筛选出需要的列
    temp_df = temp_df[columns_to_keep]
    merged_df = pd.concat([merged_df, temp_df], ignore_index=True)

# 去除杂质，保留数据部分     将所有列转换为字符串类型后再处理
merged_df_cleaned = merged_df.astype(str).apply(lambda col: col.map(lambda x: x.strip('="')))
# 保存合并后的表格到新的文件
output_path = folder_path + '\\all-订单管理查询.xlsx'
merged_df_cleaned.to_excel(output_path, index=False)
print(f"表格合并完成，结果已保存到：{output_path}")

product_map_sku_path = output_path
# 映射 订单管理-查询结果的 “店铺账号”
main_df_1 = sku_mappings(
    main_df=main_df,
    main_sku='订单参考号',
    map_sku_path=product_map_sku_path,
    map_old_sku="销售参考号",
    map_new_sku="店铺账号",
    map_sku_sheet="Sheet1"
)

main_df_1 = main_df_1.rename(columns={'映射店铺账号': '映射店铺账号-1'})
# 映射 订单管理-查询结果的 “店铺账号”
main_df_2 = sku_mappings(
    main_df=main_df_1,
    main_sku='参考号',  # 有的“订单参考号”在 参考号 列
    map_sku_path=product_map_sku_path,
    map_old_sku="销售参考号",
    map_new_sku="店铺账号",
    map_sku_sheet="Sheet1"
)

main_df_2 = main_df_2.rename(columns={'映射店铺账号': '映射店铺账号-2'})

# 映射 订单管理-查询结果的 “店铺账号”
product_map_sku_path = r"\\Betohow\数据报表\RPA\二次上架-数据查询\自发货\自发货-订单查询.xlsx"
main_df_3 = sku_mappings(
    main_df=main_df_2,
    main_sku='参考号',  # 有的“订单号”在 参考号 列
    map_sku_path=product_map_sku_path,
    map_old_sku="服务号",
    map_new_sku="店铺账号",
    map_sku_sheet="Worksheet 1"
)
# 去除杂质，保留数据部分
main_df_3['SKU'] = main_df_3['SKU'].apply(
    lambda x: x.replace('900008-', '').replace('-ECO', '').replace('-BC', '').replace('-AT-01', '').strip())
# 将映射的 店铺账号 合并成一列
main_df_3['合并-映射账号'] = main_df_3['映射店铺账号'].combine_first(main_df_3['映射店铺账号-1']).combine_first(
    main_df_3['映射店铺账号-2'])
# 去掉 括号的内容
main_df_3['合并-映射账号'] = main_df_3['合并-映射账号'].apply(
    lambda x: x.split('(')[0].strip() if x and '(' in x else x)

# 替换 相同产品的SKU  避免映射不到  原始采购价
main_df_3.loc[main_df_3['SKU'] == '20007-YES', 'SKU'] = 'SK20007'

product_map_sku_path = fr"{DESKTOP_ROOT}\站点-匹配表.xlsx"  # 改成对应的映射表
# 映射   站点
main_df_4 = sku_mappings(
    main_df=main_df_3,
    main_sku='合并-映射账号',
    map_sku_path=product_map_sku_path,
    map_old_sku="平台账号",
    map_new_sku="站点",
    map_sku_sheet='站点匹配'
)

# 保存目标列
main_df_4 = main_df_4[
    ['退件号', '订单号', '参考号', '订单参考号', '合并-映射账号', '映射站点', 'SKU', '实收数量', '良品',
     '退件费用(RMB)', '退件类型']
]

# 将处理后的数据保存到新的Excel文件
output_file_path = (main_file_path.rsplit('\\', 1)[0] + '\\(已完成-1)' +
                    main_file_path.rsplit('\\', 1)[1].replace('xls', 'xlsx'))
main_df_4.to_excel(output_file_path, index=False)  # index=False表示不保存索引列
print(f"处理完成，结果已保存到{output_file_path}")
print(f'{Color.YELLOW}~~~[注意]请检查，合并-映射账号、映射站点 是否都有了！！！--- ====== ---{Color.RESET} ')
print('合并-映射账号 很多没有的话，可能是RPA查询时，查询漏了，可以RPA再查一次！')
print('合并-映射账号  没有的话，询问：惠成')
