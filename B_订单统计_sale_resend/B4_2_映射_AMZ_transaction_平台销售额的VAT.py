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
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name, transaction_date
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-3)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)
# 删掉 头程、关税 的计算列
main_file_df.drop(
    columns=['原-头程运费', '原-头程税费', '单个-头程运费', '单个-头程税费', '映射-头程运费', '映射-头程税费'],
    inplace=True
)

# 筛选 平台 == amazon
amazon_df = main_file_df[main_file_df['平台'] == 'amazon'].copy()
# 筛选 平台 != amazon
non_amazon_df = main_file_df[main_file_df['平台'] != 'amazon'].copy()

# 在 订单号 后插入映射列 订单号识别码
new_column_name = "订单号识别码"  # 映射列名
new_column_data = amazon_df["订单号"] + amazon_df["SKU"]  # 映射列数据
target_column = "订单号"  # 目标列名（在其后插入）
insert_position = amazon_df.columns.get_loc(target_column) + 1  # 计算插入位置
amazon_df.insert(insert_position, new_column_name, new_column_data)  # 插入映射列
# 按照 '订单号识别码' 列进行分组，进行汇总
amazon_df_1 = amazon_df.groupby('订单号识别码').agg({
    '平台': 'first',  # 保留每组的第一行数据
    '店铺英文名': 'first',
    '站点': 'first',
    '映射站点': 'first',
    '映射平台': 'first',
    '儿子-站点识别码': 'first',
    '儿子-平台识别码': 'first',
    '订单类型': 'first',
    '参考号': 'first',
    '订单号': 'first',
    'SKU': 'first',
    '仓库': 'first',
    '运输方式': 'first',
    '仓库SKU销量': 'first',
    '跟踪单号': 'first',
    '币种': 'first',
    '订单总金额': 'sum',
    '平台运费': 'sum',
    'fba费用': 'sum',
    '派送运费': 'sum',
    '头程运费': 'sum',
    '头程税费': 'sum',
}).reset_index()

# 映射 已发放-推迟订单 的AMZ派送费
transaction_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\transaction交易明细\(处理完成)transaction交易明细_已发放-推迟订单{transaction_date}.xlsx"
# 标记哪些行需要映射（两个费用都等于0）
mask = (amazon_df_1['派送运费'] == 0) & (amazon_df_1['fba费用'] == 0)
# 需要映射的部分
amazon_df_to_map = amazon_df_1[mask].copy()
amazon_df_mapped = sku_mappings(
    main_df=amazon_df_to_map,
    main_sku='订单号识别码',
    map_sku_path=transaction_path,
    map_old_sku="order-id识别码",
    map_new_sku="fba fees",
    map_sku_sheet='Sheet1'
)
amazon_df_mapped = amazon_df_mapped.rename(columns={'映射fba fees': '映射transaction-FBA-派送运费'})
# 不需要映射的部分（至少有一个费用不等于0）
amazon_df_not_map = amazon_df_1[~mask].copy()
# 合并回来
amazon_df_2 = pd.concat([amazon_df_mapped, amazon_df_not_map], ignore_index=True)

# 合并数据：非 Amazon 数据 + 处理后的 Amazon 数据
final_df = pd.concat([non_amazon_df, amazon_df_2], ignore_index=True)

# TODO 计算Amazon 的 平台销售额VAT
# 筛选“平台”列中包含“amazon”的行
df_filtered_amazon = final_df[final_df['平台'].str.contains('amazon', case=False, na=False)].copy()
product_map_sku_path = fr"{DESKTOP_ROOT}\VAT、平台费-映射.xlsx"
# 映射 VAT税
df_filtered_amazon_1 = sku_mappings(
    main_df=df_filtered_amazon,
    main_sku='映射站点',
    map_sku_path=product_map_sku_path,
    map_old_sku="站点",
    map_new_sku="VAT税",
    map_sku_sheet='VAT税、佣金'
)
# 计算VAT
df_filtered_amazon_1['平台销售额VAT-amazon'] = df_filtered_amazon_1['订单总金额'] * df_filtered_amazon_1['映射VAT税']
# 筛选："映射站点"不包含 'FB'的，对应‘平台销售额VAT-amazon’列的数据替换成0
condition = ~df_filtered_amazon_1['映射站点'].str.contains('FB')    # 只有FB的有VAT
df_filtered_amazon_1.loc[condition, '平台销售额VAT-amazon'] = 0
# 更新回数据
cols_to_update = ['映射VAT税', '平台销售额VAT-amazon']
final_df = final_df.join(df_filtered_amazon_1[cols_to_update])  # 按索引回填

# 得到：派送费-映射分类
# 仓库 为 '--'，目前只出现在Amazon平台，统一换成 'FBA仓库'
final_df['仓库'] = final_df['仓库'].str.replace('--', 'FBA仓库')


# 站点分类函数
def classify_site(row):
    platform = str(row['平台']).upper()
    site = row['站点']

    # AMAZON 的站点分类：['FR', 'ES', 'IT', 'BE', 'DE']为本身；其余为分类站点为：DE
    if 'AMAZON' in platform:
        return site if site in ['FR', 'ES', 'IT', 'BE', 'DE'] else 'DE'
    # TEMU 都是 HY仓库,站点分类：['DE', 'ES', 'FR', 'IT', 'PL', 'CZ', 'BE', 'NL']为本身；其余为分类站点为：IT
    elif 'TEMU' in platform:
        return site if site in ['DE', 'ES', 'FR', 'PL', 'CZ', 'BE', 'NL', 'IT'] else 'IT'
    else:
        return site


# 仓库分类函数
def classify_warehouse(warehouse):
    wh = str(warehouse).upper()
    if '4PX' in wh: return '4PX'
    if 'HY' in wh: return 'HY'
    if 'FBA' in wh: return 'FBA'
    if 'MF' in wh: return 'MF'
    if 'ZHG' in wh: return 'ZHG'
    raise ValueError(f'无法识别的仓库代码: {warehouse}，请检查“仓库”数据！')


# 生成'派送费-映射分类'列
final_df['派送费-映射分类'] = final_df.apply(
    lambda row: classify_warehouse(row['仓库']) + '-' + classify_site(row), axis=1
)

# 保存结果
output_path = main_file_path.replace('已完成-3', '已完成-4')
final_df.to_excel(output_path, index=False)
print(f'处理完成，output_path：{output_path}')
