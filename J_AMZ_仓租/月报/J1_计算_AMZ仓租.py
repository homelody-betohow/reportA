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
from A_报表.Z_method.platform_shop import map_shop_platform_region
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name, fba_date
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

# 忽略特定的 UserWarning
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
print(f"{fba_date}")
# TODO 文件路径！！！   上月的 利润报表
# main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\FBA仓租\FBA仓租明细{fba_date}.xlsx"
# FBA仓租 引用的是 SellerSku利润报表， 将 对应日期文件 重命名 即可
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\FBA仓租\FBA仓租明细{fba_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)
# 使用 sellerSku 列的数据填充 仓库sku 列的空值
main_file_df['仓库sku'] = main_file_df['仓库sku'].fillna(main_file_df['sellerSku'])


def extract_values(s):
    if pd.isna(s):  # 检查是否为 NaN
        return None  # 如果是 NaN，返回 None 或其他默认值
    if 'amzn.gr.' in s:
        return s.split(r'amzn.gr.')[-1].split('-')[0].split('_')[0]
    else:
        return s.split('#')[0].split('BCFBAFL')[0].split('FBFBAFL')[0]


# 应用提取规则，清洗 仓库sku
main_file_df['仓库sku'] = main_file_df['仓库sku'].apply(extract_values)
main_file_df = main_file_df.rename(columns={'仓库sku': 'SKU'})

# 映射 商品ID
product_map_sku_path = r"\\Betohow\数据报表\数据库\产品信息库2025.xlsx"
main_file_df_1 = sku_mappings(
    main_df=main_file_df,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="产品编码",
    map_new_sku="商品ID",
    map_sku_sheet='产品信息表'
)
main_file_df_1 = main_file_df_1.rename(columns={'映射商品ID': '商品ID'})

# 映射站点 / 映射平台（数据源：platform_shop）
main_file_df_3 = map_shop_platform_region(main_file_df_1, shop_col='店铺', site_col=None)

# 在 映射站点 后插入新列 站点商品ID识别码
new_column_name = "站点商品ID识别码"
new_column_data = main_file_df_3["映射站点"] + main_file_df_3["商品ID"]
target_column = "映射站点"
insert_position = main_file_df_3.columns.get_loc(target_column) + 1
main_file_df_3.insert(insert_position, new_column_name, new_column_data)

# 在 站点商品ID识别码后插入 平台商品ID识别码
new_column_name = "平台商品ID识别码"
new_column_data = main_file_df_3["映射平台"] + main_file_df_3["商品ID"]
target_column = "站点商品ID识别码"
insert_position = main_file_df_3.columns.get_loc(target_column) + 1
main_file_df_3.insert(insert_position, new_column_name, new_column_data)

# 得到上月的  FBA仓租费
main_file_df_3['FBA仓租费'] = main_file_df_3['仓储费用（已分摊）'] + main_file_df_3['长期仓储费（已分摊）']
main_file_df_3['FBA仓租费'] = main_file_df_3['FBA仓租费'].fillna(0).abs()  # 去掉 负号

# 删除"sellerSku"为空的行
main_file_df_4 = main_file_df_3.dropna(subset=['sellerSku'])
# 删除"FBA仓租费"为0的行
main_file_df_4 = main_file_df_4[main_file_df_4['FBA仓租费'] != 0]

main_file_df_4 = main_file_df_4[
    ['sellerSku', 'ASIN', '产品信息', 'SKU','商品ID', '店铺', '映射站点', '映射平台', '站点商品ID识别码',
     '平台商品ID识别码', '仓储费用（已分摊）', '长期仓储费（已分摊）', 'FBA仓租费']]
# 保存修改
output_path = main_file_path.rsplit('\\', 1)[0] + '\\(已完成-1)' + main_file_path.rsplit('\\', 1)[1]
main_file_df_4.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')

# 获取sellerSku为空的数据
empty_sellerSku_df = main_file_df_3[main_file_df_3['sellerSku'].isna()]
empty_sellerSku_FBA = empty_sellerSku_df['FBA仓租费'].sum()
print(f'\n---------------------sellerSku为空的FBA仓租费是：{empty_sellerSku_FBA}EUR------------------------------')
