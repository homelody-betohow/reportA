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
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT, MONTH_GOAL_EXCEL_PATH

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-21)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)

# 拆分数据
EU_df = main_file_df[main_file_df['平台'].isin(['AMAZON-EU'])]  # 平台 包含 AMAZON-EU
US_df = main_file_df[main_file_df['平台'].isin(['AMAZON-US'])]  # 平台 包含 AMAZON-US
no_EU_US_df = main_file_df[~main_file_df['平台'].isin(['AMAZON-EU', 'AMAZON-US'])]  # 平台 不 包含 AMAZON-EU、AMAZON-US

# 映射 销售负责人  AMAZON-EU
product_map_sku_path = MONTH_GOAL_EXCEL_PATH
EU_df_1 = sku_mappings(
    main_df=EU_df,
    main_sku='商品ID',
    map_sku_path=product_map_sku_path,
    map_old_sku="商品ID",
    map_new_sku="负责人",
    map_sku_sheet='AMAZON-EU'
)

# 再次映射  销售负责人，映射自己记录的  销售负责人-SKU（AMAZON-EU）
product_map_sku_path = fr'{DESKTOP_ROOT}\信息-映射.xlsx'
EU_df_2 = sku_mappings(
    main_df=EU_df_1,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="SKU",
    map_new_sku="销售负责人-SKU（AMAZON-EU）",
    map_sku_sheet='销售负责人'
)
# 将 映射销售负责人-SKU（AMAZON-EU） 不为空的行，将这些值赋给对应的 映射负责人 列
EU_df_2.loc[EU_df_2['映射销售负责人-SKU（AMAZON-EU）'].notna(), '映射负责人'] = EU_df_2['映射销售负责人-SKU（AMAZON-EU）']

# 删除 EU_df_2 的"销售负责人"列
EU_df_2 = EU_df_2.drop(columns=['销售负责人'])
# 重命名
EU_df_2 = EU_df_2.rename(columns={'映射负责人': '销售负责人'})
# EU_df_2 销售负责人 为空的，则填入 无负责人
EU_df_2['销售负责人'] = EU_df_2['销售负责人'].fillna('无负责人')

# 映射 销售负责人  AMAZON-US
product_map_sku_path = MONTH_GOAL_EXCEL_PATH
US_df_1 = sku_mappings(
    main_df=US_df,
    main_sku='商品ID',
    map_sku_path=product_map_sku_path,
    map_old_sku="商品ID",
    map_new_sku="负责人",
    map_sku_sheet='AMAZON-US'
)
# 删除 US_df_1 的"销售负责人"列
US_df_1 = US_df_1.drop(columns=['销售负责人'])
# 重命名
US_df_1 = US_df_1.rename(columns={'映射负责人': '销售负责人'})
# AMAZON-US 的 SKU，以“U” 开头的，销售负责人 是 官雪婷US
US_df_1.loc[US_df_1['SKU'].str.startswith('U', na=False), '销售负责人'] = '官雪婷US'
# US_df_1 销售负责人 为空的，则填入 无负责人
US_df_1['销售负责人'] = US_df_1['销售负责人'].fillna('无负责人')

# 合并数据
main_file_df_1 = pd.concat([EU_df_2, US_df_1, no_EU_US_df]).reset_index(drop=True)

# 保存结果到新的 Excel 文件
output_path = main_file_path.replace('已完成-21', '已完成-22')
main_file_df_1.to_excel(output_path, index=False)
print(f"结果已保存到 {output_path}")
