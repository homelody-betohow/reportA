import re
import importlib.util
from pathlib import Path

# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

import pandas as pd
from A_报表.Z_method.style import Color
from A_报表.Z_method.sku_映射 import sku_mappings
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-1)订单统计-{shared_date}.xlsx"
main_df = pd.read_excel(main_file_path)

# 一、映射 REAL-FB 没有站点（国家）的
real_fb_order_list_path = r"\\Betohow\数据报表\RPA\报表-无站点-订单查询\REAL-FB\REAL-FB无站点-订单管理-近6个月.csv"
real_fb_order_df = pd.read_csv(real_fb_order_list_path)
# 清理CSV中的 =" " 格式，并转换为xlsx（sku_映射需要.xlsx格式）
for col in ['销售参考号', '国家或地区代码']:
    real_fb_order_df[col] = real_fb_order_df[col].str.replace(r'^="(.*)"$', r'\1', regex=True)
product_map_sku_path = real_fb_order_list_path.replace('.csv', '.xlsx')
real_fb_order_df.to_excel(product_map_sku_path, index=False)  # 保存 为 xlsx文件
# 映射订单管理-查询结果的“国家或地区代码”
main_df = sku_mappings(
    main_df=main_df,
    main_sku='订单号',
    map_sku_path=product_map_sku_path,
    map_old_sku="销售参考号",
    map_new_sku="国家或地区代码",
    map_sku_sheet="Sheet1"
)
# 如果“映射国家或地区代码”非空，则替换“站点”
main_df['站点'] = main_df.apply(
    lambda row: row['映射国家或地区代码']
    if pd.notna(row['映射国家或地区代码']) and str(row['映射国家或地区代码']).strip() != ''
    else row['站点'], axis=1
)

# 二、映射 LM-BC的重发订单对应的原订单的平台SKU ——区分负责人
# LM_BC、LM_RP 的  重发订单  的 平台SKU 映射
shops = {
    'LM_BC_FR': r'\\Betohow\数据报表\RPA\报表-无站点-订单查询\LM-BC-重发\LM-BC-重发-订单管理-近6个月.csv',
    'LM_RP_FR': r'\\Betohow\数据报表\RPA\报表-无站点-订单查询\LM-RP-重发\LM-RP-重发-订单管理-近6个月.csv',
}
have_order_shop_list = []
for shop, path in shops.items():
    mask = (main_df['店铺英文名'] == shop) & (main_df['订单类型'] == '重发订单')
    if not mask.any():
        print(f"\n{shop}，没有重发订单！\n")
        continue
    have_order_shop_list.append(shop)
    df_order = pd.read_csv(path)
    for c in ['销售参考号', 'SKU']: df_order[c] = df_order[c].str.replace(r'^="(.*)"$', r'\1', regex=True)
    sku_map = df_order.set_index('销售参考号')['SKU']
    # 创建去掉尾缀的 临时订单号 用于匹配，去掉 订单号 尾缀的 -数字（例如： -1、-2、-3、-4、-5）
    order_no_clean = main_df.loc[mask, '订单号'].str.replace(r'-\d$', '', regex=True)
    # 用清理后的订单号判断是否匹配
    matched = mask & order_no_clean.isin(sku_map.index)
    # 保留第一个匹配的 SKU
    sku_map = sku_map[~sku_map.index.duplicated(keep='first')]
    # 映射SKU时使用清理后的订单号，但原订单号不变
    main_df.loc[matched, '平台sku'] = order_no_clean[matched].map(sku_map)
    main_df.loc[matched, '订单号'] += '——已映射"' + main_df.loc[matched, '平台sku'] + '"'

# 保存结果
output_path = main_file_path.replace('已完成-1', '已完成-1-1')
main_df.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
# print(f'{Color.YELLOW}~~~~~~~~~~~~~~~~~请检查，"店铺英文名" == FB_REAL，"站点"是否都有 站点（国家）！！！{Color.RESET}')
print('-' * 100)
for have_order_shop in have_order_shop_list:
    print(f'{Color.YELLOW}~~~~~~~~~~~~~~~~~请检查，"店铺英文名" == {have_order_shop}，"重发订单"是否都已映射 "平台SKU"！！！{Color.RESET}')
