import pandas as pd
import importlib.util
from pathlib import Path
# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.Z_method.style import Color
from A_报表.Z_method.sku_映射 import sku_mappings
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-4)订单统计-{shared_date}.xlsx"
main_df = pd.read_excel(main_file_path)

# 包含MF站点、MF仓库的，都要映射派送费（订单统计的偏低）
product_map_sku_path = fr"{DESKTOP_ROOT}\MANO-MF 尾程.xlsx"  # 改成对应的映射表
# 筛选条件  包含MF站点、MF仓库的
mask_mf_site = main_df['映射站点'].str.contains('MF', na=False)
mask_mf_category = main_df['派送费-映射分类'].str.startswith('MF', na=False)
# 同时满足两个条件的行
mf_df = main_df[mask_mf_site & mask_mf_category].copy()
# 剩下的行（不满足上述条件）
non_mf_df = main_df[~(mask_mf_site & mask_mf_category)].copy()
# 获取当前筛选，“映射站点”列的唯一值
unique_site_list = mf_df['映射站点'].unique()
# 初始化最终合并的数据框
mf_df_1 = pd.DataFrame()
# 循环筛选并处理每个分类
for site in unique_site_list:
    # 筛选出当前筛选站点后的数据
    site_df = mf_df[mf_df['映射站点'] == site]
    # 分站点映射 MANO平台 的尾程
    site_df_1 = sku_mappings(
        main_df=site_df,
        main_sku='SKU',
        map_sku_path=product_map_sku_path,
        map_old_sku=site,
        map_new_sku=f"尾程-{site}",
        map_sku_sheet='Sheet1'
    )
    # 如果是第一次循环，直接赋值
    if mf_df_1.empty:
        mf_df_1 = site_df_1
    else:
        # 否则合并数据框
        mf_df_1 = pd.concat([mf_df_1, site_df_1], ignore_index=True)
# 合并所有派送费-映射分类的新列数据到一个新列：单个-MF-派送费
mf_df_1['单个-MF-派送费'] = (mf_df_1.filter(regex='映射尾程-').apply(lambda s: s.dropna().iloc[0], axis=1))
# 计算 MF站点的 派送运费 = 映射FBA费  *  仓库SKU销量    '单个-MF-派送费'没有映射到的话，对应位置为  空
mf_df_1['单个-MF-派送费'] = pd.to_numeric(mf_df_1['单个-MF-派送费'], errors='coerce')
# MF-派送费
mf_df_1['MF-派送费'] = mf_df_1['单个-MF-派送费'] * mf_df_1['仓库SKU销量']  # 包含了sale、resend

# 合并数据
main_df_1 = pd.concat([mf_df_1, non_mf_df], ignore_index=True)
# 保存结果
output_path = main_file_path.replace('已完成-4', '已完成-5')
main_df_1.to_excel(output_path, index=False)
print(f'处理完成，output_path：{output_path}')
print(f'{Color.YELLOW}~~~~~~~~~~~~~~~~~请检查，“派送费-映射分类”含’MF‘的 “MF-派送费”，是否有空的（注意"仓库SKU销量"的数量）！！！{Color.RESET}')
print(f"单个-MF-派送费 {Color.RED} =VLOOKUP(G列,'[手动-二次映射.xlsx]MF-派送费'!$A:$B,2,FALSE){Color.RESET}")
print(
    f'MF仓库有空的话，找：王园芳，补充基础表：{DESKTOP_ROOT}\\MANO-MF 尾程.xlsx，COMMF和OHPAMF是一样的，同时补充！（没有出单的，先映射定价表（MF-站点））')
