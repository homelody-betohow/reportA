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

from A_报表.Z_method.sku_映射 import sku_mappings
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.Z_method.style import Color
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-5)订单统计-{shared_date}.xlsx"
main_df = pd.read_excel(main_file_path)
# 删除所有以 "映射尾程-" 开头的列
main_df = main_df.drop(columns=[col for col in main_df.columns if col.startswith("映射尾程-")])

# HY、4PX、FBA仓库的，派送运费、fba 都等于0，映射transaction-FBA-派送运费 为空的，进行派送费映射
# 标记哪些行需要映射派送费（派送运费、fba 都等于0，映射transaction-FBA-派送运费 为空0）
mask = (main_df['派送运费'] == 0) & (main_df['fba费用'] == 0) & (main_df['映射transaction-FBA-派送运费'].isna())
# 需要映射的部分
main_df_map = main_df[mask].copy()
# 不需要映射的部分（至少有一个费用不等于0）
main_df_not_map = main_df[~mask].copy()
product_map_sku_dir = r"\\Betohow\数据报表\2-定价表"
product_map_sku_file = "欧洲平台定价表 2026.0708.xlsx"
product_map_sku_path = f"{product_map_sku_dir}\{product_map_sku_file}"
# 获取表头，转换为list，并处理NaN值
header_list = pd.read_excel(product_map_sku_path, sheet_name='基础表').iloc[1].fillna('').tolist()  # 将NaN转为空字符串
print(f'定价表的表头:{header_list}')
# 获取“派送费-映射分类”列的唯一值
unique_site_classes = main_df_map['派送费-映射分类'].unique()
# 初始化最终合并的数据框
main_df_map_1 = pd.DataFrame()
# 循环筛选并处理每个分类
for site_class in unique_site_classes:
    # 筛选出当前分类的数据
    site_df = main_df_map[main_df_map['派送费-映射分类'] == site_class]
    #  当前'派送费-映射分类' 在 定价表的表头，and 'MF' 不在 '派送费-映射分类'，则映射 定价表
    if site_class in header_list and 'MF' not in site_class:
        # 分站点映射 amazon 的尾程
        site_df_1 = sku_mappings(
            main_df=site_df,
            main_sku='SKU',
            map_sku_path=product_map_sku_path,  # 读取文件时，会跳过前2行，第3行当列名
            map_old_sku="百途鸿SKU",
            map_new_sku=site_class,
            map_sku_sheet='基础表'
        )
    else:
        site_df_1 = site_df
    # 如果是第一次循环，直接赋值
    if main_df_map_1.empty:
        main_df_map_1 = site_df_1
    else:
        # 否则合并数据框
        main_df_map_1 = pd.concat([main_df_map_1, site_df_1], ignore_index=True)
# 保存结果
output_path = main_file_path.replace('已完成-5', '已完成-51(定价表-映射尾程)')
main_df_map_1.to_excel(output_path, index=False)
# print(f'处理完成，output_path：{output_path}')


# 分仓库 合并映射的数据，并删掉 细分国家的数据
def merge_prefix_columns(df, prefix, new_col_name):
    """
    将以 prefix 开头的列按行合并（取第一个非空值），生成新列，并删除原列
    """
    # 找出需要合并的列
    cols = [col for col in df.columns if col.startswith(prefix)]
    if cols:
        # 合并：每行取第一个非空值
        df[new_col_name] = df[cols].apply(lambda row: next((v for v in row if pd.notna(v)), None), axis=1)
        # 删除原列
        df = df.drop(columns=cols)
    else:
        # 如果没有匹配的列，创建空列
        df[new_col_name] = None
    return df


# 批量处理
prefix_mapping = [
    ("映射FBA-", "单个-FBA-派送费"),
    ("映射HY-", "单个-HY-派送费"),
    ("映射4PX-", "单个-4PX-派送费"),
]
for prefix, new_name in prefix_mapping:
    main_df_map_1 = merge_prefix_columns(main_df_map_1, prefix, new_name)

# 合并回来
main_df_1 = pd.concat([main_df_map_1, main_df_not_map], ignore_index=True)

cols = ['单个-FBA-派送费', '单个-HY-派送费', '单个-4PX-派送费']
existing_cols = [c for c in cols if c in main_df_1.columns]
# 转换为数值类型
for col in existing_cols:
    main_df_1[col] = pd.to_numeric(main_df_1[col], errors='coerce')
# 计算每行非空值数量
non_null_count = main_df_1[existing_cols].notna().sum(axis=1)
# 只有刚好1个非空值时才取sum（因为只有一个值，sum就是它本身）
main_df_1['映射-单个-定价派送费'] = main_df_1[existing_cols].sum(axis=1, numeric_only=True)
main_df_1.loc[non_null_count != 1, '映射-单个-定价派送费'] = None
# 映射-定价派送费
main_df_1['映射-定价派送费'] = main_df_1['映射-单个-定价派送费'] * main_df_1['仓库SKU销量']  # 包含了sale、resend

# 重命名
main_df_1 = main_df_1.rename(columns={'派送运费': '原-派送运费'})

# 需要相加的列（用于非MF开头的情况）
fee_cols = [
    'fba费用',
    '原-派送运费',
    '映射transaction-FBA-派送运费',
    '映射-定价派送费'
]
# 统一把费用列清洗为数值（避免出现 "float + str"）
def _to_num(s: pd.Series) -> pd.Series:
    if s is None:
        return s
    # 兼容：空字符串、带逗号的数字、混合类型
    return pd.to_numeric(
        s.astype(str).str.replace(",", "", regex=False).replace({"": np.nan, "nan": np.nan, "None": np.nan}),
        errors="coerce",
    )

for _c in [*fee_cols, "MF-派送费"]:
    if _c in main_df_1.columns:
        main_df_1[_c] = _to_num(main_df_1[_c])
# 条件判断：派送费-映射分类 是否以 MF 开头
mask = main_df_1['派送费-映射分类'].astype(str).str.startswith('MF', na=False)
# 根据条件计算派送运费
main_df_1['派送运费'] = np.where(
    mask,
    main_df_1['MF-派送费'],  # MF开头：直接用 MF-派送费
    main_df_1[fee_cols].sum(axis=1, numeric_only=True)  # 其他情况：四列相加
)
# 把等于0的值替换为 NaN
main_df_1['派送运费'] = main_df_1['派送运费'].replace(0, np.nan)
# 将 '派送费-映射分类' 列中以 'ZHG' 开头的行对应的 '派送运费' 列设为 0
main_df_1.loc[main_df_1['派送费-映射分类'].str.startswith('ZHG', na=False), '派送运费'] = 0

# 保存结果
output_path = main_file_path.replace('已完成-5', '已完成-5-1')
main_df_1.to_excel(output_path, index=False)
print(f'处理完成，output_path：{output_path}')
print(f'{Color.YELLOW}~~~~~~~~~~~~~~~~~请检查，派送运费，是否有空的，分销的派送运费 为 0 ！！！{Color.RESET}')
print(r'~~~~~~~~~~~~~~~~~有空的话，去看看"欧洲平台定价表"，是否没及时更新，path：\\Betohow\数据报表\2-定价表')
print('================================================================================"')
print(f"{Color.GREEN} FBA-DE =VLOOKUP(R列,'[{product_map_sku_file}]基础表'!$B:$AF,30,FALSE){Color.RESET}")
print(f"{Color.GREEN} FBA-FR =VLOOKUP(R列,'[{product_map_sku_file}]基础表'!$B:$AF,31,FALSE){Color.RESET}")
print(f"{Color.GREEN} HY-AT，HY-PT =VLOOKUP(R列,'[手动-二次映射.xlsx]派送费-HY-PT、HY-AT'!$A:$C,3,FALSE){Color.RESET}")
print('================================================================================"')
print(f'{Color.YELLOW}~~~（注意"仓库SKU销量"的数量）~~~~~{Color.RESET}')
print('~~~~~~~~~~~~~~~~~"欧洲平台定价表"没有的话，联系：李杨，更新定价表的数据')
