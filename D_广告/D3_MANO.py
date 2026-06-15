import importlib.util
import os
import glob
from pathlib import Path

import pandas as pd

# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.Z_method.sku_映射 import sku_mappings
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.Z_method.split_rows_data_拆分SKU_1个加号_逗号 import split_one_rows_data
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

# 一级子文件夹名与站点的映射关系（CSV 放在对应子文件夹内即可，与文件名无关）
SITE_MAPPING = {
    "Betohow-DE": "MANO-DE-BTH",
    "Betohow-FR": "MANO-FR-BTH",
    "Betohow-IT": "MANO-IT-BTH",

    "B2C RP COMMERCE SARL-FR": "MANO-FR-COM",
    "B2C OHPA- FR B2C": "MANO-FR-OHPA",
    "B2C DE - B2C OHPA": "MANO-DE-OHPA",
    "OHPA- FR": "MANO-FR-OHPA",
    "OHPA MF-FR": "MANO-FR-OHPAMF",
    "DE - MMF OHPA": "MANO-DE-OHPAMF",
    "MMF Betohow MF-FR": "MANO-FR-BTHMF",
    "MMF Ubeegol MF-FR": "MANO-FR-UBGMF",
    "MMF DE - MMF OHPA": "MANO-DE-OHPAMF",
    "MMF FR - MF - OHPA-FR B2B": "MANO-FR-OHPAMF-B2B",
    "MMF OHPA MF-FR": "MANO-FR-OHPAMF",
    "COM-B2C DE": "MANO-DE-COM",
    "COM-B2C IT": "MANO-IT-COM",
    "COM-B2C FR": "MANO-FR-COM",
    "COM-MMF DE": "MANO-DE-COMMF",
    "COM-MMF FR": "MANO-FR-COMMF",
"onemanofr@outlook.com_FR-B2C OHPA- FR B2C":"MANO-FR-OHPA",
"onemanofr@outlook.com_FR-MMF OHPA MF":"MANO-FR-OHPAMF",
"onemanofr@outlook.com_DE-MMF DE - MMF OHPA":"MANO-DE-OHPAMF",
"onemanofr@outlook.com_DE-MMF DE - MFX_DE_OHPA-B2B":"MANO-DE-OHPAMF-B2B",
"onemanofr@outlook.com_FR-MMF FR - MF - OHPA-FR B2B":"MANO-FR-OHPAMF-B2B",

"RPCOMMERCE@yeah.net_FR-B2C RP COMMERCE SARL":"MANO-FR-COM",
"RPCOMMERCE@yeah.net_DE-B2C DE B2C":"MANO-DE-COM",
"RPCOMMERCE@yeah.net_FR-MMF FR - MF":"MANO-FR-COMMF",
"RPCOMMERCE@yeah.net_DE-MMF DE - MF":"MANO-DE-COMMF",
"RPCOMMERCE@yeah.net_FR-MMF FR - MF B2B":"MANO-FR-COMMF-B2B",

}


def read_csv_auto_encode(file_path):
    """智能读取CSV，自动尝试gb18030/utf-8/latin1编码"""
    for enc in ['gb18030', 'utf-8', 'latin1']:
        try:
            return pd.read_csv(file_path, encoding=enc)
        except UnicodeDecodeError:
            continue
    raise RuntimeError(f"文件编码无法识别: {file_path}")


def parse_site_from_folder(file_path, root_path):
    """从 CSV 所在的一级子文件夹名解析站点标识"""
    rel = os.path.relpath(file_path, root_path)
    parts = rel.split(os.sep)
    if len(parts) < 2:
        raise RuntimeError(
            f"CSV 须放在 MANO 下的一级子文件夹中（每文件夹对应一个站点），当前: {rel}"
        )
    folder_name = parts[0]
    if folder_name in SITE_MAPPING:
        return SITE_MAPPING[folder_name]
    if folder_name in SITE_MAPPING.values():
        return folder_name
    raise RuntimeError(f"文件夹无法匹配站点: {folder_name}（路径: {rel}）")


def merge_mano_ad_data(root_path):
    """
    批量合并MANO广告CSV文件
    参数: root_path - 搜索根目录
    返回: 生成的Excel文件路径
    """
    # 递归查找所有包含日期的CSV文件
    search_pattern = os.path.join(root_path, '**', f'*{shared_date}.csv')
    csv_files = glob.glob(search_pattern, recursive=True)

    if not csv_files:
        raise RuntimeError(f"未找到包含 {shared_date} 的CSV文件，请检查路径")
    # 逐个读取并标记站点
    df_list = []
    for file_path in csv_files:
        df = read_csv_auto_encode(file_path)
        df['站点'] = parse_site_from_folder(file_path, root_path)
        # 剔除空列防止合并异常
        df = df.dropna(axis=1, how='all')
        df_list.append(df)
    # 纵向合并所有数据
    merged_df = pd.concat(df_list, ignore_index=True, sort=False)
    return merged_df


MANO_dir = fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\广告\MANO'
mano_df = merge_mano_ad_data(MANO_dir)
# 将处理后的数据保存到新的Excel文件
output_file_path = MANO_dir + '\\(已完成-1)MANO广告.xlsx'
mano_df.to_excel(output_file_path, index=False, engine='openpyxl')
print(f"合并完成，输出文件: {output_file_path}")

# 删除"广告消耗"列为0的行
mano_df = mano_df[mano_df['广告消耗'] != 0]  # 广告消耗：广告花费

product_map_sku_path = fr"{DESKTOP_ROOT}\广告-SKU关系对应.xlsx"
#  映射sku（儿子）
mano_df_1 = sku_mappings(
    main_df=mano_df,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="EAS",
    map_new_sku="仓库SKU",
    map_sku_sheet='MANO EAS对应表'
)
# 去除 '映射仓库SKU' 列的前后空格
mano_df_1['映射仓库SKU'] = mano_df_1['映射仓库SKU'].str.strip()

#  拆分有“+”的  映射仓库sku
mano_df_2 = split_one_rows_data(
    input_df=mano_df_1,
    data_column='映射仓库SKU',
    value_column='广告消耗'
)
# 映射 平台
product_map_sku_path = fr"{DESKTOP_ROOT}\站点-匹配表.xlsx"  # 改成对应的映射表
mano_df_3 = sku_mappings(
    main_df=mano_df_2,
    main_sku='站点',
    map_sku_path=product_map_sku_path,
    map_old_sku="站点",
    map_new_sku='平台',
    map_sku_sheet='站点匹配'
)
# 重命名
mano_df_3 = mano_df_3.rename(columns={'SKU': '原SKU'})
mano_df_3 = mano_df_3.rename(columns={'映射仓库SKU': 'SKU'})

mano_df_3['SKU'] = mano_df_3['SKU'].str.strip().str.replace(r'AE|OHE', 'E', regex=True)

# 构建识别码
mano_df_3['儿子-站点识别码'] = mano_df_3['站点'] + mano_df_3['SKU']
mano_df_3['儿子-平台识别码'] = mano_df_3['映射平台'] + mano_df_3['SKU']
# 保存目标列
mano_df_3 = mano_df_3[['原SKU', 'SKU', '站点', '映射平台', '儿子-站点识别码', '儿子-平台识别码', '广告消耗']]
# 更改列名，将’广告消耗‘  改为 ’广告费(非AMZ)‘
mano_df_3 = mano_df_3.rename(columns={'广告消耗': '广告费(非AMZ)'})
# 将处理后的数据保存到新的Excel文件
output_file_path = MANO_dir + '\\(处理完成)MANO广告.xlsx'
mano_df_3.to_excel(output_file_path, index=False)  # index=False表示不保存索引列
print(f"处理完成，结果已保存到{output_file_path}")
