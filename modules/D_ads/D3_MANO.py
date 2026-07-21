"""
D3_MANO.py — ManoMano 平台广告数据处理脚本

功能概述：
  1. 从桌面指定日期文件夹下，批量读取各站点子目录中的 MANO 广告 CSV
  2. 按文件夹名识别站点，合并为一张表
  3. 映射 EAS → 仓库 SKU，拆分组合 SKU（含 + 号），再映射站点 → 平台
  4. 生成「儿子-站点/平台识别码」，输出最终广告费报表

输出文件（均在 ...\\广告\\MANO 目录下）：
  - (已完成-1)MANO广告.xlsx  — 原始合并结果（含广告消耗为 0 的行）
  - (处理完成)MANO广告.xlsx    — 清洗、映射后的最终结果
"""

import importlib.util
import os
import glob
from pathlib import Path

import pandas as pd

# 须在 import config/common 之前：向上查找项目根目录并加入 sys.path，保证包导入可用
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.sku_mapping import sku_mappings
from common.platform_shop import map_region_to_platform
from config.A0_set_date import shared_date, folder_name
from common.split_rows_data_SKU import split_one_rows_data
from config.A0_paths import DESKTOP_ROOT

# CSV 所在「一级子文件夹名」→ 内部统一站点编码
# 文件夹名来自 ManoMano 后台导出时的店铺/账号目录，与 CSV 文件名无关
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

    # 邮箱前缀形式的文件夹名（新版导出目录结构）
    "onemanofr@outlook.com_FR-B2C OHPA- FR B2C": "MANO-FR-OHPA",
    "onemanofr@outlook.com_FR-MMF OHPA MF": "MANO-FR-OHPAMF",
    "onemanofr@outlook.com_DE-MMF DE - MMF OHPA": "MANO-DE-OHPAMF",
    "onemanofr@outlook.com_DE-MMF DE - MFX_DE_OHPA-B2B": "MANO-DE-OHPAMF-B2B",
    "onemanofr@outlook.com_FR-MMF FR - MF - OHPA-FR B2B": "MANO-FR-OHPAMF-B2B",

    "RPCOMMERCE@yeah.net_FR-B2C RP COMMERCE SARL": "MANO-FR-COM",
    "RPCOMMERCE@yeah.net_DE-B2C DE B2C": "MANO-DE-COM",
    "RPCOMMERCE@yeah.net_FR-MMF FR - MF": "MANO-FR-COMMF",
    "RPCOMMERCE@yeah.net_DE-MMF DE - MF": "MANO-DE-COMMF",
    "RPCOMMERCE@yeah.net_FR-MMF FR - MF B2B": "MANO-FR-COMMF-B2B",
}

# ManoMano 后台导出列名会随版本变化；统一映射为脚本内部使用的列名
COLUMN_RENAME = {
    'Ad spend': '广告消耗',
}


def coalesce_duplicate_columns(df):
    """合并重复列名：同名的多列按“从左到右取第一个非空值”折叠为一列。

    MANO 新旧版导出可能同时含 'Ad spend' 与 '广告消耗'（或带空格的重名列），
    统一重命名后会产生同名列，导致后续 df['广告消耗'] 返回 DataFrame 而报错。
    """
    if not df.columns.duplicated().any():
        return df
    merged = {}
    for col in pd.unique(df.columns):
        same = df.loc[:, df.columns == col]
        # bfill(axis=1) 让每行从左到右取第一个非空值，再取第一列即为合并结果
        merged[col] = same.iloc[:, 0] if same.shape[1] == 1 else same.bfill(axis=1).iloc[:, 0]
    return pd.DataFrame(merged)


def normalize_mano_columns(df):
    """将新旧版 CSV 列名统一为脚本内部列名，并去除列名首尾空格。"""
    df = df.rename(columns=lambda c: c.strip() if isinstance(c, str) else c)
    df = df.rename(columns=COLUMN_RENAME)
    df = coalesce_duplicate_columns(df)  # 折叠重名列，避免下游布尔索引报错
    if '广告消耗' not in df.columns:
        raise RuntimeError(
            f"未找到广告消耗列，当前列名: {df.columns.tolist()}；"
            f"请在 COLUMN_RENAME 中补充映射"
        )
    return df


def read_csv_auto_encode(file_path):
    """智能读取 CSV，依次尝试 gb18030 / utf-8 / latin1 编码。"""
    for enc in ['gb18030', 'utf-8', 'latin1']:
        try:
            return pd.read_csv(file_path, encoding=enc)
        except UnicodeDecodeError:
            continue
    raise RuntimeError(f"文件编码无法识别: {file_path}")


def parse_site_from_folder(file_path, root_path):
    """
    根据 CSV 相对根目录的一级子文件夹名，解析出统一站点编码。

    目录结构示例：MANO/Betohow-DE/xxx_202501.csv → MANO-DE-BTH
    若文件夹名已是站点编码（在 SITE_MAPPING 的 value 中），则直接返回。
    """
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
    在 root_path 下递归查找文件名含 shared_date 的 CSV，合并为一张 DataFrame。

    每个文件会新增「站点」列；全空列会被丢弃，避免 concat 时列错位。
    """
    search_pattern = os.path.join(root_path, '**', f'*{shared_date}.csv')
    csv_files = glob.glob(search_pattern, recursive=True)

    if not csv_files:
        raise RuntimeError(f"未找到包含 {shared_date} 的CSV文件，请检查路径")

    df_list = []
    for file_path in csv_files:
        df = read_csv_auto_encode(file_path)
        df['站点'] = parse_site_from_folder(file_path, root_path)
        df = df.dropna(axis=1, how='all')  # 剔除全空列，防止不同文件列数不一致
        df_list.append(df)

    merged_df = pd.concat(df_list, ignore_index=True, sort=False)
    return merged_df


# ── 主流程 ──────────────────────────────────────────────────────────

# 输入目录：桌面 \\ {folder_name}{shared_date} \\ 广告 \\ MANO
MANO_dir = fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\广告\MANO'

# 步骤 1：合并各站点原始 CSV
mano_df = normalize_mano_columns(merge_mano_ad_data(MANO_dir))
output_file_path = MANO_dir + '\\(已完成-1)MANO广告.xlsx'
mano_df.to_excel(output_file_path, index=False, engine='openpyxl')
print(f"合并完成，输出文件: {output_file_path}")

# 步骤 2：过滤无广告花费的行（广告消耗 = 0 表示该 SKU 当期无投放）
mano_df = mano_df[mano_df['广告消耗'] != 0]

product_map_sku_path = fr"{DESKTOP_ROOT}\广告-SKU关系对应.xlsx"

# 步骤 3：平台 SKU（EAS）→ 仓库 SKU（左连接映射表「MANO EAS对应表」）
mano_df_1 = sku_mappings(
    main_df=mano_df,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="EAS",
    map_new_sku="仓库SKU",
    map_sku_sheet='MANO EAS对应表'
)
mano_df_1['映射仓库SKU'] = mano_df_1['映射仓库SKU'].str.strip()

# 步骤 4：拆分组合 SKU（如 A+B+C），广告消耗按子 SKU 数量均摊到多行
mano_df_2 = split_one_rows_data(
    input_df=mano_df_1,
    data_column='映射仓库SKU',
    value_column='广告消耗'
)

# 步骤 5：站点编码 → 平台名称（数据源：platform_shop）
mano_df_3 = map_region_to_platform(mano_df_2, site_col='站点')

# 步骤 6：列重命名与 SKU 规范化
# 原平台 SKU 保留为「原SKU」，映射后的仓库 SKU 作为最终「SKU」
mano_df_3 = mano_df_3.rename(columns={'SKU': '原SKU'})
mano_df_3 = mano_df_3.rename(columns={'映射仓库SKU': 'SKU'})
# AE/OHE 前缀统一替换为 E，与仓库 SKU 命名规则对齐
mano_df_3['SKU'] = mano_df_3['SKU'].str.strip().str.replace(r'AE|OHE', 'E', regex=True)

# 步骤 7：生成识别码，供后续 D5/D6 与其他平台广告、订单数据关联
mano_df_3['SKU-站点识别码'] = mano_df_3['站点'] + mano_df_3['SKU']
mano_df_3['SKU-平台识别码'] = mano_df_3['映射平台'] + mano_df_3['SKU']

# 步骤 8：只保留下游需要的列，广告消耗改名为「广告费(非AMZ)」
mano_df_3 = mano_df_3[['原SKU', 'SKU', '站点', '映射平台', 'SKU-站点识别码', 'SKU-平台识别码', '广告消耗']]
mano_df_3 = mano_df_3.rename(columns={'广告消耗': '广告费(非AMZ)'})

output_file_path = MANO_dir + '\\(处理完成)MANO广告.xlsx'
mano_df_3.to_excel(output_file_path, index=False)
print(f"处理完成，结果已保存到{output_file_path}")
