import numpy as np
import pandas as pd
import importlib.util
import openpyxl
from openpyxl.styles import Alignment
from pathlib import Path

# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_PROJECT_ROOT = _epr_mod.bootstrap(__file__)

from common.style import Color
from common.platform_shop import map_site_vat_commission
from common.castorama_commission import (
    apply_castorama_commission_from_json,
    castorama_commission_path,
    merge_missing_into_castorama_commission_json,
)
from config.A0_set_date import shared_date, folder_name
from config.A0_paths import DESKTOP_ROOT

# 本机映射（取代桌面「castorama - SKU类目佣金比例.xlsx」）
CASTORAMA_COMMISSION_PATH = castorama_commission_path(_PROJECT_ROOT)


# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-7)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)
# 保留「分销」列（C4 fillna(0) 可能将其变为 0）
if '分销' not in main_file_df.columns:
    main_file_df['分销'] = '否'
else:
    main_file_df['分销'] = main_file_df['分销'].replace({0: '否', '0': '否'})
    main_file_df['分销'] = main_file_df['分销'].fillna('否')
# 重命名
main_file_df = main_file_df.rename(columns={'映射VAT税': 'amazon-VAT税'})

# 映射 castorama 的 佣金比例（本机 JSON，取代桌面 xlsx）
main_file_df = apply_castorama_commission_from_json(
    main_file_df, CASTORAMA_COMMISSION_PATH, log_tag="C5"
)

# 映射 平台费（佣金）、VAT税（依据数据表platform_shop设置：佣金费率、VAT费率）
main_file_df_1 = map_site_vat_commission(
    main_df=main_file_df, site_col='站点', excel_fallback=False
)

# castorama：平台费不用 DB，一律用 SKU 类目「映射佣金比」
_castorama = main_file_df_1['平台'].astype(str).str.lower() == 'castorama'
main_file_df_1.loc[_castorama, '映射平台费（佣金）'] = main_file_df_1.loc[_castorama, '映射佣金比']

# 其他平台：用“映射佣金比”填补“映射平台费（佣金）”的空值
main_file_df_1['映射平台费（佣金）'] = main_file_df_1['映射平台费（佣金）'].fillna(main_file_df_1['映射佣金比'])

# castorama 仍为空：按 SKU 第4-5位兜底（01/02→0.1，03→0.12；非数字如 BF 不转换）
_sku_rule = main_file_df_1['SKU'].astype(str).str.slice(3, 5).map(
    {'01': 0.1, '02': 0.1, '03': 0.12}
)
_need_rule = _castorama & main_file_df_1['映射平台费（佣金）'].isna()
main_file_df_1.loc[_need_rule, '映射平台费（佣金）'] = _sku_rule.loc[_need_rule]
_rule_hit = int((_need_rule & main_file_df_1['映射平台费（佣金）'].notna()).sum())
if _rule_hit:
    print(f"{Color.CYAN}[C5] SKU 第4-5位规则兜底：补全 {_rule_hit} 行{Color.RESET}")

# AMAZON-EU / AMAZON-US：SKU 以 E39、E61 开头 → 映射平台费（佣金）= 15%
_amz_e39_e61 = (
    main_file_df_1['平台'].isin(['AMAZON-EU', 'AMAZON-US'])
    & main_file_df_1['SKU'].astype(str).str.startswith(('E39', 'E61'))
)
_amz_e39_e61_n = int(_amz_e39_e61.sum())
if _amz_e39_e61_n:
    main_file_df_1.loc[_amz_e39_e61, '映射平台费（佣金）'] = 0.15
    print(f"{Color.CYAN}[C5] AMAZON E39/E61 佣金 15%：覆盖 {_amz_e39_e61_n} 行{Color.RESET}")

main_file_df_2 = main_file_df_1

"""
如果 平台 在["AMAZON-EU", "AMAZON-US", "DLZ-EU"]，则  平台费、VAT = 平台销售额 * 对应比例
(AMAZON的VAT、平台费先这样计算，后面没有VAT的会替换0；平台费会去 - 退款的佣金)
否则，平台费、VAT = 销售额 * 对应比例
"""
cond = (
    (main_file_df_2['平台'].isin(["AMAZON-EU", "AMAZON-US", "DLZ-EU"]))
)

# 1 计算平台费
main_file_df_2['平台费'] = np.where(
    cond,
    main_file_df_2['平台销售额'] * main_file_df_2['映射平台费（佣金）'],
    main_file_df_2['销售额'] * main_file_df_2['映射平台费（佣金）']
)
# 2 VAT
main_file_df_2['销售税-本土'] = np.where(
    cond,
    main_file_df_2['平台销售额'] * main_file_df_2['映射VAT税'],
    main_file_df_2['销售额'] * main_file_df_2['映射VAT税']
)

# 将“平台”中包含 "AMAZON" 的"销售税-本土"替换 0
main_file_df_2.loc[main_file_df_2["平台"].str.contains("AMAZON", na=False), "销售税-本土"] = 0
# 亚马逊的VAT = 平台销售额VAT-amazon - 销售退款金额VAT-amazon （本土平台对应位置为0）
main_file_df_2['销售税'] = main_file_df_2['销售税-本土'] + main_file_df_2['平台销售额VAT-amazon'] - main_file_df_2[
    '销售退款金额VAT-amazon']
# 亚马逊的平台费 = 平台销售额的平台费 - 销售退款金额的佣金
main_file_df_2['平台费'] = main_file_df_2['平台费'] - main_file_df_2['销售退款金额的佣金']
# 确保相关列的数据类型为数值类型
main_file_df_2['平台费'] = np.round(pd.to_numeric(main_file_df_2['平台费'], errors='coerce'), 2)
main_file_df_2['销售税'] = np.round(pd.to_numeric(main_file_df_2['销售税'], errors='coerce'), 2)

# 创建新列“平台费(AMZ)”和“销售税(AMZ)”，初始值为NaN
main_file_df_2['平台费(AMZ)'] = np.nan
main_file_df_2['销售税(AMZ)'] = np.nan

# 如果“平台”列等于“amazon”，则将“平台费”的值移动到“平台费(AMZ)”，并将“销售税”的值移动到“销售税(AMZ)”
mask = main_file_df_2['平台'].str.contains('AMAZON', case=False, na=False)
main_file_df_2.loc[mask, '平台费(AMZ)'] = main_file_df_2.loc[mask, '平台费']
main_file_df_2.loc[mask, '销售税(AMZ)'] = main_file_df_2.loc[mask, '销售税']
# 重命名列
main_file_df_2 = main_file_df_2.rename(columns={'平台费': '平台费(非AMZ)'})
main_file_df_2 = main_file_df_2.rename(columns={'销售税': '销售税(非AMZ)'})

# 将“平台”列等于“amazon”的“平台费(非AMZ)”和“销售税(非AMZ)”中对应的值设置为0
mask = main_file_df_2['平台'].str.contains('AMAZON', case=False, na=False)
main_file_df_2.loc[mask, '平台费(非AMZ)'] = 0
main_file_df_2.loc[mask, '销售税(非AMZ)'] = 0

# 若 platform_shop.vat_rate = 0（映射VAT税 = 0），则 销售税(非AMZ) = 0
mask_vat0 = pd.to_numeric(main_file_df_2['映射VAT税'], errors='coerce') == 0
main_file_df_2.loc[mask_vat0, '销售税(非AMZ)'] = 0

# 将所有相关列的空值填充为0
main_file_df_2[['平台费(非AMZ)', '平台费(AMZ)', '销售税(非AMZ)', '销售税(AMZ)']] = main_file_df_2[
    ['平台费(非AMZ)', '平台费(AMZ)', '销售税(非AMZ)', '销售税(AMZ)']].fillna(0)

main_file_df_2['平台费合计'] = main_file_df_2['平台费(AMZ)'] + main_file_df_2['平台费(非AMZ)']
main_file_df_2['销售税合计'] = main_file_df_2['销售税(AMZ)'] + main_file_df_2['销售税(非AMZ)']

# 3 提现费
"""
如果 站点 == OTTO-BTH，则 提现费 = 销售额 * 0.03；
否则，提现费 = 平台销售额 * 0.01
"""
main_file_df_2['提现费'] = np.select(
    [
        main_file_df_2['站点'] == 'OTTO-BTH'
    ],
    [
        np.round(main_file_df_2['销售额'] * 0.03, 2)
    ],
    default=np.round(main_file_df_2['平台销售额'] * 0.01, 2)  # 其余所有站点
)

# 当销售额 = 0 时，把平台费、VAT、提现费强制设为 0
main_file_df_2.loc[
    main_file_df_2['销售额'] == 0, ['平台费(AMZ)', '平台费(非AMZ)', '平台费合计',
                                    '销售税(AMZ)', '销售税(非AMZ)', '销售税合计', '提现费']] = 0

# 按照 'SKU-站点识别码' 列进行分组，进行汇总
main_file_df_2 = main_file_df_2.groupby('SKU-站点识别码').agg({
    'SKU': 'first',  # 保留每组的第一行数据
    '站点': 'first',  # 保留每组的第一行数据
    '平台': 'first',  # 保留每组的第一行数据
    'SKU-平台识别码': 'first',  # 保留每组的第一行数据
    '分销': 'first',  # 保留每组的第一行数据
    '映射平台费（佣金）': 'first',  # 保留每组的第一行数据
    '映射佣金比': 'first',  # 保留每组的第一行数据
    '映射VAT税': 'first',  # 保留每组的第一行数据
    '平台销售额': 'sum',  # 汇总
    '头程': 'sum',  # 汇总
    '关税': 'sum',  # 汇总
    '派送费': 'sum',  # 汇总
    '销量': 'sum',  # 汇总
    '重发数量': 'sum',  # 汇总
    '订单采购成本': 'sum',  # 汇总
    '重发采购成本': 'sum',  # 汇总
    '退款额': 'sum',  # 汇总
    '退款数量': 'sum',  # 汇总
    '销售额': 'sum',  # 汇总
    '平台费(非AMZ)': 'sum',  # 汇总
    '销售税(非AMZ)': 'sum',  # 汇总
    '平台费(AMZ)': 'sum',  # 汇总
    '销售税(AMZ)': 'sum',  # 汇总
    '平台费合计': 'sum',  # 汇总
    '销售税合计': 'sum',  # 汇总
    '提现费': 'sum'  # 汇总
}).reset_index()

# 映射平台费（佣金） 为空 => 平台费(非AMZ)、平台费合计 置空
mask_null = main_file_df_2['映射平台费（佣金）'].isna()
main_file_df_2.loc[mask_null, ['平台费(非AMZ)', '平台费合计']] = np.nan


# 将「分销」列移到最后
_cols = [c for c in main_file_df_2.columns if c != '分销'] + ['分销']
main_file_df_2 = main_file_df_2[_cols]

# 保存结果
output_path = main_file_path.replace('已完成-7', '已完成-8')
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    main_file_df_2.to_excel(writer, index=False)
    worksheet = writer.sheets['Sheet1']
    fenxiao_col_idx = main_file_df_2.columns.get_loc('分销') + 1
    fenxiao_col_letter = openpyxl.utils.get_column_letter(fenxiao_col_idx)
    center_align = Alignment(horizontal='center')
    for cell in worksheet[fenxiao_col_letter]:
        cell.alignment = center_align
print(f'处理完成，output_path：{output_path}')

# 自动检查：平台 == castorama 的「映射平台费（佣金）」是否为空
_castorama = main_file_df_2['平台'].astype(str).str.lower() == 'castorama'
_empty_commission = _castorama & (
    main_file_df_2['映射平台费（佣金）'].isna()
    | main_file_df_2['映射平台费（佣金）'].astype(str).str.strip().isin(['', 'nan', 'None'])
)
if _empty_commission.any():
    merge_missing_into_castorama_commission_json(
        main_file_df_2, CASTORAMA_COMMISSION_PATH, log_tag="C5"
    )
    print(
        f'{Color.RED}检查失败：castorama 的「映射平台费（佣金）」存在空值，'
        f'共 {_empty_commission.sum()} 行；'
        f'请编辑 {CASTORAMA_COMMISSION_PATH} 填写「佣金比」后重跑本脚本{Color.RESET}'
    )
    print(main_file_df_2.loc[_empty_commission, ['SKU-站点识别码', 'SKU', '站点', '平台', '映射平台费（佣金）', '映射佣金比']].drop_duplicates().to_string(index=False))
    raise SystemExit(1)

    
print(f'{Color.GREEN}一切正常，请进行下一步操作{Color.RESET}')
