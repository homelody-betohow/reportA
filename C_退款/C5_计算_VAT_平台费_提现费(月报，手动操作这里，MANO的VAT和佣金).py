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

from A_报表.Z_method.style import Color
from A_报表.Z_method.sku_映射 import sku_mappings
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-7)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)
# 重命名
main_file_df = main_file_df.rename(columns={'映射VAT税': 'amazon-VAT税'})

# 映射 castorama 的 佣金比例，映射不到的问：晓佳
product_map_sku_path = fr"{DESKTOP_ROOT}\castorama - SKU类目佣金比例.xlsx"
main_file_df = sku_mappings(
    main_df=main_file_df,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="SKU",
    map_new_sku="佣金比",
    map_sku_sheet='Sheet1'
)
# 映射 平台费（佣金）
product_map_sku_path = fr"{DESKTOP_ROOT}\VAT、平台费-映射.xlsx"
main_file_df_1 = sku_mappings(
    main_df=main_file_df,
    main_sku='站点',
    map_sku_path=product_map_sku_path,
    map_old_sku="站点",
    map_new_sku="平台费（佣金）",
    map_sku_sheet='VAT税、佣金'
)
# 用“映射佣金比”填补“映射平台费（佣金）”的空值
main_file_df_1['映射平台费（佣金）'] = main_file_df_1['映射平台费（佣金）'].fillna(main_file_df_1['映射佣金比'])
# 映射 VAT税
main_file_df_2 = sku_mappings(
    main_df=main_file_df_1,
    main_sku='站点',
    map_sku_path=product_map_sku_path,
    map_old_sku="站点",
    map_new_sku="VAT税",
    map_sku_sheet='VAT税、佣金'
)

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

# 按照 '儿子-站点识别码' 列进行分组，进行汇总
main_file_df_2 = main_file_df_2.groupby('儿子-站点识别码').agg({
    'SKU': 'first',  # 保留每组的第一行数据
    '站点': 'first',  # 保留每组的第一行数据
    '平台': 'first',  # 保留每组的第一行数据
    '儿子-平台识别码': 'first',  # 保留每组的第一行数据
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

# 保存结果
output_path = main_file_path.replace('已完成-7', '已完成-8')
main_file_df_2.to_excel(output_path, index=False)
print(f'处理完成，output_path：{output_path}')

# 自动检查：平台 == castorama 的「映射平台费（佣金）」是否为空
_castorama = main_file_df_2['平台'].astype(str).str.lower() == 'castorama'
_empty_commission = _castorama & (
    main_file_df_2['映射平台费（佣金）'].isna()
    | main_file_df_2['映射平台费（佣金）'].astype(str).str.strip().isin(['', 'nan', 'None'])
)
if _empty_commission.any():
    print(
        f'{Color.RED}检查失败：castorama 的「映射平台费（佣金）」存在空值，'
        f'共 {_empty_commission.sum()} 行；请联系陈晓佳补充'
        f'「castorama - SKU类目佣金比例.xlsx」后重跑本脚本{Color.RESET}'
    )
    print(main_file_df_2.loc[_empty_commission, ['儿子-站点识别码', 'SKU', '站点', '平台', '映射平台费（佣金）', '映射佣金比']].drop_duplicates().to_string(index=False))
    raise SystemExit(1)
print(f'{Color.GREEN}一切正常，请进行下一步操作{Color.RESET}')
