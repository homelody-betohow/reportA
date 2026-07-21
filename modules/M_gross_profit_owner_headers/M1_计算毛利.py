import numpy as np
import pandas as pd
import importlib.util
from pathlib import Path
# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from config.A0_set_date import shared_date, folder_name, USD_to_EUR, SKU_NW_DISCOUNT
from config.A0_paths import DESKTOP_ROOT
from common.style import Color

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-19)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)

# 删掉 月租 的计算列
df_name_list = main_file_df.columns.tolist()  # 表头
if '月租-整月' in df_name_list:
    main_file_df.drop(columns=['月租-站点识别', '月租-整月'], inplace=True)
else:
    main_file_df.drop(columns=['月租-站点识别'], inplace=True)

# 需要调整的列
cost_cols = ['采购成本']    # 目前是EUR
# 易速 的成本价是 美元。这里先 *7.3 变回 美元，再 * USD_to_EUR(美元转欧元)
mask = main_file_df['供应商'] == '易速'
main_file_df.loc[mask, cost_cols] *= 7.3 * USD_to_EUR
# 不需要 * 1.13 的 SKU（以这些SKU开头）
exclude_skus = ('U56033001', 'U56032001', 'U56031001', 'E56033001', 'E56033101', 'E56033201', 'E56032001', 'E56032101')
# 条件1：SKU 以指定 exclude_skus 开头 且 平台包含 AMAZON
sku_condition = (
    main_file_df['SKU'].str.startswith(exclude_skus, na=False) &
    main_file_df['平台'].str.contains('AMAZON', na=False, case=False)
)
# 条件2：平台包含 AMAZON 且 站点包含 FB 或 NF（店铺：fenbin、hknovaflow）
platform_site_condition = (
    (main_file_df['平台'].str.contains('AMAZON', na=False, case=False)) &
    (main_file_df['站点'].str.contains('FB|NF', na=False, case=False))
)
# 合并“不需要 * 1.13”的条件（满足任一条件就不乘1.13）
no_adjust_mask = sku_condition | platform_site_condition
# 对“需要 * 1.13”的行（no_adjust_mask 为 False）做 *1.13 并保留两位小数
main_file_df.loc[~no_adjust_mask, cost_cols] = np.round(main_file_df.loc[~no_adjust_mask, cost_cols] * 1.13, 2)


# # 2026-06-05 调整：SKU以 -NW 结尾的
nw_suffix_mask = main_file_df['SKU'].astype(str).str.endswith('-NW')
main_file_df.loc[nw_suffix_mask, cost_cols] *= SKU_NW_DISCOUNT
print(f"{Color.YELLOW}2026-06-05 调整：SKU以 -NW 结尾的，采购成本 打折：{SKU_NW_DISCOUNT}{Color.RESET} \n")

# 参与毛利计算的列先补 0，避免 NaN（如月租未摊分）导致整行毛利变 NaN 后被 fillna 清零
_gross_profit_cols = (
    '销售额', '测评费', '秒杀费', '广告费合计', '平台费合计', '销售税合计',
    '派送费', '提现费', '赔偿金额', '其他分摊费用', '二次上架金额',
    '采购成本', '头程', '关税', '仓租合计', '月租',
)
for _col in _gross_profit_cols:
    if _col in main_file_df.columns:
        main_file_df[_col] = main_file_df[_col].fillna(0)

# 计算毛利列
# main_file_df['毛利'] = main_file_df['销售额'] - main_file_df['测评费'] - main_file_df[
#     '秒杀费'] - main_file_df['广告费合计'] - main_file_df['平台费合计'] - main_file_df['销售税合计'] - main_file_df[
#                            '派送费'] - main_file_df['提现费'] + main_file_df['赔偿金额'] - main_file_df[
#                            '其他分摊费用'] - main_file_df['二次上架金额'] - main_file_df['采购成本'] - \
#                        main_file_df['头程'] - main_file_df['关税'] - main_file_df['仓租合计'] - main_file_df['月租']

# 2026-06-05 调整：二次上架金额 不计入 毛利
main_file_df['毛利'] = main_file_df['销售额'] - main_file_df['测评费'] - main_file_df[
    '秒杀费'] - main_file_df['广告费合计'] - main_file_df['平台费合计'] - main_file_df['销售税合计'] - main_file_df[
                           '派送费'] - main_file_df['提现费'] + main_file_df['赔偿金额'] - main_file_df[
                           '其他分摊费用'] - main_file_df['采购成本'] - \
                       main_file_df['头程'] - main_file_df['关税'] - main_file_df['仓租合计'] - main_file_df['月租']

# 二次上架金额&二次上架采购成本 与销售日报无关，所以设置为 0， 公司承担
print(f"{Color.GREEN}{Color.BOLD}(二次上架金额 & 二次上架采购成本) 统一设置为 0{Color.RESET} \n")
main_file_df['二次上架金额'] = 0
main_file_df['二次上架采购成本'] = 0

# 供应商 为 智慧谷，且 销售额 不为 0（这样 无销售额的，毛利就会显示实际的花费）；毛利 固定为 销售额的 5%
main_file_df.loc[(main_file_df['供应商'] == '智慧谷') & (main_file_df['销售额'] != 0), '毛利'] = main_file_df[
                                                                                                     '销售额'] * 0.05
main_file_df['毛利'] = np.round(main_file_df['毛利'], 2)
# 计算  毛利率
main_file_df['毛利率'] = (main_file_df['毛利'] / main_file_df['销售额']).apply(lambda x: f"{x:.2%}")  # 换为百分比格式
# 使用正则表达式替换“毛利率”列中包含“inf%”、“-inf%”、nan% 的值为“0.00%”
main_file_df['毛利率'] = main_file_df['毛利率'].str.replace(r'[-+]?inf%|nan%', '0.00%', regex=True)
# 销售额 <= 0 ，则 毛利率 = 0.00%
main_file_df['毛利率'] = main_file_df.apply(
    lambda row: f"{0.00:.2f}%" if row['销售额'] <= 0 else f"{float(str(row['毛利率']).rstrip('%')):.2f}%", axis=1)

# 空值的地方——补 0
main_file_df = main_file_df.fillna(0)
# 去除 整张表 的前后空格
for col in main_file_df.columns:
    main_file_df[col] = main_file_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

# 保存结果到新的 Excel 文件
output_path = main_file_path.replace('已完成-19', '已完成-20')
main_file_df.to_excel(output_path, index=False)
print(f"结果已保存到 {output_path}")
