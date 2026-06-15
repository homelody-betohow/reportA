import warnings
import numpy as np
import pandas as pd
from pathlib import Path
import importlib.util
# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT, SELLERSKU_PROFIT_FILE_NAME, SELLERSKU_PROFIT_REPORT_DIR

# 忽略 FutureWarning
warnings.filterwarnings("ignore", category=FutureWarning)

# TODO 文件路径！！！
# 判断是否有 二次上架费用，自动选择文件路径
# 有  二次上架费用  的路径
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-12)订单统计-{shared_date}.xlsx"
# 判断是否有 二次上架费用
if not Path(main_file_path).is_file():
    print(f"文件不存在，无  二次上架费用！！！")
    # 有  测评费  的路径
    main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-11)订单统计-{shared_date}.xlsx"
    if not Path(main_file_path).is_file():
        print(f"文件不存在，无  测评费！！！")
        # 有  秒杀费  的路径
        main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-10)订单统计-{shared_date}.xlsx"
        if not Path(main_file_path).is_file():
            print(f"文件不存在，无  秒杀费！！！")
            main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-9)订单统计-{shared_date}.xlsx"

main_file_df = pd.read_excel(main_file_path)
# 检查 下面这些 列名 是否存在，不存在，则：新增列，数据为 0
cols_to_check = ['秒杀费', '测评费', '二次上架数量', '二次上架金额', '二次上架采购成本']
for col in cols_to_check:
    if col not in main_file_df.columns:
        main_file_df[col] = 0
        print(f"新增一列:{col}，数据全是 0")

if '采购成本' not in main_file_df.columns:
    main_file_df['采购成本'] = main_file_df['订单采购成本'] + main_file_df['重发采购成本'] - main_file_df[
        '二次上架采购成本']

# TODO 文件路径！！！
guang_gao_path = fr"{SELLERSKU_PROFIT_REPORT_DIR}\(处理完成){SELLERSKU_PROFIT_FILE_NAME}"
guang_gao_df = pd.read_excel(guang_gao_path)

# 以儿子-站点识别码为键进行合并，选择左连接（left join），这样可以确保表1的所有数据都被保留
result_df = pd.merge(main_file_df, guang_gao_df[
    ['儿子-站点识别码', '广告费(AMZ)', '赔偿金额', '其他分摊费用', 'EU-其他分摊费用-需要分摊的',
     'US-其他分摊费用-需要分摊的']], on='儿子-站点识别码', how='left')

# 找出表2中在表1中不存在的行
missing_rows = guang_gao_df[~guang_gao_df['儿子-站点识别码'].isin(main_file_df['儿子-站点识别码'])]

# 将这些缺失的行添加到结果中
result_df = pd.concat([result_df, missing_rows], ignore_index=True)

# 确保所有期望的列都存在
expected_columns = list(main_file_df.columns) + ['广告费(AMZ)', '赔偿金额', '其他分摊费用',
                                                 'EU-其他分摊费用-需要分摊的', 'US-其他分摊费用-需要分摊的']
for col in expected_columns:
    if col not in result_df.columns:
        result_df[col] = None  # 如果列不存在，添加该列并填充为 None

# 对于表1中没有的行，将新增的列填充为0
result_df[expected_columns] = result_df[expected_columns].fillna(0)
# 重新排序，确保列的顺序符合要求
result_df = result_df[expected_columns]

# 空值的地方——补 0
result_df = result_df.fillna(0)

# TODO 按照”订单统计“中的 AMAZON-EU、AMAZON-US ‘平台’的 ‘平台销售额‘
#  去分摊  “利润报表”中的‘EU-其他分摊费用-需要分摊的’、‘US-其他分摊费用-需要分摊的’
# 1. 先把原“其他分摊费用”保存下来
result_df = result_df.rename(columns={'其他分摊费用': '原-其他分摊费用'})

# 2. 待摊总额
total_eu_alloc = result_df['EU-其他分摊费用-需要分摊的'].sum()
total_us_alloc = result_df['US-其他分摊费用-需要分摊的'].sum()

# 3. 清空
result_df['EU-其他分摊费用-需要分摊的'] = 0
result_df['US-其他分摊费用-需要分摊的'] = 0

# 4. EU 分摊：平台 == "AMAZON-EU"
eu_df = result_df[result_df['平台'] == 'AMAZON-EU'].copy()
total_eu_sales = eu_df['平台销售额'].sum()
eu_df['EU-其他分摊费用-需要分摊的'] = np.round(
    eu_df['平台销售额'] / total_eu_sales * total_eu_alloc, 2)
result_df.loc[result_df['平台'] == 'AMAZON-EU',
'EU-其他分摊费用-需要分摊的'] = eu_df['EU-其他分摊费用-需要分摊的'].values

# 5. US 分摊：平台 == "AMAZON-US"
us_df = result_df[result_df['平台'] == 'AMAZON-US'].copy()
total_us_sales = us_df['平台销售额'].sum()
us_df['US-其他分摊费用-需要分摊的'] = np.round(
    us_df['平台销售额'] / total_us_sales * total_us_alloc, 2)
result_df.loc[result_df['平台'] == 'AMAZON-US',
'US-其他分摊费用-需要分摊的'] = us_df['US-其他分摊费用-需要分摊的'].values

# 6. 合并得到最终“其他分摊费用”
result_df['其他分摊费用'] = (
        result_df['原-其他分摊费用']
        + result_df['EU-其他分摊费用-需要分摊的']
        + result_df['US-其他分摊费用-需要分摊的']
)

# 计算 广告费合计
result_df['广告费合计'] = result_df['广告费(AMZ)'] + result_df['广告费(非AMZ)']

# 保存结果到新的 Excel 文件
output_path = main_file_path.replace('已完成-12', '已完成-13')
# 没有 二次上架费用 的话， 则是：11 直接跳到 13
output_path = output_path.replace('已完成-11', '已完成-13')
# 没有二次上架费用、测评费的话， 则是：10 直接跳到 13
output_path = output_path.replace('已完成-10', '已完成-13')
# 没有二次上架费用、测评费、秒杀费的话，则是：9 直接跳到 13
output_path = output_path.replace('已完成-9', '已完成-13')
result_df.to_excel(output_path, index=False)
print(f"结果已保存到文件：{output_path}")
