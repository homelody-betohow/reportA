import pandas as pd
import importlib.util
from pathlib import Path
# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.A0_设置_时间段.A0_paths import SELLERSKU_PROFIT_FILE_NAME, SELLERSKU_PROFIT_REPORT_DIR

main_file_path = fr"{SELLERSKU_PROFIT_REPORT_DIR}\(已完成-1){SELLERSKU_PROFIT_FILE_NAME}"
main_file_df = pd.read_excel(main_file_path)
# 重命名
main_file_df = main_file_df.rename(columns={'映射站点': '站点'})
main_file_df = main_file_df.rename(columns={'映射平台': '平台'})
main_file_df = main_file_df.rename(columns={'计算结果-广告费': '广告费(AMZ)'})  # 全是 负的
main_file_df = main_file_df.rename(columns={'计算结果-赔偿': '赔偿金额'})  # 全是 正数，不用处理
main_file_df = main_file_df.rename(columns={'计算结果-其他分摊费用': '其他分摊费用'})  # 有正，有负
# 去掉 负号
main_file_df['广告费(AMZ)'] = main_file_df['广告费(AMZ)'].apply(
    lambda x: abs(x) if isinstance(x, (int, float)) else x)
# 负数变正数; 正数变负数
s = main_file_df['其他分摊费用']
main_file_df['其他分摊费用'] = pd.to_numeric(s).mul(-1).fillna(s)

# 按照 'SKU-站点识别码' 列进行分组，进行汇总
main_file_df_1 = main_file_df.groupby('SKU-站点识别码').agg({
    '广告费(AMZ)': 'sum',
    '赔偿金额': 'sum',
    '其他分摊费用': 'sum',
    'SKU': 'first',  # 保留每组的第一行数据
    '站点': 'first',  # 保留每组的第一行数据
    '平台': 'first',  # 保留每组的第一行数据
    'SKU-平台识别码': 'first'  # 保留每组的第一行数据
}).reset_index()

# 保留目标列
main_file_df_1 = main_file_df_1[
    ['SKU', '站点', '平台', 'SKU-站点识别码', 'SKU-平台识别码', '广告费(AMZ)', '赔偿金额', '其他分摊费用']]

# 筛选出指定列相加不为 0 的行
columns_to_check = ['广告费(AMZ)', '赔偿金额', '其他分摊费用']
main_file_df_2 = main_file_df_1[~(main_file_df_1[columns_to_check].sum(axis=1) == 0)]

# SKU != "无" 的行
main_file_df_3 = main_file_df_2[main_file_df_2['SKU'] != '无'].copy()

# SKU == "无" 的行
main_file_df_NO = main_file_df_2[main_file_df_2['SKU'] == '无'].copy()
# 按平台分组，计算各平台的“其他分摊费用”合计
platform_sum = main_file_df_NO.groupby('平台')['其他分摊费用'].sum()
# 提取 EU 和 US 的合计（如果不存在则为 0）
EU_sum = platform_sum.get('AMAZON-EU', 0)
US_sum = platform_sum.get('AMAZON-US', 0)
# 新增两列：EU-其他分摊费用-需要分摊的、US-其他分摊费用-需要分摊的
main_file_df_3['EU-其他分摊费用-需要分摊的'] = pd.NA
main_file_df_3['US-其他分摊费用-需要分摊的'] = pd.NA
# 将合计值填入第一行
main_file_df_3.iloc[0, main_file_df_3.columns.get_loc('EU-其他分摊费用-需要分摊的')] = EU_sum
main_file_df_3.iloc[0, main_file_df_3.columns.get_loc('US-其他分摊费用-需要分摊的')] = US_sum

output_path = fr"{SELLERSKU_PROFIT_REPORT_DIR}\(处理完成){SELLERSKU_PROFIT_FILE_NAME}"
main_file_df_3.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
