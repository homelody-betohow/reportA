import pandas as pd
import importlib.util
from pathlib import Path
# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-6)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)

# TODO 文件路径！！！
RMA_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\RMA\(处理完成)所有平台-退款.xlsx"
RMA_file_df = pd.read_excel(RMA_file_path)

# 去除 整张表 的前后空格
for col in main_file_df.columns:
    main_file_df[col] = main_file_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

# 以SKU-站点识别码为键进行合并，选择左连接（left join），这样可以确保表1的所有数据都被保留
result_df = pd.merge(main_file_df, RMA_file_df[
    ['SKU-站点识别码', '退款额', '退款数量', '销售退款金额VAT-amazon', '销售退款金额的佣金']], on='SKU-站点识别码',
                     how='left')

# 找出表2中在表1中不存在的行
missing_rows = RMA_file_df[~RMA_file_df['SKU-站点识别码'].isin(main_file_df['SKU-站点识别码'])]

# 将这些缺失的行添加到结果中
result_df = pd.concat([result_df, missing_rows], ignore_index=True)

# 确保所有期望的列都存在
expected_columns = list(main_file_df.columns) + ['退款额', '退款数量', '销售退款金额VAT-amazon', '销售退款金额的佣金']
if '分销' not in expected_columns:
    expected_columns.append('分销')
for col in expected_columns:
    if col not in result_df.columns:
        result_df[col] = None  # 如果列不存在，添加该列并填充为 None

# 表1中没有和表2相同的SKU-站点识别码，将新增的两列在对应行单元格填充为0
result_df[['退款额', '退款数量', '销售退款金额VAT-amazon', '销售退款金额的佣金']] = result_df[
    ['退款额', '退款数量', '销售退款金额VAT-amazon', '销售退款金额的佣金']].fillna(0)
# 重新排序，确保列的顺序符合要求
result_df = result_df[expected_columns]

# 空值的地方——补 0（分销列除外）
if '分销' not in result_df.columns:
    result_df['分销'] = '否'
else:
    result_df['分销'] = result_df['分销'].replace({0: '否', '0': '否'}).fillna('否')
_fill_cols = [c for c in result_df.columns if c != '分销']
result_df[_fill_cols] = result_df[_fill_cols].fillna(0)
# 将「分销」列移到最后
_cols = [c for c in result_df.columns if c != '分销'] + ['分销']
result_df = result_df[_cols]

result_df['销售额'] = result_df['平台销售额'] - result_df['退款额']

# 保存结果到新的 Excel 文件
output_path = main_file_path.replace('已完成-6', '已完成-7')
result_df.to_excel(output_path, index=False)
print(f"结果已保存到文件：{output_path}")
