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

# 表1
# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-8)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)

# 表2
# TODO 文件路径！！！
guang_gao_path = fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\广告\(处理完成)所有平台广告费用.xlsx'
guang_gao_df = pd.read_excel(guang_gao_path)

# 以SKU-站点识别码为键进行合并，选择左连接（left join），这样可以确保表1的所有数据都被保留
result_df = pd.merge(main_file_df, guang_gao_df[['SKU-站点识别码', '广告费(非AMZ)']], on='SKU-站点识别码', how='left')

# 找出表2中在表1中不存在的行
missing_rows = guang_gao_df[~guang_gao_df['SKU-站点识别码'].isin(main_file_df['SKU-站点识别码'])]

# 将这些缺失的行添加到结果中
result_df = pd.concat([result_df, missing_rows], ignore_index=True)

# 空值的地方——补 0（分销列除外）
if '分销' not in result_df.columns:
    result_df['分销'] = '否'
else:
    result_df['分销'] = result_df['分销'].replace({0: '否', '0': '否'}).fillna('否')
_fill_cols = [c for c in result_df.columns if c != '分销']
result_df[_fill_cols] = result_df[_fill_cols].fillna(0)

# 确保所有期望的列都存在
expected_columns = list(main_file_df.columns) + ['广告费(非AMZ)']
for col in expected_columns:
    if col not in result_df.columns:
        result_df[col] = None  # 如果列不存在，添加该列并填充为 None

# 表1中没有和表2相同的SKU-站点识别码，将新增的两列在对应行单元格填充为0
result_df[['广告费(非AMZ)']] = result_df[['广告费(非AMZ)']].fillna(0)
# 重新排序，确保列的顺序符合要求
result_df = result_df[expected_columns]

# 保存结果到新的 Excel 文件
output_path = main_file_path.replace('已完成-8', '已完成-9')
result_df.to_excel(output_path, index=False)
print(f"结果已保存到文件：{output_path}")
