import pandas as pd
import importlib.util
from pathlib import Path
# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from config.A0_set_date import shared_date, folder_name, fba_date
from config.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-14)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)

# TODO 文件路径！！！
fba_rent_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\FBA仓租\(处理完成)FBA仓租明细{fba_date}.xlsx"
fba_rent_df = pd.read_excel(fba_rent_path)

# 以站点商品ID识别码为键进行合并，选择左连接（left join），这样可以确保表1的所有数据都被保留
result_df = pd.merge(main_file_df, fba_rent_df[['站点商品ID识别码', 'FBA仓租费']], on='站点商品ID识别码', how='left')

# 找出表2中在表1中不存在的行
missing_rows = fba_rent_df[~fba_rent_df['站点商品ID识别码'].isin(main_file_df['站点商品ID识别码'])]

# 将这些缺失的行添加到结果中
result_df = pd.concat([result_df, missing_rows], ignore_index=True)

# 确保所有期望的列都存在
expected_columns = list(main_file_df.columns) + ['FBA仓租费']
for col in expected_columns:
    if col not in result_df.columns:
        result_df[col] = None  # 如果列不存在，添加该列并填充为 None

# 对于表1中没有的行，将新增的列填充为0
result_df[['FBA仓租费']] = result_df[['FBA仓租费']].fillna(0)

# 重新排序，确保列的顺序符合要求
result_df = result_df[expected_columns]

# 保存结果到新的 Excel 文件
output_path = main_file_path.replace('已完成-14', '已完成-15')
result_df.to_excel(output_path, index=False)
print(f"结果已保存到文件：{output_path}")
