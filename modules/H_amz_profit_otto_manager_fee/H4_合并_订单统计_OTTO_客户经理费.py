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

from config.A0_set_date import *
from config.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-13)订单统计-{shared_date}.xlsx"
main_df = pd.read_excel(main_file_path)

main_df_1 = main_df.copy()
# 1. 计算总的分摊费用
# 获取 当前报表的天数
d1, d2 = (int(part.split('.')[1]) for part in shared_date.split('-'))
days = d2 - d1 + 1
# 总分摊费用（一个月 540 的 客户经理费用）
total_cost = 540
if folder_name == '日报':
    # 获取 当前年份
    current_year = datetime.now().year
    # 获取 日报月份
    month = int(shared_date.split('.')[0])
    # 获取 日报 所在月的天数（当前年）
    days_in_month = pd.Timestamp(year=current_year, month=month, day=1).daysinmonth
    # 日报则按天数占比 * 540
    total_cost = 540 * days / days_in_month
print(f'OTTO的客户经理费用：{total_cost}')
# 2. 筛选OTTO平台数据
otto_mask = main_df_1['平台'] == 'OTTO'
otto_data = main_df_1[otto_mask]
# 计算OTTO平台总销量
total_otto_sales = otto_data['销量'].sum()
# 3. 按销量占比分摊费用
# 计算每个OTTO记录的销量占比
otto_sales_ratio = otto_data['销量'] / total_otto_sales
# 计算每个OTTO记录应分摊的费用
allocated_cost = np.round(otto_sales_ratio * total_cost, 2)
# 更新其他分摊费用列
main_df_1.loc[otto_mask, '其他分摊费用'] = allocated_cost

# 保存结果到新的 Excel 文件
output_path = main_file_path.replace('已完成-13', '已完成-13-1')
main_df_1.to_excel(output_path, index=False)
print(f"结果已保存到文件：{output_path}")
