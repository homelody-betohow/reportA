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

from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name, fba_date
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-14)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)

# TODO 文件路径！！！  亚马逊上个月的仓租(实际：上2个月的利润报表)，按天数摊分给日报
fba_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\FBA仓租\(处理完成)FBA仓租明细{fba_date}.xlsx"
fba_file_df = pd.read_excel(fba_file_path)
# 上个月 总的 FBA仓租费
total_fba_fee = fba_file_df["FBA仓租费"].sum()

# 获取当前月份的天数
days_in_now_month = pd.Timestamp.now().daysinmonth
# 获取 当前日报的天数
d1, d2 = (int(part.split('.')[1]) for part in shared_date.split('-'))
days = d2 - d1 + 1

# 按 销量比例，用 上月总的FBA仓租费 去计算 日报 的  FBA仓租费
# 筛选 平台 包含 AMAZON
mask = main_file_df['平台'].str.contains('AMAZON', case=False, na=False)
# 计算销量占比（仅在 AMAZON 行）
amazon_sales = main_file_df.loc[mask, '销量']
# 得到 当前月份、按销量比例 * 上月的FBA仓租费 的   一天的FBA仓租费
main_file_df.loc[mask, 'FBA仓租费'] = np.round(amazon_sales / amazon_sales.sum() * total_fba_fee / days_in_now_month * days, 2)
# 空值的地方——补 0
main_file_df = main_file_df.fillna(0)
# 保存修改
output_path = main_file_path.replace('已完成-14', '已完成-15')
main_file_df.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
print('太小的FBA仓租费（日租），保留2位小数后，会变成0！！！！')
