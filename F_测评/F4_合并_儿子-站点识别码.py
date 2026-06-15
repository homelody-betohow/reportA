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

from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！
test_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\测评表\(已完成-4)测评表.xlsx"
test_file_df = pd.read_excel(test_file_path)

test_file_df = test_file_df.rename(columns={'映射平台': '平台'})
# 按照 '儿子-站点识别码' 列进行分组，并对 '测评费' 列进行汇总
grouped_df = test_file_df.groupby('儿子-站点识别码').agg({
    '测评费': 'sum',  # 汇总 测评费
    'SKU': 'first',  # 保留每组的第一行数据
    '站点': 'first',  # 保留每组的第一行数据
    '平台': 'first',  # 保留每组的第一行数据
    '儿子-平台识别码': 'first'  # 保留每组的第一行数据
}).reset_index()

grouped_df['测评费'] = np.round(grouped_df['测评费'], 2)

# 添加币种列，默认值为 EUR
grouped_df['币种'] = 'EUR'

grouped_df = grouped_df[
    ['SKU', '站点', '平台', '儿子-站点识别码', '儿子-平台识别码', '测评费', '币种']]

# 删除"测评费"列为测评费的行
grouped_df = grouped_df[grouped_df['测评费'] != 0]

# 输出文件路径
output_file_path = test_file_path.replace('(已完成-4)', '(处理完成)')
# 将结果保存到新的Excel文件
grouped_df.to_excel(output_file_path, index=False)

print(f"处理完成，结果已保存到{output_file_path}")
