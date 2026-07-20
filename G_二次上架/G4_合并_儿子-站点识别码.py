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
main_file_path = fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\二次上架\(已完成-2)鸿羽仓-二次上架明细-{shared_date}.xlsx'
main_df = pd.read_excel(main_file_path)

# 返还采购成本：物流退件时等于二次上架采购成本，其余为 0
main_df['返还采购成本'] = np.where(
    main_df['退件类型'] == '物流退件/Logistics',
    main_df['二次上架采购成本（RMB）'],
    0
)

# 按照 'SKU-站点识别码' 列进行分组，并对 '数量' 和 '退件费用(EUR)' 列进行汇总
result_df = main_df.groupby('SKU-站点识别码').agg({
    '实收数量': 'sum',  # 汇总 数量
    '退件费用(EUR)': 'sum',  # 汇总 退件费用
    '二次上架采购成本（RMB）': 'sum',  # 汇总 二次上架采购成本
    '返还采购成本': 'sum',  # 汇总 返还采购成本
    'SKU': 'first',  # 保留每组的第一行数据
    '站点': 'first',  # 保留每组的第一行数据
    '平台': 'first',  # 保留每组的第一行数据
    'SKU-平台识别码': 'first'  # 保留每组的第一行数据
}).reset_index()
# 重命名
result_df = result_df.rename(columns={'实收数量': '二次上架数量'})
result_df = result_df.rename(columns={'退件费用(EUR)': '二次上架金额'})
result_df = result_df.rename(columns={'二次上架采购成本（RMB）': '二次上架采购成本'})
# 保留2位小数
result_df['二次上架金额'] = np.round(result_df['二次上架金额'], 2)
result_df['二次上架采购成本'] = np.round(result_df['二次上架采购成本'], 2)
result_df['返还采购成本'] = np.round(result_df['返还采购成本'], 2)

result_df = result_df[
    ['SKU', '站点', '平台', 'SKU-站点识别码', 'SKU-平台识别码', '二次上架数量', '二次上架金额', '二次上架采购成本', '返还采购成本']]

# 输出文件路径
output_file_path = main_file_path.replace('(已完成-2)', '(处理完成)')
# 将结果保存到新的Excel文件
result_df.to_excel(output_file_path, index=False)

print(f"处理完成，结果已保存到{output_file_path}")
