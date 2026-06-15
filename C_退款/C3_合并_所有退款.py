import re
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
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT, SELLERSKU_PROFIT_FILE_NAME, SELLERSKU_PROFIT_REPORT_DIR

# TODO 文件路径！！！
RMA_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\RMA\(处理完成-无Amazon)RMA-{shared_date}.xlsx"
RMA_file_df = pd.read_excel(RMA_file_path)

# TODO 文件路径！！！
amazon_file_path = fr"{SELLERSKU_PROFIT_REPORT_DIR}\(处理完成-退款){SELLERSKU_PROFIT_FILE_NAME}"
amazon_file_df = pd.read_excel(amazon_file_path)
# 过滤掉全为 NA 的列
RMA_file_df = RMA_file_df.dropna(axis=1, how='all')
amazon_file_df = amazon_file_df.dropna(axis=1, how='all')
# 合并 两个退款表格
main_df = pd.concat([RMA_file_df, amazon_file_df], ignore_index=True)
# SKU、儿子-站点识别码、儿子-平台识别码 去掉尾缀 -1、-2、-3、-4、-5、-6、-7、-8
main_df['SKU'] = main_df['SKU'].apply(
    lambda x: re.sub(r'(-1|-2|-3|-4|-5|-6|-7|-8)$', '', x) if re.search(r'(-1|-2|-3|-4|-5|-6|-7|-8)$', x) else x)
main_df['儿子-站点识别码'] = main_df['儿子-站点识别码'].apply(
    lambda x: re.sub(r'(-1|-2|-3|-4|-5|-6|-7|-8)$', '', x) if re.search(r'(-1|-2|-3|-4|-5|-6|-7|-8)$', x) else x)
main_df['儿子-平台识别码'] = main_df['儿子-平台识别码'].apply(
    lambda x: re.sub(r'(-1|-2|-3|-4|-5|-6|-7|-8)$', '', x) if re.search(r'(-1|-2|-3|-4|-5|-6|-7|-8)$', x) else x)

# 按照 '儿子-站点识别码' 列进行分组，汇总
main_df_1 = main_df.groupby('儿子-站点识别码').agg({
    'SKU': 'first',  # 保留每组的第一行数据
    '站点': 'first',  # 保留每组的第一行数据
    '平台': 'first',  # 保留每组的第一行数据
    '儿子-平台识别码': 'first',  # 保留每组的第一行数据
    '退款数量': 'sum',
    '退款额': 'sum',
    '销售退款金额VAT-amazon': 'sum',
    '销售退款金额的佣金': 'sum',
}).reset_index()

main_df_1['退款额'] = np.round(main_df_1['退款额'], 2)

output_path = RMA_file_path.rsplit('\\', 1)[0] + '\\(处理完成)所有平台-退款.xlsx'
main_df_1.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
