import importlib.util
import warnings
from pathlib import Path

# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

import pandas as pd
from config.A0_set_date import shared_date, folder_name, transaction_date
from config.A0_paths import DESKTOP_ROOT

# 忽略警告
warnings.filterwarnings("ignore", category=UserWarning)


# transaction 导出表头行（新格式前 3 行为筛选条件，第 4 行为列名）
_TRANSACTION_HEADER_KEYS = frozenset({"order id", "seller sku", "fba fees"})


def _find_transaction_header_row(xlsx_path, max_scan=12):
    """定位 transaction 表头行（兼容新旧 ERP 导出：header=0 或 header=3）。"""
    preview = pd.read_excel(xlsx_path, header=None, nrows=max_scan, engine="openpyxl")
    for i in range(len(preview)):
        row_vals = {
            str(v).replace("\n", " ").strip()
            for v in preview.iloc[i]
            if pd.notna(v) and str(v).strip()
        }
        if _TRANSACTION_HEADER_KEYS.issubset(row_vals):
            return i
    return 0


def _read_transaction_excel(xlsx_path):
    header_row = _find_transaction_header_row(xlsx_path)
    df = pd.read_excel(xlsx_path, header=header_row, engine="openpyxl")
    df.columns = [("" if c is None else str(c)).replace("\n", " ").strip() for c in df.columns]
    return df.dropna(how="all")


# 定义 sku 提取规则
def extract_values(s):
    if pd.isna(s):  # 检查是否为 NaN
        return None  # 如果是 NaN，返回 None 或其他默认值
    if 'amzn.gr.' in s:
        return s.split(r'amzn.gr.')[-1].split('-')[0].split('_')[0]
    else:
        return s.split('#')[0].split('BCFBAFL')[0]

print(f'{transaction_date}')
# 读取两个文件
transaction_path_1 = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\transaction交易明细\transaction交易明细-已发放订单{transaction_date}.xlsx"
transaction_path_2 = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\transaction交易明细\transaction交易明细-已推迟订单{transaction_date}.xlsx"
transaction_df_1 = _read_transaction_excel(transaction_path_1)
transaction_df_2 = _read_transaction_excel(transaction_path_2)

# 合并（纵向拼接）
merged_df = pd.concat([transaction_df_1, transaction_df_2], ignore_index=True)

# 去除 整张表 的前后空格
for col in merged_df.columns:
    merged_df[col] = merged_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

# 应用提取规则，清洗 仓库sku
merged_df['seller sku'] = merged_df['seller sku'].apply(extract_values)
merged_df_1 = merged_df.rename(columns={'seller sku': 'SKU'})

# 在 order id 后插入新列 order-id识别码
new_column_name = "order-id识别码"  # 新列名
new_column_data = merged_df_1["order id"] + merged_df_1["SKU"]  # 新列数据
target_column = "order id"  # 目标列名（在其后插入）
insert_position = merged_df_1.columns.get_loc(target_column) + 1  # 计算插入位置
merged_df_1.insert(insert_position, new_column_name, new_column_data)  # 插入新列

# 筛选 order id 列不为空、fba fees 列不为0的行
merged_df_2 = merged_df_1[merged_df_1['order id'].notna() & (merged_df_1['fba fees'] != 0)].copy()
# fba fees的正数变负数，负数变正数
merged_df_2['fba fees'] = merged_df_2['fba fees'].apply(lambda x: -x)

# 保留指定列
merged_df_2 = merged_df_2[["order id", "SKU", "order-id识别码", "fba fees"]]
# 按照 'order-id识别码' 列进行分组，进行汇总
merged_df_3 = merged_df_2.groupby('order-id识别码').agg({
    'order id': 'first',  # 保留每组的第一行数据
    'SKU': 'first',  # 保留每组的第一行数据
    'fba fees': 'sum'  # 汇总
}).reset_index()

# 保存结果
output_path = transaction_path_1.rsplit('\\', 1)[
                  0] + fr'\(处理完成)transaction交易明细_已发放-推迟订单{transaction_date}.xlsx'
merged_df_3.to_excel(output_path, index=False)
print(f'处理完成，output_path：{output_path}')
