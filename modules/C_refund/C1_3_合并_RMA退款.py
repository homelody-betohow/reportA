import pandas as pd
import importlib.util
from pathlib import Path
# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from config.A0_set_date import shared_date, folder_name
from config.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！
RMA_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\RMA\(已完成-1-1)RMA-{shared_date}.xlsx"
RMA_file_df = pd.read_excel(RMA_file_path)

# 重命名
RMA_file_df = RMA_file_df.rename(columns={'站点': '原-站点'})
RMA_file_df = RMA_file_df.rename(columns={'映射站点': '站点'})
RMA_file_df = RMA_file_df.rename(columns={'平台': '原-平台'})
RMA_file_df = RMA_file_df.rename(columns={'映射平台': '平台'})
RMA_file_df = RMA_file_df.rename(columns={'RMA产品数量': '退款数量'})
RMA_file_df = RMA_file_df.rename(columns={'退款金额': '退款额'})
# 为了和 利润报表退款 的表头 一致
RMA_file_df['销售退款金额VAT-amazon'] = 0
RMA_file_df['销售退款金额的佣金'] = 0
# 按照 'SKU-站点识别码' 列进行分组，汇总
RMA_file_df_1 = RMA_file_df.groupby('SKU-站点识别码').agg({
    'SKU': 'first',  # 保留每组的第一行数据
    '站点': 'first',  # 保留每组的第一行数据
    '平台': 'first',  # 保留每组的第一行数据
    'SKU-平台识别码': 'first',  # 保留每组的第一行数据
    '分销': lambda x: '是' if (x == '是').any() else '否',
    '退款数量': 'sum',
    '退款额': 'sum',
    '销售退款金额VAT-amazon': 'sum',
    '销售退款金额的佣金': 'sum',
}).reset_index()

# 筛选出“平台”列中不包含“AMAZON”的行
df_no_amazon = RMA_file_df_1[~RMA_file_df_1['平台'].str.contains('AMAZON', na=False)]

# 保存结果
output_path = RMA_file_path.replace('已完成-1-1', '处理完成-无Amazon')
df_no_amazon.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
