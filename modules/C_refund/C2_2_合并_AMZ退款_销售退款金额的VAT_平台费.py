import pandas as pd
import importlib.util
from pathlib import Path
# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from config.A0_paths import SELLERSKU_PROFIT_FILE_NAME, SELLERSKU_PROFIT_REPORT_DIR

main_file_path = fr"{SELLERSKU_PROFIT_REPORT_DIR}\(退款-1){SELLERSKU_PROFIT_FILE_NAME}"
main_file_df = pd.read_excel(main_file_path)
main_file_df_1 = main_file_df.groupby('SKU-站点识别码').agg({
    'SKU': 'first',  # 保留每组的第一行数据
    '映射站点': 'first',  # 保留每组的第一行数据
    '映射平台': 'first',  # 保留每组的第一行数据
    'SKU-平台识别码': 'first',  # 保留每组的第一行数据
    '退款量': 'sum',
    '退款额': 'sum',
    '销售退款金额VAT-amazon': 'sum',
    '销售退款金额的佣金': 'sum',
}).reset_index()

# 重命名
main_file_df_1 = main_file_df_1.rename(columns={'映射站点': '站点'})
main_file_df_1 = main_file_df_1.rename(columns={'映射平台': '平台'})
main_file_df_1 = main_file_df_1.rename(columns={'退款量': '退款数量'})

output_path = fr"{SELLERSKU_PROFIT_REPORT_DIR}\(处理完成-退款){SELLERSKU_PROFIT_FILE_NAME}"
main_file_df_1.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
