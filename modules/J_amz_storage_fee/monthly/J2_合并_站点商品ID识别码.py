import pandas as pd
import importlib.util
from pathlib import Path
# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.sku_mapping import sku_mappings
from config.A0_set_date import shared_date, folder_name, fba_date
from config.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！   上月的 利润报表
Amazon_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\FBA仓租\(已完成-1)FBA仓租明细{fba_date}.xlsx"
Amazon_file_df = pd.read_excel(Amazon_file_path)

# 重命名
Amazon_file_df = Amazon_file_df.rename(columns={'映射站点': '站点'})
Amazon_file_df = Amazon_file_df.rename(columns={'映射平台': '平台'})
# 删除列 'FBA仓租费' 中值为 0 的行
Amazon_file_df = Amazon_file_df[Amazon_file_df['FBA仓租费'] != 0]

# 按照 '站点商品ID识别码' 列进行分组，并对 '数量' 和 '退件费用(EUR)' 列进行汇总
result_df = Amazon_file_df.groupby('站点商品ID识别码').agg({
    'SKU': 'first',  # 保留每组的第一行数据
    '商品ID': 'first',  # 保留每组的第一行数据
    '站点': 'first',  # 保留每组的第一行数据
    '平台': 'first',  # 保留每组的第一行数据
    '平台商品ID识别码': 'first',  # 保留每组的第一行数据
    'FBA仓租费': 'sum'  # 汇总 数量
}).reset_index()

product_map_sku_path = r"\\Betohow\数据报表\数据库\产品信息库2025.xlsx"  # 改成对应的映射表
#  商品ID 去映射 产品信息库 的 第一个 产品编码（SKU）
result_df_1 = sku_mappings(
    main_df=result_df,
    main_sku='商品ID',
    map_sku_path=product_map_sku_path,
    map_old_sku="商品ID",
    map_new_sku="产品编码",
    map_sku_sheet='产品信息表'
)
result_df_1 = result_df_1.rename(columns={'SKU': '原-SKU'})
result_df_1 = result_df_1.rename(columns={'映射产品编码': 'SKU'})

result_df_1 = result_df_1[['SKU','商品ID', '站点', '平台', '站点商品ID识别码', '平台商品ID识别码', 'FBA仓租费']]

# 保存修改
output_path = Amazon_file_path.replace('已完成-1', '处理完成')
result_df_1.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
