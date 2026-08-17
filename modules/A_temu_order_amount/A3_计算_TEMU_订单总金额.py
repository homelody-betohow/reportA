import warnings
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

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")
# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\订单统计-{shared_date}.xlsx"
main_df = pd.read_excel(main_file_path, skiprows=4)  # 跳过前4行
# # ====================================================================================#
# # 从2026-07-01 从 付款时间维度调整到发货时间维度，6月的数据需要过滤 07-01后发货的订单       #
# # # main_df = main_df[main_df['发货时间'] > '2026-07-01']                              #
# # ====================================================================================#


# 过滤掉 “订单销售状态” 为 问题件 的行
main_df = main_df[main_df['订单销售状态'] != '问题件']
# ============================================================================================

# 过滤掉 “订单销售状态” 为 “冻结中” 的行
main_df = main_df[main_df['订单销售状态'] != '冻结中']

# 过滤指定订单，不统计利润
filter_order_nos = [
    'WEC0382608170069'
] 
main_df = main_df[~main_df['订单号'].isin(filter_order_nos)]

# 替换列名
main_df.rename(columns={'仓库sku销量': '仓库SKU销量', '仓库sku': '仓库SKU'}, inplace=True)
# 筛选不等于'semitemu'的行
no_temu_df = main_df[main_df['平台'] != 'semitemu']
# 保留指定列  保证 拼接前  列名一致
no_temu_df = no_temu_df[
    ['平台', '店铺英文名', '站点', '订单类型', '参考号', '订单号', '平台sku', '仓库属性', '仓库',  '运输方式','国家','邮编',
    '跟踪单号','仓库SKU','仓库SKU销量',  '订单总金额','币种', '平台运费', 'fba费用', '头程运费', '头程税费', '派送运费','发货时间']
]

# TODO 文件路径！！！！！
temu_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\只有TEMU(已完成-1)订单统计-{shared_date}.xlsx"
temu_df = pd.read_excel(temu_path)
# # ====================================================================================#
# # 从2026-07-01 从 付款时间维度调整到发货时间维度，6月的数据需要过滤 07-01后发货的订单       #
# # temu_df = temu_df[temu_df['发货时间'] > '2026-07-01']                                #
# # ====================================================================================#

# 替换列名
temu_df.rename(columns={'仓库sku销量': '仓库SKU销量', '仓库sku': '仓库SKU'}, inplace=True)

temu_df['订单总金额'] = temu_df['映射产品单价（EUR）'].astype(float) * temu_df['仓库SKU销量'] + temu_df[
    '映射运费回款（EUR）']

# 保留指定列  保证 拼接前  列名一致
temu_df = temu_df[
    ['平台', '店铺英文名', '站点', '订单类型', '参考号', '订单号', '平台sku','仓库属性', '仓库', '运输方式','国家','邮编',
    '跟踪单号','仓库SKU','仓库SKU销量', '订单总金额','币种', '平台运费', 'fba费用', '头程运费', '头程税费', '派送运费','发货时间']
]

# 按行拼接（追加新行）
main_df_1 = pd.concat([no_temu_df, temu_df], ignore_index=True)

# 筛选  “店铺英文名”不包含 ECO、Biancca、yiqianshangmao_DE 的行
main_df_2 = main_df_1[
        ~main_df_1['店铺英文名'].str.contains('ECO|Biancca|yiqianshangmao_DE', na=False)].copy()


# 保存修改后的
output_path = main_file_path.rsplit('\\', 1)[0] + '\\(已完成-1)' + main_file_path.rsplit('\\', 1)[1]
main_df_2.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
"""
合并上去的TEMU的订单总金额是str，合并SKU-站点识别码的时候，会自动变成 数值
"""
