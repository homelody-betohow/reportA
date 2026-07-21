import numpy as np
import pandas as pd
from pathlib import Path
import importlib.util
# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from config.A0_set_date import shared_date, folder_name
from config.A0_set_date import RMB_di_EUR
from config.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！
# 判断是否有 测评费，自动选择文件路径
# 有  测评费  的路径
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-11)订单统计-{shared_date}.xlsx"
if not Path(main_file_path).is_file():
    print(f"文件不存在，无  测评费！！！")
    # 有  秒杀费  的路径
    main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-10)订单统计-{shared_date}.xlsx"
    if not Path(main_file_path).is_file():
        print(f"文件不存在，无  秒杀费！！！")
        main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-9)订单统计-{shared_date}.xlsx"

main_file_df = pd.read_excel(main_file_path)
# 检查 下面这些 列名 是否存在，不存在，则：新增列，数据为 0
cols_to_check = ['秒杀费', '测评费']
for col in cols_to_check:
    if col not in main_file_df.columns:
        main_file_df[col] = 0
        print(f"新增一列:{col}，数据全是 0")

# TODO 文件路径！！！
guang_gao_path = fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\二次上架\(处理完成)鸿羽仓-二次上架明细-{shared_date}.xlsx'
guang_gao_df = pd.read_excel(guang_gao_path)

# 以SKU-站点识别码为键进行合并，选择左连接（left join），这样可以确保表1的所有数据都被保留
result_df = pd.merge(main_file_df,
                     guang_gao_df[['SKU-站点识别码', '二次上架数量', '二次上架金额', '二次上架采购成本', '返还采购成本']],
                     on='SKU-站点识别码', how='left')

# 找出表2中在表1中不存在的行
missing_rows = guang_gao_df[~guang_gao_df['SKU-站点识别码'].isin(main_file_df['SKU-站点识别码'])]

# 将这些缺失的行添加到结果中
result_df = pd.concat([result_df, missing_rows], ignore_index=True)

# 空值的地方——补 0（分销列除外）
if '分销' not in result_df.columns:
    result_df['分销'] = '否'
else:
    result_df['分销'] = result_df['分销'].replace({0: '否', '0': '否'}).fillna('否')
_fill_cols = [c for c in result_df.columns if c != '分销']
result_df[_fill_cols] = result_df[_fill_cols].fillna(0)

# 确保所有期望的列都存在
expected_columns = list(main_file_df.columns) + ['二次上架数量', '二次上架金额', '二次上架采购成本', '返还采购成本']
for col in expected_columns:
    if col not in result_df.columns:
        result_df[col] = None  # 如果列不存在，添加该列并填充为 None

# 表1中没有和表2相同的SKU-站点识别码，将新增的两在对应行单元格填充为0
result_df = result_df.fillna({
    '二次上架数量': 0,
    '二次上架金额': 0,
    '二次上架采购成本': 0,
    '返还采购成本': 0
})
# 重新排序，确保列的顺序符合要求
result_df = result_df[expected_columns]
# 空值的地方——补 0
result_df = result_df.fillna(0)

# RMB转EUR
result_df['采购成本'] = np.round((result_df['订单采购成本'] + result_df['重发采购成本'] - result_df['返还采购成本']) / RMB_di_EUR, 2)


# 保存结果到新的 Excel 文件
output_path = main_file_path.replace('已完成-11', '已完成-12')
# 没有测评费的话， 则是：10 直接跳到 12
output_path = output_path.replace('已完成-10', '已完成-12')
# 没有测评费、秒杀费的话，则是：9 直接跳到 12
output_path = output_path.replace('已完成-9', '已完成-12')
result_df.to_excel(output_path, index=False)
print(f"结果已保存到文件：{output_path}")
