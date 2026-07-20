import openpyxl
import pandas as pd
import importlib.util
from pathlib import Path
# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.Z_method.platform_shop import map_region_to_platform
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT
from A_报表.A0_设置_时间段.A0_set_date import RMB_di_EUR

# TODO 文件路径！！！
test_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\测评表\(已完成-3)测评表.xlsx"
test_file_df = pd.read_excel(test_file_path)

# 计算总的 测评费
test_file_df['测评费'] = 0.0

# 筛选条件：退款类型是 "空包退订单金额"
mask_1 = test_file_df['退款类型'] == '空包退订单金额'
test_file_df.loc[mask_1, '测评费'] = test_file_df['平台费'] + test_file_df['销售税'] + test_file_df['提现费']
#
# 筛选条件：退款类型是 "测评退订单金额"
mask_2 = test_file_df['退款类型'] == '测评退订单金额'
test_file_df.loc[mask_2, '测评费'] = test_file_df['报表金额'] * 1.05 + 0.34
#
# 筛选条件：退款类型是 "测评退订单金额70%"
mask_3 = test_file_df['退款类型'] == '测评退订单金额70%'
test_file_df.loc[mask_3, '测评费'] = test_file_df['报表金额'] * 1.05 * 0.7 + 0.34
#
# 筛选条件：退款类型是 "佣金"
mask_4 = test_file_df['退款类型'] == '佣金'
test_file_df.loc[mask_4, '测评费'] = test_file_df['报表金额']
#
# 筛选条件：退款类型是 "好评返现"
mask_5 = test_file_df['退款类型'] == '好评返现'
test_file_df.loc[mask_5, '测评费'] = test_file_df['报表金额']

# 保存目标列
test_file_df = test_file_df[
    ['订单号', '数量', '站点', '退款日期', '退款类型', '订单金额', '原币种', '支付方式', 'SKU', 'SKU-站点识别码',
     '报表币种', '报表金额', '头程', '关税', '采购价', '运费', '提现费', '销售税', '平台费', '测评费']]

# 映射 平台（数据源：platform_shop）
test_file_df_5 = map_region_to_platform(test_file_df, site_col='站点')
# 在 SKU-站点识别码 后插入 SKU-平台识别码
new_column_name = "SKU-平台识别码"  # 新列名
new_column_data = test_file_df_5["映射平台"] + test_file_df_5["SKU"]  # 新列数据
target_column = "SKU-站点识别码"  # 目标列名（在其后插入）
insert_position = test_file_df_5.columns.get_loc(target_column) + 1  # 计算插入位置
test_file_df_5.insert(insert_position, new_column_name, new_column_data)  # 插入新列
test_file_df_5 = test_file_df_5.rename(columns={'映射平台': '平台'})

# # 四舍五入，强制 2 位小数，保存文件
output_path = test_file_path.replace('(已完成-3)', '(已完成-4)')
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    test_file_df_5.to_excel(writer, index=False, float_format="%.2f")  # 四舍五入，强制 2 位小数
    worksheet = writer.sheets['Sheet1']  # 默认 Sheet 名
    # 强制 2 位小数 的列名
    set_2_list = ['头程', '关税', '采购价', '运费', '提现费', '销售税', '平台费', '测评费']
    for idx, col in enumerate(test_file_df_5.columns):
        if col in set_2_list:
            col_letter = openpyxl.utils.get_column_letter(idx + 1)  # +1 因为 Excel 从 1 开始
            for cell in worksheet[col_letter]:
                cell.number_format = '0.00'  # 设置 Excel 单元格格式
print(f'处理完成，文件另存为：{output_path}')
