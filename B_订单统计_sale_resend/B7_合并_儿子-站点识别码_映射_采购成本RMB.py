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

from A_报表.Z_method.style import Color
from A_报表.A0_设置_时间段.A0_set_date import RATE_SHIP_FEE, SKU_NW_DISCOUNT
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import BTH_ALL_SKU_DETAIL_PATH, DESKTOP_ROOT
from A_报表.Z_method.sku_映射 import sku_mappings


# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-5-1)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)

# 去除 整张表 的前后空格
for col in main_file_df.columns:
    main_file_df[col] = main_file_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

# 重命名
main_file_df = main_file_df.rename(columns={'站点': '原-站点'})
main_file_df = main_file_df.rename(columns={'映射站点': '站点'})
main_file_df = main_file_df.rename(columns={'平台': '原-平台'})
main_file_df = main_file_df.rename(columns={'映射平台': '平台'})

# 将列 "订单类型" 中的 "销售订单" 替换为 "sale"、"线下订单" 替换为 "sale"
main_file_df["订单类型"] = main_file_df["订单类型"].replace("销售订单", "sale")
main_file_df["订单类型"] = main_file_df["订单类型"].replace("线下订单", "sale")
# 将列 "订单类型" 中的 "重发订单" 替换为 "resend"、"FBA换货单" 替换为 "resend"
main_file_df["订单类型"] = main_file_df["订单类型"].replace("重发订单", "resend")
main_file_df["订单类型"] = main_file_df["订单类型"].replace("FBA换货单", "resend")

# 获取“儿子-站点识别码”的所有唯一值（自动去重）
col1_conditions = main_file_df["儿子-站点识别码"].dropna().unique().tolist()  # 条件

# 当仓库属性 == 第三方 时，派送费需乘以费率
print(f"{Color.GREEN}\n 第三方仓库尾程派送系数：{RATE_SHIP_FEE}{Color.RESET}")
third_party_mask = main_file_df['仓库属性'] == '第三方'
main_file_df.loc[third_party_mask, '派送运费'] = main_file_df.loc[third_party_mask, '派送运费'] * RATE_SHIP_FEE

output_path = main_file_path.replace('已完成-5-1', '已完成-6')
writer = pd.ExcelWriter(output_path, engine='openpyxl')

# 创建空DataFrame存储所有结果
final_results_df = pd.DataFrame()
# 主循环：按“儿子-站点识别码”列的所有值循环
for status in col1_conditions:
    # 第一次筛选：当前儿子-站点识别码的数据
    temp_main_file_df = main_file_df[main_file_df["儿子-站点识别码"] == status]
    # 初始化一个字典，用于存储每种订单类型的销量总和
    order_type_counts = {category: 0 for category in main_file_df["订单类型"].unique()}
    # 初始化总和变量
    sum_1 = 0
    sum_2 = 0
    sum_3 = 0
    sum_4 = 0
    sum_5 = 0
    sum_6 = 0
    # 第二次循环：按“订单类型”列的唯一值循环
    for category in main_file_df["订单类型"].unique():
        # 第二次筛选：当前类别的数据
        filtered_main_file_df = temp_main_file_df[temp_main_file_df["订单类型"] == category]
        # 对筛选结果的某列（如“订单总金额”）求和
        sum_1 += filtered_main_file_df["订单总金额"].sum()
        sum_2 += filtered_main_file_df["头程运费"].sum()
        sum_3 += filtered_main_file_df["头程税费"].sum()
        sum_4 += filtered_main_file_df["派送运费"].sum()
        sum_5 += filtered_main_file_df["平台销售额VAT-amazon"].sum()
        # 记录当前订单类型的销量总和
        order_type_counts[category] = filtered_main_file_df["仓库SKU销量"].sum()
    # 获取第一行的某些列数据（如“ID”和“Name”）
    first_row_data = temp_main_file_df.iloc[0][["平台", "站点", "儿子-平台识别码", "SKU", "儿子-站点识别码"]]
    # 创建当前条件的临时结果DataFrame
    temp_result = pd.DataFrame({
        "SKU": [first_row_data["SKU"]],
        "站点": [first_row_data["站点"]],
        "平台": [first_row_data["平台"]],
        "儿子-站点识别码": [first_row_data["儿子-站点识别码"]],
        "儿子-平台识别码": [first_row_data["儿子-平台识别码"]],
        "平台销售额": [sum_1],
        "头程": [sum_2],
        "关税": [sum_3],
        "派送费": [sum_4],
        "平台销售额VAT-amazon": [sum_5],
    })
    # 添加每种订单销量类型的总和列
    for category, count in order_type_counts.items():
        temp_result[f"{category}数量总和"] = [count]
    # 将临时结果追加到最终结果
    final_results_df = pd.concat([final_results_df, temp_result], ignore_index=True)
final_results_df = final_results_df.rename(columns={'sale数量总和': '销量'})
final_results_df = final_results_df.rename(columns={'resend数量总和': '重发数量'})
# 如果 没有 重发数量 列，则新增列：重发数量
if '重发数量' not in final_results_df.columns:
    final_results_df['重发数量'] = 0

final_results_df['平台销售额'] = np.round(final_results_df['平台销售额'], 2)
final_results_df['头程'] = np.round(final_results_df['头程'], 2)
final_results_df['关税'] = np.round(final_results_df['关税'], 2)
final_results_df['派送费'] = np.round(final_results_df['派送费'], 2)


# 计算 sale、resend的采购成本
# 获取 采购价
product_map_sku_path = BTH_ALL_SKU_DETAIL_PATH
final_results_df_1 = sku_mappings(
    main_df=final_results_df,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="SKU",
    map_new_sku="原始采购价",
    map_sku_sheet='基础数据维护'

)

# 将字符串转换为数值类型（空字符串或非数字字符串会转换为 NaN）
final_results_df_1['映射原始采购价'] = pd.to_numeric(final_results_df_1['映射原始采购价'], errors='coerce')

final_results_df_1['订单采购成本'] = np.round(final_results_df_1['销量'] * final_results_df_1['映射原始采购价'], 2)
final_results_df_1['重发采购成本'] = np.round(final_results_df_1['重发数量'] * final_results_df_1['映射原始采购价'], 2)

# 2026-06-05 调整：SKU以 -NW 结尾的，订单采购成本&重发采购成本 打折
nw_suffix_mask = final_results_df_1['SKU'].astype(str).str.endswith('-NW')
final_results_df_1.loc[nw_suffix_mask, '订单采购成本'] = final_results_df_1.loc[nw_suffix_mask, '订单采购成本'] * SKU_NW_DISCOUNT
final_results_df_1.loc[nw_suffix_mask, '重发采购成本'] = final_results_df_1.loc[nw_suffix_mask, '重发采购成本'] * SKU_NW_DISCOUNT
print(f"{Color.YELLOW}2026-06-05 调整：SKU以 -NW 结尾的，订单采购成本&重发采购成本 打折：{SKU_NW_DISCOUNT}{Color.RESET} \n")

# CD平台的重发采购替换为 0
final_results_df_1.loc[final_results_df_1['平台'] == 'CD', '重发采购成本'] = 0

# 保存Excel文件将所有结果写入
final_results_df_1.to_excel(writer, index=False)
writer.close()
print(f'处理完成，文件另存为：{output_path}')
print(f'{Color.YELLOW}检查：映射原始采购价，是否有空的，空的去查——是否是智慧谷的，25开头的 都是智慧谷的分销，智慧谷的采购成本为0{Color.RESET}')
