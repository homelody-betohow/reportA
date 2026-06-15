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
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-17)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)

# 如果供应商是“易速”或“智慧谷”，则将“产品状态”、“二级分类”、“三级分类”替换为 “分销”
main_file_df.loc[main_file_df['供应商'].isin(['易速', '智慧谷']), ['产品状态', '二级分类', '三级分类']] = '分销'
# 如果供应商是“智慧谷”，则将“采购成本”、“订单采购成本”、“重发采购成本”、“二次上架采购成本”替换为 0
main_file_df.loc[
    main_file_df['供应商'] == '智慧谷', ['采购成本', '订单采购成本', '重发采购成本', '二次上架采购成本']] = 0
# 如果产品状态是“分销”，则将“运营模式”替换为 “自运营”；“二级分类”、“三级分类”替换为 “分销”
main_file_df.loc[main_file_df['产品状态'] == '分销', ['运营模式', '二级分类', '三级分类']] = ['自运营', '分销', '分销']
# 如果产品状态是“分销”，则将“头程”、“关税”、“派送费”替换为 0
main_file_df.loc[main_file_df['产品状态'] == '分销', ['头程', '关税', '派送费']] = 0

# 分摊-仓租
# 1. 计算需要分摊的总费用
total_cost = main_file_df['所有仓库-无平台-需要分摊的费用'].sum()
# 2. 筛选参与分摊的行      选择 产品状态 不等于 分销  的行
participating_rows = main_file_df[main_file_df['产品状态'] != '分销']
# 3. 计算分摊值
num_participating = len(participating_rows)  # '产品状态' != '分销'的，总的行数
if num_participating > 0:
    allocation_per_row = total_cost / num_participating
else:
    allocation_per_row = 0.0
# 4. 创建新列“仓租分摊”
main_file_df['仓租分摊'] = 0.0  # 初始化新列为0
main_file_df.loc[main_file_df['产品状态'] != '分销', '仓租分摊'] = allocation_per_row
# 重命名
main_file_df = main_file_df.rename(columns={'海外仓仓租费': '原-海外仓仓租费'})
main_file_df['海外仓仓租费'] = np.round(main_file_df['原-海外仓仓租费'] + main_file_df['仓租分摊'], 2)

# 空值的地方——补 0   使 仓租合计  可以正常合计
main_file_df = main_file_df.fillna(0)

main_file_df['仓租合计'] = np.round(main_file_df['FBA仓租费'] + main_file_df['海外仓仓租费'], 2)

# 保存修改后的文件
output_path = main_file_path.replace('已完成-17', '已完成-18')
main_file_df.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
