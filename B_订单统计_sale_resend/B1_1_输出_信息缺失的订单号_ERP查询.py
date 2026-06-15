import re
import importlib.util
from pathlib import Path

# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

import pandas as pd
from A_报表.Z_method.style import Color
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-1)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)

print(f"{main_file_path}")

# 一、REAL-FB  没有站点（国家）的
# 筛选条件  店铺英文名 == FB_REAL  且 站点 == UNKNOW
real_df = main_file_df[(main_file_df['店铺英文名'] == 'FB_REAL') & (main_file_df['站点'] == 'UNKNOW')]
# 获取订单号列表
REAL_UNKNOW_order_list = real_df['订单号'].tolist()
# 判断是否有符合条件的订单号
if REAL_UNKNOW_order_list:
    print(f'需要手动查询，FB_REAL，站点为"UNKNOW"的订单号：{len(REAL_UNKNOW_order_list)}个')
    REAL_UNKNOW_order_str = ' '.join(REAL_UNKNOW_order_list)  # 改成批量查询的格式
    print(f'REAL_UNKNOW_order_str：（复制下面这行，直接 Ctrl + C）{Color.YELLOW}\n{REAL_UNKNOW_order_str}{Color.RESET}')
    print("存储文件夹_path：")
    print(r"\\Betohow\数据报表\RPA\报表-无站点-订单查询\REAL-FB")
else:
    print('FB_REAL，没有 站点为"UNKNOW"的订单号！！！')

print('-' * 100)

# 二、LM-BC的重发订单
# 筛选条件  店铺英文名 == LM_BC_FR  且 订单类型 == 重发订单
lm_bc_df = main_file_df[(main_file_df['店铺英文名'] == 'LM_BC_FR') & (main_file_df['订单类型'] == '重发订单')]
# 获取订单号列表  且 去掉订单号 的后缀
LM_BC_resend_order_list = [re.sub(r'-[1-5]', '', i) for i in lm_bc_df['订单号']]
# 判断是否有符合条件的订单号
if LM_BC_resend_order_list:
    print(f'需要手动查询，LM_BC_FR，订单类型 == "重发订单"的订单号：{len(LM_BC_resend_order_list)}个')
    LM_BC_resend_order_str = ' '.join(LM_BC_resend_order_list)  # 改成批量查询的格式
    print(f'LM_BC_resend_order_str：（复制下面这行，直接 Ctrl + C）{Color.YELLOW}\n{LM_BC_resend_order_str}{Color.RESET}')
    print("存储文件夹_path：")
    print(r"\\Betohow\数据报表\RPA\报表-无站点-订单查询\LM-BC-重发")
else:
    print('LM_BC_FR，没有 "重发订单"！！！')

print('-' * 100)

# 三、LM-RP的重发订单
# 筛选条件  店铺英文名 == LM_RP_FR  且 订单类型 == 重发订单
lm_rp_df = main_file_df[(main_file_df['店铺英文名'] == 'LM_RP_FR') & (main_file_df['订单类型'] == '重发订单')]
# 获取订单号列表  且 去掉订单号 的后缀
LM_RP_resend_order_list = [re.sub(r'-[1-5]', '', i) for i in lm_rp_df['订单号']]
# 判断是否有符合条件的订单号
if LM_RP_resend_order_list:
    print(f'需要手动查询，LM_RP_FR，订单类型 == "重发订单"的订单号：{len(LM_RP_resend_order_list)}个')
    LM_RP_resend_order_str = ' '.join(LM_RP_resend_order_list)  # 改成批量查询的格式
    print(f'{Color.YELLOW}LM_RP_resend_order_str：（复制下面这行，直接 Ctrl + C）\n{LM_RP_resend_order_str}{Color.RESET}')
    print("存储文件夹_path：")
    print(r"\\Betohow\数据报表\RPA\报表-无站点-订单查询\LM-RP-重发")
else:
    print('LM_RP_FR，没有 "重发订单"！！！')
