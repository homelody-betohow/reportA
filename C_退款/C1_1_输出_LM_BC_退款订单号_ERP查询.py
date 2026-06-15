import warnings
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
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")
# TODO 文件路径！！！
RMA_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\RMA\RMA-{shared_date}.xlsx"
RMA_file_df = pd.read_excel(RMA_file_path, skiprows=2)  # 跳过前2行 如果文件读取失败的话，则手动删掉多余的列，只保留想要的列
# 去除 整张表 的前后空格
for col in RMA_file_df.columns:
    RMA_file_df[col] = RMA_file_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

# LM-BC的退款订单
# 筛选条件  店铺英文名 == LM_BC_FR
lm_bc_df = RMA_file_df[(RMA_file_df['店铺英文名'] == 'LM_BC_FR')]
# 获取 退款原订单号 列表
LM_BC_refund_order_list = lm_bc_df['退款原订单号'].tolist()
# 判断是否有符合条件的 退款原订单号
if LM_BC_refund_order_list:
    print(f'需要手动查询，LM_BC_FR，退款原订单号：{len(LM_BC_refund_order_list)}个')
    LM_BC_refund_order_str = ' '.join(LM_BC_refund_order_list)  # 改成批量查询的格式
    print(f'LM_BC_refund_order_str：（复制下面这行，直接 Ctrl + C）{Color.YELLOW}\n{LM_BC_refund_order_str}{Color.RESET}')
    print("存储文件夹_path：")
    print(r"\\Betohow\数据报表\RPA\报表-无站点-订单查询\LM-BC-退款")
else:
    print('LM_BC_FR，没有 "退款订单"！！！')

print('-' * 100)

# LM_RP的退款订单
# 筛选条件  店铺英文名 == LM_RP_FR
lm_rp_df = RMA_file_df[(RMA_file_df['店铺英文名'] == 'LM_RP_FR')]
# 获取 退款原订单号 列表
LM_RP_refund_order_list = lm_rp_df['退款原订单号'].tolist()
# 判断是否有符合条件的 退款原订单号
if LM_RP_refund_order_list:
    print(f'需要手动查询，LM_RP_FR，退款原订单号：{len(LM_RP_refund_order_list)}个')
    LM_RP_refund_order_str = ' '.join(LM_RP_refund_order_list)  # 改成批量查询的格式
    print(f'LM_RP_refund_order_str：（复制下面这行，直接 Ctrl + C）{Color.YELLOW}\n{LM_RP_refund_order_str}{Color.RESET}')
    print("存储文件夹_path：")
    print(r"\\Betohow\数据报表\RPA\报表-无站点-订单查询\LM-RP-退款")
else:
    print('LM_RP_FR，没有 "退款订单"！！！')

