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
from A_报表.Z_method.style import Color
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！
main_file_path = fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\二次上架\鸿羽仓-二次上架明细-{shared_date}.xls'
main_df = pd.read_excel(main_file_path)

# 打印——需要手动查询的“参考号”、“订单参考号”
# 获取列“订单参考号”的数据list，且不要空值
order_list_ = main_df['订单参考号'].dropna().tolist()
# 在筛选列’映射账号‘为空的基础上，再筛选列 '订单参考号' 为空的行
filtered_df_2 = main_df[main_df['订单参考号'].isnull()]
# 获取列“参考号”的数据list，且不要空值
refer_list_ = filtered_df_2['参考号'].dropna().tolist()
refer_list = [i for i in refer_list_ if '900008-' in i]  # 得到所有的参考号，去“自发货”查
refer_str = ' '.join(refer_list)  # 改成批量查询的格式
print(
    f'需要手动查询——自发货！参考号_len：{len(refer_list)}，参考号_str：（复制下面这行，直接 Ctrl + C）{Color.YELLOW}\n{refer_str}{Color.RESET}')
order_list = order_list_ + [i for i in refer_list_ if i not in refer_list]  # 得到所有的订单参考号，去“订单管理”查
order_str = ' '.join(order_list)  # 改成批量查询的格式
print(
    f'\n需要手动查询——订单管理！订单参考号_len：{len(order_list)}，订单参考号_str：（复制下面这行，直接 Ctrl + C）{Color.YELLOW}\n{order_str}{Color.RESET}')
# TODO 批量去查（RPA）
