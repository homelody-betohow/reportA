# 网络共享盘路径（版本更新时只改这里）
BTH_ALL_SKU_DETAIL_PATH = r"\\Betohow\数据报表\数据库\BTH全部SKU明细-v2026.06.29.xlsx"

# 桌面根目录（换电脑或用户名时只改这里）
DESKTOP_ROOT = r"C:\Users\BTH-windows\Desktop"

# 月目标拆解表目录（换盘符或文件夹时只改这里）
MONTH_GOAL_DIR = r"F:\月目标拆解及跟进"

import importlib.util
from pathlib import Path
# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.A0_设置_时间段.A0_set_date import folder_name, shared_date, month_goal_excel

# 月目标拆解表完整路径（文件名在 A0_set_date.month_goal_excel 中配置）
MONTH_GOAL_EXCEL_PATH = fr"{MONTH_GOAL_DIR}\{month_goal_excel}"

# 当前报表周期目录（日报/月报 + 时间段）
REPORT_PERIOD_DIR = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}"

# SellerSku 始利润报表路径
SELLERSKU_PROFIT_REPORT_DIR = (
    fr"{REPORT_PERIOD_DIR}\SellerSku利润报表"
)
# SellerSku 原始利润报表文件
SELLERSKU_PROFIT_FILE_NAME = (
    fr"SellerSku利润报表{shared_date}.xlsx"
)