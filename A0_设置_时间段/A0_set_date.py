from calendar import monthrange
from datetime import datetime, timedelta
import importlib.util
from pathlib import Path
# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.Z_method.time import get_month_range

# 汇率 别的币种 转 EUR
RMB_di_EUR = 7.3  # 人民币转欧元(除法)
USD_to_EUR = 0.858  # 美元
kc_to_EUR = 0.04133  # 捷克克朗
zl_to_EUR = 0.237  # 波兰兹罗提
Ft_to_EUR = 0.002611  # 匈牙利福林
CAD_to_EUR = 0.6179  # 加拿大元
kr_to_EUR = 0.0934  # 瑞典克朗
Lei_to_EUR = 0.196  # 罗马尼亚列伊

RATE_SHIP_FEE = 1.05  # 运费费率
SKU_NW_DISCOUNT = 0.8  # NW尾缀SKU 折扣

# 切换报表类型：'日报' | '月报'
# folder_name = '月报'
folder_name = '日报'

_VALID_FOLDER_NAMES = ('日报', '月报')
# 日报 获取多少天前的数据
_X_DAY = 3

def _format_md(month, day):
    return f"{month}.{day}"


def _last_day_of_month(year, month):
    return monthrange(year, month)[1]


def _period_for_daily(today):
    """日报：统计日为今天-3天，区间为当月1号至该日。"""
    anchor = today - timedelta(days=_X_DAY)
    shared_date = f"{_format_md(anchor.month, 1)}-{_format_md(anchor.month, anchor.day)}"
    ku_cun_date = f"{anchor.year}.{anchor.month}.{anchor.day}"
    return shared_date, anchor, ku_cun_date


def _period_for_monthly(today):
    """月报：上一个自然月 1 号至月末。"""
    first_this_month = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    last_day_prev = first_this_month - timedelta(days=1)
    y, m = last_day_prev.year, last_day_prev.month
    shared_date = (
        f"{_format_md(m, 1)}-{_format_md(m, _last_day_of_month(y, m))}"
    )
    ku_cun_date = f"{y}.{m}.{last_day_prev.day}"
    return shared_date, last_day_prev, ku_cun_date


def _resolve_period(folder):
    today = datetime.now()
    if folder == '日报':
        return _period_for_daily(today)
    if folder == '月报':
        return _period_for_monthly(today)
    raise ValueError(f"folder_name 只能是 {_VALID_FOLDER_NAMES}，当前为: {folder!r}")


shared_date, report_date, ku_cun_date = _resolve_period(folder_name)

# 获取当前年月
current_date = datetime.now()
current_year_month = f"{current_date.year}.{current_date.month}"

# fba_date = '3.1-3.31'
_start_m, _ = map(int, shared_date.split('-')[0].split('.'))



# 测评月份-定位Sheet页
# test_file_sheet_name = '2026.5'
test_file_sheet_name = f"{report_date.year}.{report_date.month}"
# 测评开始时间、结束时间
# test_start_date = '2026-5-1'
# test_end_date = '2026-5-29'
test_start_date = f"{report_date.year}-{report_date.month}-1"
test_end_date = f"{report_date.year}-{report_date.month}-{report_date.day}"

# TODO 目标拆解表，用于 映射 Amazon的销售负责人（手动改，一月改一次）
# 文件目录见 A0_paths.MONTH_GOAL_DIR（当前：F:\月目标拆解及跟进）
month_goal_excel = '2026-06月目标拆解及跟进.xlsx'

# 获取 transaction文件命名的日期
today = datetime.today()
if folder_name == '日报':
    fba_date = get_month_range(2, datetime(report_date.year, _start_m, 1))
    m, d = map(int, shared_date.split('-')[-1].split('.'))
    # 加天数  +3天
    formatted_day = f'{(datetime(today.year, m, d) + timedelta(days=3)).strftime("%#m.%#d")}'
else:
    fba_date = get_month_range(1, datetime(report_date.year, _start_m, 1))
    # 月报 每月5号
    formatted_day = f'{today.month}.5'
    
transaction_date = shared_date.split('-')[0] + '-' + formatted_day

print(f"Run-{folder_name}: {shared_date}")
# print(f"测评时间： {test_file_sheet_name}")
# print(f"transaction_date: {transaction_date}")
# print(f"fba_date: {fba_date}")