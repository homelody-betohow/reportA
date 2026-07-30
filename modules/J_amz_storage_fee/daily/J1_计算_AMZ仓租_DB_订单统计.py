import importlib.util
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import pymysql.cursors

# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from config.A0_paths import DESKTOP_ROOT
from config.A0_set_date import folder_name, report_date, shared_date
from database.db_connection import get_db_manager

RENT_TABLE = "amz_warehouse_rent_snapshot"


def resolve_snapshot_month() -> str:
    """与 A0_set_date.fba_date / V3 快照一致：日报往前 2 月，格式 yyyy-mm。"""
    months_ago = 2 if folder_name == "日报" else 1
    start_m = int(shared_date.split("-")[0].split(".")[0])
    target = date(report_date.year, start_m, 1)
    for _ in range(months_ago):
        target = (target.replace(day=1) - timedelta(days=1)).replace(day=1)
    return f"{target.year:04d}-{target.month:02d}"


def fetch_total_fba_fee(snapshot_month: str) -> float:
    """从 amz_warehouse_rent_snapshot 汇总指定月的仓租（替代 Excel「FBA仓租费」合计）。"""
    sql = f"""
        SELECT COALESCE(SUM(`rent_fee`), 0) AS total_fba_fee
        FROM `{RENT_TABLE}`
        WHERE `snapshot_month` = %s
          AND `is_deleted` = 0
    """
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            cur.execute(sql, (snapshot_month,))
            row = cur.fetchone() or {}
            return float(row.get("total_fba_fee") or 0)
    finally:
        conn.close()


# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-14)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)

# 用快照表代替 (处理完成)FBA仓租明细{fba_date}.xlsx；snapshot_month 对齐 fba_date 归属月
snapshot_month = resolve_snapshot_month()
total_fba_fee = fetch_total_fba_fee(snapshot_month)
print(f"snapshot_month={snapshot_month}, total_fba_fee={total_fba_fee}")

# 获取当前月份的天数
days_in_now_month = pd.Timestamp.now().daysinmonth
# 获取 当前日报的天数
d1, d2 = (int(part.split('.')[1]) for part in shared_date.split('-'))
days = d2 - d1 + 1

# 按 销量比例，用 上月总的FBA仓租费 去计算 日报 的  FBA仓租费
# 筛选 平台 包含 AMAZON
mask = main_file_df['平台'].str.contains('AMAZON', case=False, na=False)
# 计算销量占比（仅在 AMAZON 行）
amazon_sales = main_file_df.loc[mask, '销量']
# 得到 当前月份、按销量比例 * 上月的FBA仓租费 的   一天的FBA仓租费
main_file_df.loc[mask, 'FBA仓租费'] = np.round(amazon_sales / amazon_sales.sum() * total_fba_fee / days_in_now_month * days, 2)
# 空值的地方——补 0
main_file_df = main_file_df.fillna(0)
# 保存修改
output_path = main_file_path.replace('已完成-14', '已完成-15')
main_file_df.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
print('太小的FBA仓租费（日租），保留2位小数后，会变成0！！！！')
