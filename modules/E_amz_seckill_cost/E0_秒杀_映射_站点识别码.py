"""E0：从 DB amz_seckill_cost 读取秒杀费用，映射站点/平台识别码。

相对 E1（读 xlsx）的差异：
  - 数据源改为 amz_seckill_cost
  - seckill_fee 含「+」的行忽略（多为「€4.00 per day +0.75% of sales」未结算公式）
  - SKU 优先 seckill_sku，为空则用 seckill_goods
  - 站点由 shop_name_en → platform_shop.market_region 映射

输出：与 E1 相同，供 E2 合并使用
  {DESKTOP_ROOT}/{folder_name}{shared_date}/秒杀/(处理完成)秒杀数据-{shared_date}.xlsx
"""

from __future__ import annotations

import importlib.util
import re
from pathlib import Path

import numpy as np
import pandas as pd
import pymysql.cursors

# 须在 import config/common 之前：加载项目根到 sys.path
_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.platform_shop import map_region_to_platform, map_shop_to_region
from common.split_rows_data_SKU import split_one_rows_data
from config.A0_paths import DESKTOP_ROOT
from config.A0_set_date import folder_name, shared_date, test_end_date, test_start_date
from database.db_connection import get_db_manager

TABLE = "amz_seckill_cost"
_FEE_NUM_RE = re.compile(r"[\d]+(?:\.\d+)?")


def _parse_seckill_fee(val) -> float | None:
    """从 seckill_fee 文本提取金额；含「+」或无法解析则返回 None（忽略）。"""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    s = str(val).strip()
    if not s or "+" in s:
        return None
    m = _FEE_NUM_RE.search(s.replace(",", ""))
    if not m:
        return None
    try:
        return float(m.group())
    except ValueError:
        return None


def fetch_seckill_cost_df(start_date: str, end_date: str) -> pd.DataFrame:
    """读取报表区间内与秒杀活动日期有重叠的记录。"""
    sql = f"""
        SELECT
            shop_name_en,
            marketplace,
            seckill_sku,
            seckill_goods,
            seckill_fee,
            start_date,
            end_date
        FROM `{TABLE}`
        WHERE DATE(start_date) <= %s
          AND DATE(end_date) >= %s
    """
    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        cur.execute(sql, (end_date, start_date))
        rows = cur.fetchall()
    finally:
        cur.close()
        conn.close()
    return pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=[
            "shop_name_en",
            "marketplace",
            "seckill_sku",
            "seckill_goods",
            "seckill_fee",
            "start_date",
            "end_date",
        ]
    )


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

output_dir = Path(DESKTOP_ROOT) / f"{folder_name}{shared_date}" / "秒杀"
output_file_path = output_dir / f"(处理完成)秒杀数据-{shared_date}.xlsx"

main_df = fetch_seckill_cost_df(test_start_date, test_end_date)
print(
    f"[DB] {TABLE} 区间 {test_start_date}~{test_end_date} 读到 {len(main_df)} 行"
)

# seckill_fee 含「+」或无法解析 → 忽略
main_df["秒杀费"] = main_df["seckill_fee"].map(_parse_seckill_fee)
_ignored = main_df["秒杀费"].isna().sum()
main_df = main_df[main_df["秒杀费"].notna()].copy()
if _ignored:
    print(f"[DB] 忽略 seckill_fee 含「+」或无法解析 {_ignored} 行，剩余 {len(main_df)} 行")

# SKU：优先 seckill_sku，为空则用 seckill_goods
_sku = main_df["seckill_sku"].astype(str).str.strip()
_goods = main_df["seckill_goods"].astype(str).str.strip()
_sku = _sku.mask(_sku.isin(["", "nan", "None", "none"]), "")
main_df["SKU"] = _sku.where(_sku != "", _goods)

# 店铺 → 站点（platform_shop.market_region）
main_df = map_shop_to_region(main_df, shop_col="shop_name_en", region_col="站点")

# 拆分有「+」的 SKU，秒杀费均摊
main_df_1 = split_one_rows_data(
    input_df=main_df,
    data_column="SKU",
    value_column="秒杀费",
)

# 使用字符串操作合并两列
main_df_1["SKU-站点识别码"] = main_df_1["站点"] + main_df_1["SKU"]
# 映射 平台（数据源：platform_shop）
main_df_2 = map_region_to_platform(main_df_1, site_col="站点")
# 在 SKU-站点识别码 后插入 SKU-平台识别码
new_column_name = "SKU-平台识别码"
new_column_data = main_df_2["映射平台"] + main_df_2["SKU"]
target_column = "SKU-站点识别码"
insert_position = main_df_2.columns.get_loc(target_column) + 1
main_df_2.insert(insert_position, new_column_name, new_column_data)

main_df_2 = main_df_2.rename(columns={"映射平台": "平台"})

# 按照 'SKU-站点识别码' 列进行分组，并对 '秒杀费' 列进行汇总
grouped_main_df = (
    main_df_2.groupby("SKU-站点识别码")
    .agg(
        {
            "秒杀费": "sum",
            "SKU": "first",
            "站点": "first",
            "平台": "first",
            "SKU-平台识别码": "first",
        }
    )
    .reset_index()
)

grouped_main_df["秒杀费"] = pd.to_numeric(grouped_main_df["秒杀费"], errors="coerce")
grouped_main_df["秒杀费"] = np.round(grouped_main_df["秒杀费"], 2)
grouped_main_df = grouped_main_df[
    ["SKU", "站点", "平台", "SKU-站点识别码", "SKU-平台识别码", "秒杀费"]
]

output_dir.mkdir(parents=True, exist_ok=True)
grouped_main_df.to_excel(output_file_path, index=False)
print(f"处理完成，输出文件路径：{output_file_path}")
