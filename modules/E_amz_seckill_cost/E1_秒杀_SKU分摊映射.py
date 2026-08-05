"""E0：从 DB amz_seckill_cost 读取秒杀费用，映射站点/平台识别码。

相对 E1（读 xlsx）的差异：
  - 数据源改为 amz_seckill_cost
  - seckill_fee 含「+」的行忽略（多为「€4.00 per day +0.75% of sales」未结算公式）
  - seckill_fee 可解析且 settle_status≠1 时回填 charge_amount / currency_code，
    用 updated_at 回填 charge_date；同时 settle_status=1，settle_batch_no=charge_date 月份（yyyy-mm）；
    若 seckill_sku 为空则回填 seckill_sku=seckill_goods
  - 已结算（settle_status=1）不再回填
  - 输出 xlsx 使用 settle_status=1 且 settle_batch_no 与报表 end_date 同月的数据
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
from config.A0_set_date import folder_name, shared_date, test_end_date
from database.db_connection import get_db_manager

TABLE = "amz_seckill_cost"
_FEE_NUM_RE = re.compile(r"[\d]+(?:\.\d+)?")
_CURRENCY_CODE_RE = re.compile(r"\b(USD|EUR|GBP|JPY|CAD|MXN|AUD|SEK|PLN|TRY|INR|BRL)\b", re.I)
_CURRENCY_SYMBOL_MAP = {
    "€": "EUR",
    "£": "GBP",
    "$": "USD",
    "¥": "JPY",
    "￥": "CNY",
}


def _batch_no_from_end_date(end_date: str) -> str:
    """报表 end_date → settle_batch_no（yyyy-mm）。"""
    return pd.to_datetime(end_date).strftime("%Y-%m")


def _parse_seckill_fee_detail(val) -> tuple[float, str | None] | None:
    """从 seckill_fee 提取 (金额, 币种)；含「+」或无法解析金额则返回 None。

    例：``Fees €68.52`` → ``(68.52, "EUR")``；无币种符号时币种为 None。
    """
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    s = str(val).strip()
    if not s or "+" in s:
        return None
    m = _FEE_NUM_RE.search(s.replace(",", ""))
    if not m:
        return None
    try:
        amount = float(m.group())
    except ValueError:
        return None

    currency: str | None = None
    for sym, code in _CURRENCY_SYMBOL_MAP.items():
        if sym in s:
            currency = code
            break
    if currency is None:
        cm = _CURRENCY_CODE_RE.search(s)
        if cm:
            currency = cm.group(1).upper()
    return amount, currency


def _is_blank(val) -> bool:
    """空值 / 空白 / 常见空字符串占位视为 blank。"""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return True
    s = str(val).strip()
    return s == "" or s.lower() in ("nan", "none")


def fetch_unsettled_for_backfill() -> pd.DataFrame:
    """读取尚未结算（settle_status≠1）的记录，供回填。"""
    sql = f"""
        SELECT
            id,
            seckill_fee,
            seckill_sku,
            seckill_goods,
            updated_at,
            settle_status
        FROM `{TABLE}`
        WHERE settle_status <> 1
    """
    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        cur.execute(sql)
        rows = cur.fetchall()
    finally:
        cur.close()
        conn.close()
    return pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=[
            "id",
            "seckill_fee",
            "seckill_sku",
            "seckill_goods",
            "updated_at",
            "settle_status",
        ]
    )


def fetch_settled_for_report(batch_no: str) -> pd.DataFrame:
    """读取已结算且 settle_batch_no 与报表 end_date 同月的记录，供输出 xlsx。"""
    sql = f"""
        SELECT
            shop_name_en,
            marketplace,
            seckill_sku,
            seckill_goods,
            charge_amount,
            settle_batch_no
        FROM `{TABLE}`
        WHERE settle_status = 1
          AND settle_batch_no = %s
    """
    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        cur.execute(sql, (batch_no,))
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
            "charge_amount",
            "settle_batch_no",
        ]
    )


def backfill_charge_from_seckill_fee(df: pd.DataFrame) -> int:
    """seckill_fee 可解析且未结算时回填扣费字段，并标记已结算。

    settle_status=1 的行跳过（调用方已过滤；UPDATE 再加条件防并发重复回填）。
    seckill_sku 为空时同步回填 seckill_sku = seckill_goods。
    """
    if df.empty:
        return 0

    updates: list[tuple] = []
    skipped_no_currency = 0
    for _, row in df.iterrows():
        if int(row.get("settle_status") or 0) == 1:
            continue
        detail = _parse_seckill_fee_detail(row.get("seckill_fee"))
        if detail is None:
            continue
        amount, currency = detail
        if not currency:
            skipped_no_currency += 1
            continue
        updated_at = row.get("updated_at")
        if updated_at is None or (isinstance(updated_at, float) and np.isnan(updated_at)):
            continue
        charge_date = pd.to_datetime(updated_at).date()
        settle_batch_no = charge_date.strftime("%Y-%m")
        row_id = row.get("id")
        if row_id is None or (isinstance(row_id, float) and np.isnan(row_id)):
            continue
        # seckill_sku 为空时用 seckill_goods；否则保留原值（SQL 用 COALESCE 语义由参数表达）
        sku = row.get("seckill_sku")
        goods = row.get("seckill_goods")
        fill_sku = None if not _is_blank(sku) else (
            None if _is_blank(goods) else str(goods).strip()
        )
        updates.append(
            (amount, currency, charge_date, settle_batch_no, fill_sku, int(row_id))
        )

    if skipped_no_currency:
        print(f"[DB] seckill_fee 可解析金额但缺币种，跳过回填 {skipped_no_currency} 行")

    if not updates:
        return 0

    sql = f"""
        UPDATE `{TABLE}`
        SET charge_amount = %s,
            currency_code = %s,
            charge_date = %s,
            settle_status = 1,
            settle_batch_no = %s,
            seckill_sku = COALESCE(%s, seckill_sku)
        WHERE id = %s
          AND settle_status <> 1
    """
    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor()
    try:
        cur.executemany(sql, updates)
        conn.commit()
        return cur.rowcount if cur.rowcount is not None and cur.rowcount >= 0 else len(updates)
    finally:
        cur.close()
        conn.close()


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

output_dir = Path(DESKTOP_ROOT) / f"{folder_name}{shared_date}" / "秒杀"
output_file_path = output_dir / f"(处理完成)秒杀数据-{shared_date}.xlsx"
settle_batch_no = _batch_no_from_end_date(test_end_date)

# 1) 未结算行：可解析则回填（已 settle_status=1 的不处理）
unsettled_df = fetch_unsettled_for_backfill()
print(f"[DB] {TABLE} 未结算 settle_status≠1 共 {len(unsettled_df)} 行")
_n_backfill = backfill_charge_from_seckill_fee(unsettled_df)
if _n_backfill:
    print(f"[DB] 已回填扣费并标记结算 {_n_backfill} 行")
else:
    print("[DB] 无需回填（无可解析费用或均已结算）")

# 2) 已结算且 settle_batch_no 与报表 end_date 同月 → 输出 xlsx
main_df = fetch_settled_for_report(settle_batch_no)
print(
    f"[DB] settle_status=1 且 settle_batch_no={settle_batch_no!r} "
    f"读到 {len(main_df)} 行（报表 end_date={test_end_date}）"
)
if main_df.empty:
    print(f"[WARN] 无已结算数据，仍写出空表：{output_file_path}")

main_df = main_df.copy()
main_df["秒杀费"] = pd.to_numeric(main_df.get("charge_amount"), errors="coerce")
main_df = main_df[main_df["秒杀费"].notna()].copy()

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
