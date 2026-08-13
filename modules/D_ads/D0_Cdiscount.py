"""
D0_Cdiscount.py — Cdiscount 广告费：数据库 → D5 同款 Excel

在广告流水线中的位置（runAll_D.py）：
  D1 OTTO / D2 REAL / D3 MANO / D4 DLZ 读桌面 CSV；
  本脚本读数据库 fee_advertising，只处理 Cdiscount；
  然后 D5 把各平台「处理完成」Excel 合成一张，D6 再并入订单统计。

为什么单独写这个脚本：
  Cdiscount 广告已经落在表 fee_advertising 里，没有桌面 CSV。
  D5 / D6 认的是「仓库 SKU + 站点」识别码，不能直接拿刊登码去对订单。

本脚本只取两行条件（不要改成读全表）：
  1) platform = cdiscount          （库里小写；报表平台码是 CD）
  2) charge_month = 报表开始日      （YYYY-MM-DD，见 _period_start_ymd）

处理步骤：
  1. 按 charge_month + platform 从 fee_advertising 取数
  2. 站点 = market_region（这张表里已经是站点，不是账号名）
  3. SKU 优先 product_sku（仓库 SKU）；为空才用 sku_code
  4. 金额转到欧元；去掉 0 花费
  5. 组合 SKU（含 + 或 ,）拆行，广告费均摊
  6. 映射平台固定写 CD；生成 SKU-站点识别码、SKU-平台识别码
  7. 按 SKU-站点识别码 汇总（同一 SKU 多个 ad_id 要加总）

输出（列必须与 D1~D4 一致，D5 用 usecols 读这 6 列）：
  SKU, 站点, 映射平台, SKU-站点识别码, SKU-平台识别码, 广告费(非AMZ)
  路径：桌面/{日报|月报}{shared_date}/广告/CD/(处理完成)CD广告.xlsx

用法：
  python modules/D_ads/D0_Cdiscount.py
"""
from __future__ import annotations

import importlib.util
from datetime import date
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
import pymysql.cursors

# 须在 import config/common 之前：把项目根加入 sys.path，否则下面的包导入会失败
_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.split_rows_data_SKU import split_one_rows_data
from common.style import Color
from config.A0_paths import DESKTOP_ROOT
from config.A0_set_date import (
    CAD_to_EUR,
    Ft_to_EUR,
    Lei_to_EUR,
    RMB_di_EUR,
    USD_to_EUR,
    folder_name,
    kc_to_EUR,
    kr_to_EUR,
    report_date,
    shared_date,
    zl_to_EUR,
)
from database.db_connection import get_db_manager

# ---------------------------------------------------------------------------
# 业务常量：库里的平台名 ≠ 订单统计里的平台名
#   fee_advertising.platform = "cdiscount"
#   订单统计 / D5 的「平台」列 = "CD"（见 B7、毛利表）
# 识别码必须用 CD，否则 D6 按 SKU-站点识别码 合并时对不上订单。
# ---------------------------------------------------------------------------
TABLE = "fee_advertising"
AD_PLATFORM = "cdiscount"
REPORT_PLATFORM = "CD"

# 写出 Excel 时用的中文列名，与 D1_OTTO / D2_REAL / D3_MANO 对齐
SITE_COL = "站点"
SKU_COL = "SKU"
FEE_COL = "广告费(非AMZ)"

# 原币种 → 欧元。汇率来自 config/A0_set_date.py（改汇率只改那一处）。
# _FX_MUL：原币 × 汇率 = EUR（美元、克朗等）
# _FX_DIV：原币 ÷ 汇率 = EUR（人民币在本项目里用除法）
_FX_MUL: dict[str, float] = {
    "EUR": 1.0,
    "USD": float(USD_to_EUR),
    "CZK": float(kc_to_EUR),
    "PLN": float(zl_to_EUR),
    "HUF": float(Ft_to_EUR),
    "CAD": float(CAD_to_EUR),
    "SEK": float(kr_to_EUR),
    "RON": float(Lei_to_EUR),
}
_FX_DIV: dict[str, float] = {
    "CNY": float(RMB_di_EUR),
    "RMB": float(RMB_di_EUR),
}

# D5 只读这 6 列；多写列也不影响，少写列会报错
_OUTPUT_COLS = [
    SKU_COL,
    SITE_COL,
    "映射平台",
    "SKU-站点识别码",
    "SKU-平台识别码",
    FEE_COL,
]


def _as_text(value: Any) -> str:
    """把数据库/Excel 里的空值、NaN、'None' 统一成空字符串，避免后续拼接识别码变成 'nan'。"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip()
    if s.lower() in ("nan", "none", "<na>"):
        return ""
    return s


def _period_start_ymd() -> str:
    """
    本报表区间的开始日，格式 YYYY-MM-DD，对应 fee_advertising.charge_month。

    report_date 来自 A0_set_date.py，是区间结束日：
      - 日报：今天再往前推 3 天（当月 1 号 ~ 该日）
      - 月报：上个月最后一天（上月 1 号 ~ 月末）
    开始日一律是 report_date 所在月的 1 号，例如 2026-08-01。
    库里必须存零填充日期，不要写成 2026-8-1。
    """
    rd = report_date.date() if hasattr(report_date, "date") else report_date
    if not isinstance(rd, date):
        rd = pd.to_datetime(rd).date()
    return f"{rd.year:04d}-{rd.month:02d}-01"


def fetch_cdiscount_ads(charge_month: str) -> pd.DataFrame:
    """
    只拉 Cdiscount、且 charge_month 等于报表开始日的行。

    LOWER(TRIM(platform))：防止库里写成 'Cdiscount ' 带空格/大小写对不上。
    %s 是参数占位，不要把日期拼进 SQL 字符串（防注入，也避免引号出错）。
    无数据时仍返回带列名的空表，后面 prepare / to_excel 不会因缺列崩溃。
    """
    sql = f"""
        SELECT
            platform,
            market_region,
            shop_name_en,
            ad_id,
            product_sku,
            sku_code,
            ean_code,
            ad_expenditure,
            ad_currency,
            charge_month
        FROM `{TABLE}`
        WHERE charge_month = %s
          AND LOWER(TRIM(platform)) = %s
    """
    db = get_db_manager()
    conn = db.get_connection()
    # DictCursor：每行是 dict，用列名取值，不要用下标
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        cur.execute(sql, (charge_month, AD_PLATFORM))
        rows = cur.fetchall()
    finally:
        # 连接来自连接池，用完必须 close，归还连接
        cur.close()
        conn.close()
    return pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=[
            "platform",
            "market_region",
            "shop_name_en",
            "ad_id",
            "product_sku",
            "sku_code",
            "ean_code",
            "ad_expenditure",
            "ad_currency",
            "charge_month",
        ]
    )


def _to_eur(amount: float, currency: str) -> float:
    """单笔金额转到欧元。未配置的币种直接报错，避免默默用错汇率。"""
    code = (currency or "EUR").strip().upper() or "EUR"
    if code in _FX_DIV:
        denom = _FX_DIV[code]
        return float(amount) / denom if denom else float(amount)
    rate = _FX_MUL.get(code)
    if rate is None:
        raise ValueError(f"未配置币种汇率：{code!r}，请补 A0_set_date 或 _FX_MUL/_FX_DIV")
    return float(amount) * rate


def prepare_ad_fee_df(raw: pd.DataFrame) -> pd.DataFrame:
    """
    把库里的活动明细，收成 D5 要的「仓库 SKU × 站点」汇总表。

    库表一行 ≈ 一个广告活动（ad_id）；D5/D6 一行 ≈ 一个 SKU-站点。
    所以最后必须 groupby 识别码，把多个活动的花费加总。
    """
    if raw.empty:
        return pd.DataFrame(columns=_OUTPUT_COLS)

    df = raw.copy()
    # 步骤 1：字符串去首尾空格，避免 'CD-FR ' 和 'CD-FR' 对不上
    for col in df.columns:
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

    # 步骤 2：落到仓库 SKU
    #   product_sku = 仓库 SKU（入库时若已映射，优先用）
    #   sku_code    = 平台刊登 SKU，对不上订单统计里的 SKU
    product_sku = df["product_sku"].map(_as_text)
    sku_code = df["sku_code"].map(_as_text)
    df[SKU_COL] = product_sku.where(product_sku != "", sku_code)
    # 步骤 3：站点直接用 market_region（本表约定：已是站点，不是账号名）
    df[SITE_COL] = df["market_region"].map(_as_text)

    n_fallback = int((product_sku == "").sum())
    if n_fallback:
        print(
            f"{Color.YELLOW}[检查] product_sku 为空、改用 sku_code 的行：{n_fallback}。"
            f"若 sku_code 不是仓库 SKU，D6 将无法对上订单。{Color.RESET}"
        )

    # 步骤 4：SKU 或站点为空无法拼识别码，丢掉并打日志
    blank_sku = df[SKU_COL] == ""
    blank_site = df[SITE_COL] == ""
    if blank_sku.any() or blank_site.any():
        n_sku = int(blank_sku.sum())
        n_site = int(blank_site.sum())
        print(
            f"{Color.YELLOW}[检查] 将丢弃 SKU 空 {n_sku} 行、站点空 {n_site} 行{Color.RESET}"
        )
        df = df.loc[~blank_sku & ~blank_site].copy()
        if df.empty:
            return pd.DataFrame(columns=_OUTPUT_COLS)

    # 步骤 5：广告支出 → 欧元。errors='coerce'：脏数据变成 NaN，再填 0
    amounts = pd.to_numeric(df["ad_expenditure"], errors="coerce").fillna(0.0)
    currencies = df["ad_currency"].map(_as_text).replace("", "EUR")
    unknown_ccy = sorted(
        {c for c in currencies.unique() if c not in _FX_MUL and c not in _FX_DIV}
    )
    if unknown_ccy:
        raise ValueError(f"fee_advertising 存在未配置汇率的币种：{unknown_ccy}")

    df[FEE_COL] = [
        _to_eur(amt, ccy) for amt, ccy in zip(amounts.tolist(), currencies.tolist())
    ]
    # 与 OTTO/REAL 一样：当期没花钱的行不进报表
    df = df.loc[pd.to_numeric(df[FEE_COL], errors="coerce").fillna(0.0) != 0].copy()
    if df.empty:
        return pd.DataFrame(columns=_OUTPUT_COLS)

    # 步骤 6：组合品如 A+B 拆成两行，广告费按子 SKU 个数均摊（不是按销量）
    df = split_one_rows_data(input_df=df, data_column=SKU_COL, value_column=FEE_COL)

    # 步骤 7：识别码规则与 B2 订单统计相同：站点+SKU、平台+SKU，中间没有分隔符
    # 例：站点 CD-FR、SKU E51001 → SKU-站点识别码 = CD-FRE51001
    df["映射平台"] = REPORT_PLATFORM
    df["SKU-站点识别码"] = df[SITE_COL].astype(str) + df[SKU_COL].astype(str)
    df["SKU-平台识别码"] = REPORT_PLATFORM + df[SKU_COL].astype(str)

    # 步骤 8：同一 SKU-站点 可能对应多个广告活动，费用加总后才交给 D5
    # first：SKU/站点/平台在同一识别码下应相同，取第一行即可
    grouped = (
        df.groupby("SKU-站点识别码", dropna=False)
        .agg(
            {
                FEE_COL: "sum",
                SKU_COL: "first",
                SITE_COL: "first",
                "映射平台": "first",
                "SKU-平台识别码": "first",
            }
        )
        .reset_index()
    )
    grouped[FEE_COL] = np.round(
        pd.to_numeric(grouped[FEE_COL], errors="coerce").fillna(0.0), 2
    )
    grouped = grouped.loc[grouped[FEE_COL] != 0].copy()
    return grouped[_OUTPUT_COLS]


def main() -> None:
    # 1) 算出报表开始日，去库里取 Cdiscount 当月（或日报当月至今）广告
    charge_month = _period_start_ymd()
    raw = fetch_cdiscount_ads(charge_month)
    print(
        f"[DB] {TABLE}: platform={AD_PLATFORM}, charge_month={charge_month} "
        f"读到 {len(raw)} 行（映射平台={REPORT_PLATFORM}，shared_date={shared_date}）"
    )

    # 2) 清洗、换汇、拆组合、按识别码汇总
    out_df = prepare_ad_fee_df(raw)
    total_eur = float(out_df[FEE_COL].sum()) if not out_df.empty else 0.0

    # 3) 写到桌面广告/CD，与 OTTO/REAL/MANO 目录并列；D5 会读这个文件
    #    无数据也写出带表头的空表，避免 D5 因文件不存在而失败
    output_dir = Path(DESKTOP_ROOT) / f"{folder_name}{shared_date}" / "广告" / "CD"
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / "(处理完成)CD广告.xlsx"
    out_df.to_excel(output_path, index=False)
    print(
        f"{Color.GREEN}处理完成：{len(out_df)} 行，合计 EUR {total_eur:.2f}，"
        f"路径：{output_path}{Color.RESET}"
    )


if __name__ == "__main__":
    main()
