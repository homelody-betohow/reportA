"""
cang_zu_site — 仓租流程：库存「平台」标签 → 标准「站点」

【用途】
  K1_HY_仓租 / K2_4PX_仓租 在按库存拆分仓租后，库存表「平台」列里的值
  并不全是订单统计口径的「站点」（例如 AMAZON-EU、MANO-EU），需要先映射成站点，
  再拼「站点商品ID识别码 = 站点 + 商品ID」。

【数据来源】
  原桌面「仓租-站点映射.xlsx」Sheet1；现固化为本模块 PLATFORM_TO_SITE。

【字典约定】
  - 只收录「需要改名」的映射（键 ≠ 值），例如 AMAZON-EU → AMAZON-DE。
  - 键 = 值的恒等项已删除；未命中时由 map_platform_to_site 回填原平台名
    （相当于恒等：LM-BTH → LM-BTH、TEMU-AIH → TEMU-AIH 等）。
  - 「无」「其他」及空值不作为有效站点，映射结果置空（与下游 K3 过滤一致）。
"""

from __future__ import annotations

import pandas as pd

# 无效 / 占位「平台」标签：映射站点置空（不回填原文）
_INVALID_PLATFORM_LABELS: frozenset[str] = frozenset({"", "nan", "None", "无", "其他"})

# 平台（库存标签）→ 站点：仅「需要改名」的条目
# 同键多写时以后者为准；恒等映射不必写，见模块说明
PLATFORM_TO_SITE: dict[str, str] = {
    # Amazon 区域标签 → 统一落到 AMAZON-DE 站点
    "AMAZON-EU": "AMAZON-DE",
    "AMAZON-US": "AMAZON-US",
    # 简称 / 别名 → 标准站点
    "CD": "chengyi-CD",
    "DLZ-EU": "DLZ-DE",
    "MANO-EU": "MANO-FR",
    "OTTO": "OTTO-BTH",
    "REAL": "REAL-FB",
    "LM-FR": "LM-BTH",
}


def map_platform_to_site(
    main_df: pd.DataFrame,
    platform_col: str = "销售平台",
    result_col: str = "映射站点",
) -> pd.DataFrame:
    """
    按 PLATFORM_TO_SITE 写入「映射站点」列（紧挨 platform_col 之后）。

    规则
    ----
    1. 命中字典 → 用字典中的站点名
    2. 未命中 → 回填原「平台」字符串（恒等；故字典无需写键=值项）
    3. 原值为空，或为「无」「其他」→ 映射站点为 None

    参数
    ----
    main_df : DataFrame
        含库存「销售平台」列的仓租中间表
    platform_col : str
        源列名，默认「销售平台」
    result_col : str
        写出列名，默认「映射站点」
    """
    if platform_col not in main_df.columns:
        raise KeyError(f"主表缺少平台列：{platform_col}")

    df = main_df.copy()
    if result_col in df.columns:
        df = df.drop(columns=[result_col])

    insert_pos = df.columns.get_loc(platform_col) + 1
    keys = df[platform_col].astype(str).str.strip()
    is_invalid = df[platform_col].isna() | keys.isin(_INVALID_PLATFORM_LABELS)

    mapped = keys.map(PLATFORM_TO_SITE)
    # 未命中：回填原平台名（替代已删除的恒等字典项）
    mapped = mapped.where(mapped.notna(), keys)
    # 无效标签：站点置空，避免「无/其他」进入识别码
    mapped = mapped.where(~is_invalid, None)

    df.insert(insert_pos, result_col, mapped)
    return df
