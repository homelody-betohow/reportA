"""
M2_映射_销售负责人_非AMZ — 订单统计补全非 AMZ 销售负责人（20→21）

【流水线位置】
  上游：(已完成-20)订单统计（M1 计算毛利）
  下游：(已完成-21) → M3 映射 AMZ 销售负责人

【映射规则】
  1. 站点落在 config.A0_station_sales_owner → 直接用配置负责人
  2. 其余行 → 用「平台 + 商品ID」匹配 MONTH_GOAL_EXCEL_PATH（ALL sheet）的「负责人」
  3. 空白 / 「无负责人」 → nobody

"""

from __future__ import annotations

import importlib.util
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# 项目根引导：须在 import config / common 之前执行
# ---------------------------------------------------------------------------
_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from config.A0_paths import DESKTOP_ROOT, MONTH_GOAL_EXCEL_PATH  # noqa: E402
from config.A0_set_date import folder_name, shared_date  # noqa: E402
from config.A0_station_sales_owner import (  # noqa: E402
    STATION_OWNER_KEYS,
    STATION_SALES_OWNER,
)

OWNER_COL = "销售负责人"
NOBODY = "nobody"
_BLANK_OWNERS = frozenset(("", "nan", "None", "NaN", "无负责人"))

# 月目标 ALL：平台 + 商品ID → 负责人
_GOAL_SHEET = "ALL"
_GOAL_COLS = ("平台", "商品ID", "负责人")


def _input_path() -> str:
    """拼出 M1 产出的 (已完成-20) 订单统计 Excel 绝对路径。"""
    return (
        fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计"
        fr"\(已完成-20)订单统计-{shared_date}.xlsx"
    )


def _norm_key(series: pd.Series) -> pd.Series:
    """匹配键标准化：转字符串并去首尾空白。"""
    return series.astype(str).str.strip()


def _load_goal_owner_map() -> dict[tuple[str, str], str]:
    """
    从月目标 ALL sheet 构建 {(平台, 商品ID): 负责人}。

    同键多行时保留首次出现的非空负责人（AMZ 冲突由 M3 再覆盖）。
    """
    goal = pd.read_excel(
        MONTH_GOAL_EXCEL_PATH,
        sheet_name=_GOAL_SHEET,
        usecols=list(_GOAL_COLS),
    )
    goal = goal.dropna(subset=["平台", "商品ID"])
    goal["平台"] = _norm_key(goal["平台"])
    goal["商品ID"] = _norm_key(goal["商品ID"])
    owner = goal["负责人"]
    valid = owner.notna() & ~owner.astype(str).str.strip().isin(_BLANK_OWNERS)
    goal = goal.loc[valid].drop_duplicates(subset=["平台", "商品ID"], keep="first")
    return {
        (p, i): str(o).strip()
        for p, i, o in zip(goal["平台"], goal["商品ID"], goal["负责人"])
    }


def _apply_station_owners(df: pd.DataFrame) -> pd.Series:
    """按站点配置映射销售负责人（未命中为 NA）。"""
    return df["站点"].map(STATION_SALES_OWNER)


def _apply_goal_owners(df: pd.DataFrame, owner_map: dict[tuple[str, str], str]) -> pd.Series:
    """按 平台+商品ID 查月目标负责人（未命中为 NA）。"""
    keys = list(zip(_norm_key(df["平台"]), _norm_key(df["商品ID"])))
    return pd.Series([owner_map.get(k) for k in keys], index=df.index, dtype=object)


def _normalize_owner(df: pd.DataFrame) -> pd.DataFrame:
    """空白或「无负责人」统一为 nobody。"""
    out = df.copy()
    owner = out[OWNER_COL]
    blank = owner.isna() | owner.astype(str).str.strip().isin(_BLANK_OWNERS)
    out.loc[blank, OWNER_COL] = NOBODY
    return out


def map_non_amz_owners(
    df: pd.DataFrame,
    *,
    owner_map: dict[tuple[str, str], str] | None = None,
) -> pd.DataFrame:
    """
    非 AMZ 销售负责人主流程（对整表订单统计生效，保留原行序）。

    步骤：
      1. 站点在 STATION_OWNER_KEYS → 用 A0_station_sales_owner 配置
      2. 其余 → 平台+商品ID 查月目标 ALL「负责人」
      3. 空白/无负责人规范化为 nobody
    """
    out = df.copy()
    if owner_map is None:
        owner_map = _load_goal_owner_map()

    by_station = out["站点"].isin(STATION_OWNER_KEYS)
    station_owner = _apply_station_owners(out)
    goal_owner = _apply_goal_owners(out, owner_map)

    out[OWNER_COL] = goal_owner
    out.loc[by_station, OWNER_COL] = station_owner.loc[by_station]

    return _normalize_owner(out)


def main() -> int:
    """读 (已完成-20) → 映射负责人 → 写 (已完成-21)。"""
    input_path = _input_path()
    src = pd.read_excel(input_path)
    result = map_non_amz_owners(src)

    output_path = input_path.replace("已完成-20", "已完成-21")
    result.to_excel(output_path, index=False)
    print(f"结果已保存到 {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
