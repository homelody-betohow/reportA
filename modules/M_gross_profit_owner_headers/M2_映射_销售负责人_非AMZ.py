"""
M2_映射_销售负责人_非AMZ — 订单统计补全非 AMZ 销售负责人（20→21）

【流水线位置】
  上游：(已完成-20)订单统计（M1 计算毛利）
  下游：(已完成-21) → M3 映射 AMZ 销售负责人

【映射规则】
  1. 指定站点列表 → 「信息-映射.xlsx」按站点映射销售负责人
  2. 其余行 → 按平台映射销售负责人
  3. MANO-EU 且站点含 BTH → nobody
  4. 空白 / 「无负责人」 → nobody

用法：
  python modules/M_gross_profit_owner_headers/M2_映射_销售负责人_非AMZ.py
"""

from __future__ import annotations

import importlib.util
from pathlib import Path

import pandas as pd

# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.sku_mapping import sku_mappings  # noqa: E402
from config.A0_paths import DESKTOP_ROOT  # noqa: E402
from config.A0_set_date import folder_name, shared_date  # noqa: E402

OWNER_COL = "销售负责人"
NOBODY = "nobody"
_BLANK_OWNERS = frozenset(("", "nan", "None", "NaN", "无负责人"))

INFO_MAP_XLSX = fr"{DESKTOP_ROOT}\信息-映射.xlsx"
OWNER_SHEET = "销售负责人"

# 按站点映射销售负责人（其余按平台）
STATION_OWNER_KEYS = frozenset(
    {
        "TEMU-AIH",
        "TEMU-BV",
        "TEMU-HM",
        "TEMU-AL",
        "LM-TOTO",
        "LM-ES-BTH",
        "LM-FR-BTH",
        "LM-IT-BTH",
        "LM-PL-BTH",
        "LM-PT-BTH",
        "TEMU-KR-A",
        "TEMU-KR-B",
        "TEMU-KR-C",
        "TEMU-HJ-A",
        "TEMU-HJ-B",
        "TEMU-HJ-C",
        "TEMU-NF-A",
        "TEMU-NF-B",
        "TEMU-NF-C",
        "LM-FR-BC-ls",
        "LM-FR-BC-xj",
        "LM-ES-BC-ls",
        "LM-ES-BC-xj",
        "LM-PT-BC-ls",
        "LM-PT-BC-xj",
        "LM-IT-BC-ls",
        "LM-IT-BC-xj",
        "TEMU-BZ",
        "TEMU-AQ",
        "LM-FR-RP-ls",
        "LM-FR-RP-xj",
        "LM-ES-RP-ls",
        "LM-ES-RP-xj",
        "LM-PT-RP-ls",
        "LM-PT-RP-xj",
        "LM-IT-RP-ls",
        "LM-IT-RP-xj",
    }
)


def _input_path() -> str:
    return (
        fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计"
        fr"\(已完成-20)订单统计-{shared_date}.xlsx"
    )


def _map_owner(
    df: pd.DataFrame,
    *,
    main_sku: str,
    map_new_sku: str,
) -> pd.DataFrame:
    """用信息-映射表映射销售负责人，并统一列名为「销售负责人」。"""
    if df.empty:
        out = df.copy()
        if OWNER_COL not in out.columns:
            out[OWNER_COL] = None
        return out

    mapped = sku_mappings(
        main_df=df.copy(),
        main_sku=main_sku,
        map_sku_path=INFO_MAP_XLSX,
        map_old_sku=main_sku,
        map_new_sku=map_new_sku,
        map_sku_sheet=OWNER_SHEET,
    )
    return mapped.rename(columns={f"映射{map_new_sku}": OWNER_COL})


def _apply_mano_bth_nobody(df: pd.DataFrame) -> pd.DataFrame:
    """MANO-EU 且站点含 BTH → nobody。"""
    out = df.copy()
    mask = (out["平台"] == "MANO-EU") & out["站点"].str.contains("BTH", na=False)
    out.loc[mask, OWNER_COL] = NOBODY
    return out


def _normalize_owner(df: pd.DataFrame) -> pd.DataFrame:
    """空白或「无负责人」统一为 nobody。"""
    out = df.copy()
    owner = out[OWNER_COL]
    blank = owner.isna() | owner.astype(str).str.strip().isin(_BLANK_OWNERS)
    out.loc[blank, OWNER_COL] = NOBODY
    return out


def map_non_amz_owners(df: pd.DataFrame) -> pd.DataFrame:
    """非 AMZ：按站点/平台映射销售负责人，并做业务收尾。"""
    by_station = df["站点"].isin(STATION_OWNER_KEYS)
    st_df = _map_owner(
        df.loc[by_station],
        main_sku="站点",
        map_new_sku="销售负责人-站点",
    )
    platform_df = _map_owner(
        df.loc[~by_station],
        main_sku="平台",
        map_new_sku="销售负责人-平台",
    )
    out = pd.concat([st_df, platform_df], ignore_index=True)
    out = _apply_mano_bth_nobody(out)
    return _normalize_owner(out)


def main() -> int:
    input_path = _input_path()
    src = pd.read_excel(input_path)
    result = map_non_amz_owners(src)

    output_path = input_path.replace("已完成-20", "已完成-21")
    result.to_excel(output_path, index=False)
    print(f"结果已保存到 {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
