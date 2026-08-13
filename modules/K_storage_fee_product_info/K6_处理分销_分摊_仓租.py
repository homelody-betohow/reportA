"""
K6_处理分销_分摊_仓租 — 分销口径收尾 + 无平台仓租分摊（17→18）

【流水线位置】
  上游：(已完成-17)订单统计（K5；月报时 MANO 仓租已由 C2 写入「FBA仓租费」）
  下游：(已完成-18)订单统计 → 毛利等后续脚本

【核心处理】
  1. 分销收尾：智慧谷采购成本置 0；分销行运营模式/分类、头程关税派送费
  2. 将「所有仓库-无平台-需要分摊的费用」总额，仅摊给「非分销 且 原-海外仓仓租费>0」
     的行，按这些行「平台销售额」占比 →「仓租分摊」（不按全场销售额）
  3. 海外仓仓租费 = 原-海外仓仓租费 + 仓租分摊；仓租合计 = FBA仓租费 + 海外仓仓租费

"""

import importlib.util
from pathlib import Path

import pandas as pd
# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.cang_zu_decimal import round_rent, round_rent_series  # noqa: E402
from config.A0_set_date import shared_date, folder_name
from config.A0_paths import DESKTOP_ROOT

_COL_SALES = "平台销售额"
_COL_OVERSEAS = "海外仓仓租费"
_COL_OVERSEAS_ORIG = "原-海外仓仓租费"
_RENT_EPS = 1e-8


def _unallocated_rent_total(df: pd.DataFrame) -> float:
    """K4 写入的无平台仓租总额（通常仅在第 1 行有值）。"""
    col = df["所有仓库-无平台-需要分摊的费用"]
    if col.notna().any():
        return float(round_rent(col[col.notna()].iloc[0]))
    return float(round_rent(col.fillna(0).sum()))


def _allocate_unplatform_rent(
    df: pd.DataFrame, participate: pd.Series, total_cost: float
) -> pd.Series:
    """
    无平台仓租按参与行「平台销售额」占比分摊。
    销售额权重：数值化后空值/负数按 0；若合计销售额 ≤ 0，回退为按行均摊（保总额）。
    无参与行且总额 > 0 时不摊（总额留在第 1 行无平台列，避免摊到无关 SKU）。
    """
    alloc = pd.Series(0.0, index=df.index, dtype=float)
    n = int(participate.sum())
    if float(total_cost) == 0.0:
        return alloc
    if n <= 0:
        print(
            f"[无平台分摊] 总额={total_cost:.4f}，参与行=0，"
            f"无法摊到「原-海外仓仓租费>0」的非分销行（仓租分摊全 0）"
        )
        return alloc

    if _COL_SALES not in df.columns:
        raise KeyError(f"订单统计缺少列「{_COL_SALES}」，无法按销售额占比分摊无平台仓租")

    sales = (
        pd.to_numeric(df.loc[participate, _COL_SALES], errors="coerce")
        .fillna(0)
        .clip(lower=0)
    )
    sales_sum = float(sales.sum())
    if sales_sum > 0:
        alloc.loc[participate] = sales / sales_sum * total_cost
        print(
            f"[无平台分摊] 总额={total_cost:.4f}，参与行={n}，"
            f"平台销售额合计={sales_sum:.4f}（按销售额占比）"
        )
    else:
        alloc.loc[participate] = total_cost / n
        print(
            f"[无平台分摊] 总额={total_cost:.4f}，参与行={n}，"
            f"平台销售额合计=0，回退按行均摊"
        )
    return alloc

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-17)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)

# 注意：供应商→分销 的规则已在 K5 执行；此处不再重复覆盖产品状态/分类，
# 以免覆盖你在 (已完成-17) 中手动修正的 保留品/新品 等状态。
# 如果供应商是“智慧谷”，则将“采购成本”、“订单采购成本”、“重发采购成本”、“二次上架采购成本”替换为 0
main_file_df.loc[
    main_file_df['供应商'] == '智慧谷', ['采购成本', '订单采购成本', '重发采购成本', '二次上架采购成本']] = 0
# 如果产品状态是“分销”，则将“运营模式”替换为 “自运营”；“二级分类”、“三级分类”替换为 “分销”
main_file_df.loc[main_file_df['产品状态'] == '分销', ['运营模式', '二级分类', '三级分类']] = ['自运营', '分销', '分销']
# 如果产品状态是“分销”，则将“头程”、“关税”、“派送费”替换为 0
main_file_df.loc[main_file_df['产品状态'] == '分销', ['头程', '关税', '派送费']] = 0

# 先定「原-海外仓仓租费」，再筛参与行（方案 B：只摊给已挂上海外仓仓租的行）
if _COL_OVERSEAS not in main_file_df.columns:
    raise KeyError(f"订单统计缺少列「{_COL_OVERSEAS}」，无法按海外仓相关行分摊无平台仓租")
main_file_df = main_file_df.rename(columns={_COL_OVERSEAS: _COL_OVERSEAS_ORIG})
main_file_df[_COL_OVERSEAS_ORIG] = round_rent_series(
    main_file_df[_COL_OVERSEAS_ORIG]
).fillna(0)

# 参与 = 非分销 且 原-海外仓仓租费>0；权重仍为这些行的平台销售额
_participate = (main_file_df["产品状态"] != "分销") & (
    pd.to_numeric(main_file_df[_COL_OVERSEAS_ORIG], errors="coerce").fillna(0)
    > _RENT_EPS
)
# 月报时再排除 MANO-EU（其海外仓仓租已由 ManoRent 统计）
# if folder_name == '月报':
#     _mano_eu = main_file_df['平台'].astype(str).str.strip() == 'MANO-EU'
#     _participate = _participate & ~_mano_eu
#     print(f'[月报] 忽略 MANO-EU 无平台仓租分摊：排除 {_mano_eu.sum()} 行（MANO 仓租已在 FBA仓租费）')

total_cost = _unallocated_rent_total(main_file_df)
main_file_df["仓租分摊"] = _allocate_unplatform_rent(
    main_file_df, _participate, total_cost
)
main_file_df["仓租分摊"] = round_rent_series(main_file_df["仓租分摊"]).fillna(0)
main_file_df[_COL_OVERSEAS] = round_rent_series(
    main_file_df[_COL_OVERSEAS_ORIG] + main_file_df["仓租分摊"]
)

# 产品状态仍为空（含空串/空白）→ 赋 "--"（须在 fillna(0) 之前，且排除该列）
_status = main_file_df["产品状态"]
_empty_status = _status.isna() | _status.astype(str).str.strip().isin(["", "nan", "None"])
main_file_df.loc[_empty_status, "产品状态"] = "--"

# 空值的地方——补 0   使 仓租合计  可以正常合计；平台、平台商品ID识别码、产品状态 不补 0
_exclude_fill0 = {'平台', '平台商品ID识别码', '产品状态'}
_fill0_cols = [c for c in main_file_df.columns if c not in _exclude_fill0]
main_file_df[_fill0_cols] = main_file_df[_fill0_cols].fillna(0)

main_file_df["仓租合计"] = round_rent_series(
    pd.to_numeric(main_file_df["FBA仓租费"], errors="coerce").fillna(0)
    + pd.to_numeric(main_file_df["海外仓仓租费"], errors="coerce").fillna(0)
)

# 保存修改后的文件
output_path = main_file_path.replace('已完成-17', '已完成-18')
main_file_df.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
