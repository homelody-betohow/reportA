"""
G2_3_Temu合并统计.py — TEMU 罚款写入订单统计「其他分摊费用」

功能：
  1. 读取 G2_2_Temu罚款映射.py 生成的 (已完成)TEMU-罚款汇总-{shared_date}.xlsx
  2. 按「SKU-站点识别」汇总结算金额
     - 无识别码（发货未命中）时回退为「店铺__未匹配罚款」，仍计入其他分摊费用
  3. 更新 H3 生成的 (已完成-13)订单统计：其他分摊费用 += 结算金额（原有值保留相加）
  4. 订单统计中不存在的识别码，追加为新行（平台 = TEMU）
     - SKU 有值时一并填充 SKU-平台识别码 = 平台 + SKU
"""

import importlib.util
import warnings
from pathlib import Path

import numpy as np
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

from common.style import Color
from config.A0_set_date import folder_name, shared_date
from config.A0_paths import DESKTOP_ROOT

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

SKU_SITE_ID_COL = "SKU-站点识别"
SKU_SITE_ID_COL_LEGACY = "儿子-站点识别"
ORDER_SKU_SITE_COL = "SKU-站点识别码"
ORDER_SKU_PLATFORM_COL = "SKU-平台识别码"
SETTLE_AMOUNT_COL = "结算金额"
ALLOC_COL = "其他分摊费用"
SKU_COL = "SKU"
SHOP_COL = "店铺"
SITE_COL = "站点"
PLATFORM_COL = "平台"
PLATFORM_TEMU = "TEMU"
UNMATCHED_SUFFIX = "__未匹配罚款"

FINE_INPUT_NAMES = (
    f"(已完成)TEMU-罚款汇总-{shared_date}.xlsx",
    f"(已完成)TEMU-罚款汇总-{shared_date}-另存.xlsx",
    f"(已完成-1)TEMU-罚款汇总-{shared_date}.xlsx",
)

TEMU_FINE_DIR = Path(fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\TEMU-罚款")
ORDER_STATS_DIR = Path(fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计")

ORDER_STATS_DIR = Path(fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计")
ORDER_STATS_NAME = f"(已完成-13)订单统计-{shared_date}.xlsx"
originalBakFile = f"(已完成-13)订单统计-{shared_date}-original.xlsx"
originalBakPath = ORDER_STATS_DIR / originalBakFile
# 若备份不存在，则创建备份
if not Path(originalBakPath).exists():
    pd.read_excel(ORDER_STATS_DIR / ORDER_STATS_NAME).to_excel(originalBakPath, index=False)



def _strip_df_strings(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
    return df


def _norm_key(val) -> str:
    if pd.isna(val):
        return ""
    return str(val).strip()


def _fallback_sku_site_id(shop) -> str:
    shop_s = _norm_key(shop)
    if shop_s:
        return f"{shop_s}{UNMATCHED_SUFFIX}"
    return UNMATCHED_SUFFIX.lstrip("_")


def _build_sku_platform_id(sku) -> str | None:
    """SKU-平台识别码 = 平台 + SKU。"""
    sku_s = _norm_key(sku)
    if not sku_s:
        return None
    return PLATFORM_TEMU + sku_s


def _resolve_fine_path() -> Path:
    for name in FINE_INPUT_NAMES:
        path = TEMU_FINE_DIR / name
        if path.is_file():
            return path
    raise FileNotFoundError(
        f"未找到 TEMU 罚款映射文件，请先运行 G2_2_Temu罚款映射.py。尝试："
        f"{[str(TEMU_FINE_DIR / n) for n in FINE_INPUT_NAMES]}"
    )


def _ensure_sku_site_col(fine_df: pd.DataFrame) -> pd.DataFrame:
    """统一罚款表识别列名为 SKU-站点识别（兼容旧列名）。"""
    if SKU_SITE_ID_COL in fine_df.columns:
        return fine_df
    if SKU_SITE_ID_COL_LEGACY in fine_df.columns:
        return fine_df.rename(columns={SKU_SITE_ID_COL_LEGACY: SKU_SITE_ID_COL})
    raise KeyError(
        f"罚款文件缺少列 {SKU_SITE_ID_COL!r}（或旧列名 {SKU_SITE_ID_COL_LEGACY!r}）"
    )


def _group_fine_by_sku_site(fine_df: pd.DataFrame) -> tuple[pd.DataFrame, int, float]:
    """
    按 SKU-站点识别汇总结算金额。
    无识别码时回退为「店铺__未匹配罚款」。
    返回 (grouped_df, 未匹配行数, 未匹配金额合计)
    """
    if SETTLE_AMOUNT_COL not in fine_df.columns:
        raise KeyError(f"罚款文件缺少列 {SETTLE_AMOUNT_COL!r}")

    work = fine_df.copy()
    work[SETTLE_AMOUNT_COL] = pd.to_numeric(work[SETTLE_AMOUNT_COL], errors="coerce").fillna(0)
    work[SKU_SITE_ID_COL] = work[SKU_SITE_ID_COL].map(_norm_key)

    empty_mask = work[SKU_SITE_ID_COL] == ""
    unmatched_cnt = int(empty_mask.sum())
    unmatched_amt = float(work.loc[empty_mask, SETTLE_AMOUNT_COL].sum())

    if unmatched_cnt:
        if SHOP_COL in work.columns:
            work.loc[empty_mask, SKU_SITE_ID_COL] = work.loc[empty_mask, SHOP_COL].map(
                _fallback_sku_site_id
            )
        else:
            work.loc[empty_mask, SKU_SITE_ID_COL] = _fallback_sku_site_id("")

    work = work[work[SKU_SITE_ID_COL] != ""].copy()

    agg_map: dict[str, str] = {SETTLE_AMOUNT_COL: "sum"}
    for col in (SKU_COL, SHOP_COL):
        if col in work.columns:
            agg_map[col] = "first"

    grouped = work.groupby(SKU_SITE_ID_COL, as_index=False).agg(agg_map)
    grouped[SETTLE_AMOUNT_COL] = np.round(grouped[SETTLE_AMOUNT_COL], 2)
    return grouped, unmatched_cnt, unmatched_amt


def _update_order_stats(
    order_df: pd.DataFrame, fine_grouped: pd.DataFrame
) -> tuple[pd.DataFrame, int, int, list[str], pd.DataFrame]:
    """
    将罚款结算金额累加到订单统计「其他分摊费用」。
    返回 (result_df, 更新行数, 新增行数, 重复识别码列表, 新增行明细)
    """
    if ORDER_SKU_SITE_COL not in order_df.columns:
        raise KeyError(f"订单统计缺少列 {ORDER_SKU_SITE_COL!r}，请先运行 H3")

    if ALLOC_COL not in order_df.columns:
        order_df[ALLOC_COL] = 0.0
    else:
        order_df[ALLOC_COL] = pd.to_numeric(order_df[ALLOC_COL], errors="coerce").fillna(0)

    fine_lookup = fine_grouped.set_index(SKU_SITE_ID_COL)[SETTLE_AMOUNT_COL]
    order_keys = order_df[ORDER_SKU_SITE_COL].map(_norm_key)

    dup_keys = order_keys[order_keys != ""].value_counts()
    dup_keys = dup_keys[dup_keys > 1].index.tolist()

    update_cnt = 0
    for sku_site_id, amount in fine_lookup.items():
        mask = order_keys == sku_site_id
        if not mask.any():
            continue
        if sku_site_id in dup_keys:
            first_idx = order_df.index[mask][0]
            order_df.loc[first_idx, ALLOC_COL] += amount
            update_cnt += 1
        else:
            order_df.loc[mask, ALLOC_COL] += amount
            update_cnt += int(mask.sum())

    order_df[ALLOC_COL] = np.round(pd.to_numeric(order_df[ALLOC_COL], errors="coerce").fillna(0), 2)

    existing_keys = set(order_keys[order_keys != ""])
    missing = fine_grouped[~fine_grouped[SKU_SITE_ID_COL].isin(existing_keys)].copy()
    append_cnt = len(missing)

    if append_cnt:
        new_rows = pd.DataFrame({col: np.nan for col in order_df.columns}, index=range(append_cnt))
        new_rows[ORDER_SKU_SITE_COL] = missing[SKU_SITE_ID_COL].values
        new_rows[ALLOC_COL] = missing[SETTLE_AMOUNT_COL].values
        if SKU_COL in missing.columns and SKU_COL in new_rows.columns:
            new_rows[SKU_COL] = missing[SKU_COL].values
        if SHOP_COL in missing.columns and SITE_COL in new_rows.columns:
            new_rows[SITE_COL] = missing[SHOP_COL].values
        if PLATFORM_COL in new_rows.columns:
            new_rows[PLATFORM_COL] = PLATFORM_TEMU
        if ORDER_SKU_PLATFORM_COL in new_rows.columns and SKU_COL in missing.columns:
            new_rows[ORDER_SKU_PLATFORM_COL] = missing[SKU_COL].map(_build_sku_platform_id)
        keep_cols = {
            ORDER_SKU_SITE_COL,
            ALLOC_COL,
            SKU_COL,
            SITE_COL,
            PLATFORM_COL,
            ORDER_SKU_PLATFORM_COL,
        }
        for col in order_df.columns:
            if col in new_rows.columns and col not in keep_cols:
                if order_df[col].dtype.kind in "biufc":
                    new_rows[col] = 0
        order_df = pd.concat([order_df, new_rows], ignore_index=True)

    return order_df, update_cnt, append_cnt, dup_keys, missing


def _save_excel(df: pd.DataFrame, output_path: Path) -> Path:
    try:
        df.to_excel(output_path, index=False, engine="openpyxl")
        return output_path
    except PermissionError:
        alt = output_path.with_name(output_path.stem + "-另存.xlsx")
        df.to_excel(alt, index=False, engine="openpyxl")
        print(f"{Color.YELLOW}[提示]{Color.RESET} 目标文件被占用，已另存为：{alt}")
        return alt


def main() -> None:
    fine_path = _resolve_fine_path()
    order_path = ORDER_STATS_DIR / ORDER_STATS_NAME
    if not order_path.is_file():
        raise FileNotFoundError(
            f"未找到订单统计文件：{order_path}（请先运行 H3_合并_订单统计_分摊_其他费用）"
        )

    fine_df = _ensure_sku_site_col(_strip_df_strings(pd.read_excel(fine_path)))
    order_df = _strip_df_strings(pd.read_excel(order_path))
    print(f"{Color.CYAN}[读取]{Color.RESET} 罚款 {fine_path.name}：{len(fine_df)} 行")
    print(f"{Color.CYAN}[读取]{Color.RESET} 订单统计 {order_path.name}：{len(order_df)} 行")

    fine_grouped, unmatched_cnt, unmatched_amt = _group_fine_by_sku_site(fine_df)
    print(
        f"{Color.CYAN}[汇总]{Color.RESET} {len(fine_grouped)} 个SKU-站点识别，"
        f"结算金额合计 {fine_grouped[SETTLE_AMOUNT_COL].sum():.2f} EUR"
    )
    if unmatched_cnt:
        print(
            f"{Color.YELLOW}[检查]{Color.RESET} {unmatched_cnt} 行无SKU-站点识别，"
            f"已按店铺归入「店铺{UNMATCHED_SUFFIX}」，金额 {unmatched_amt:.2f} EUR"
        )

    result_df, update_cnt, append_cnt, dup_keys, missing = _update_order_stats(
        order_df, fine_grouped
    )
    saved = _save_excel(result_df, order_path)

    print(
        f"{Color.GREEN}[更新]{Color.RESET} 已写回订单统计："
        f"累加更新 {update_cnt} 行，新增 {append_cnt} 行，合计 {len(result_df)} 行"
    )
    if dup_keys:
        print(
            f"{Color.YELLOW}[检查]{Color.RESET} 订单统计中以下SKU-站点识别码存在重复行，"
            f"罚款仅累加到首行：{', '.join(dup_keys[:5])}"
            f"{' ...' if len(dup_keys) > 5 else ''}"
        )
    if append_cnt and not missing.empty:
        cols = [
            c for c in (SKU_SITE_ID_COL, SETTLE_AMOUNT_COL, SKU_COL, SHOP_COL) if c in missing.columns
        ]
        print(f"{Color.YELLOW}[新增行预览]{Color.RESET}")
        print(missing[cols].head(5).to_string(index=False))

    print(f"处理完成，output_path：{saved}")


if __name__ == "__main__":
    main()
