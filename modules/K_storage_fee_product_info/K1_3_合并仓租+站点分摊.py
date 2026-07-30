"""
K3_0_合并仓租.py — 合并 HY / 4PX「(平台分摊)」并按订单销量拆到站点

读取：
  …\\仓租\\4PX\\(平台分摊)4PX-仓租明细.xlsx
  …\\仓租\\鸿羽\\(平台分摊)HY-仓租明细.xlsx
  …\\订单统计\\(已完成-15)订单统计-{shared_date}.xlsx   # J1 日报 / J3 月报产出

规则：
  1. 合并两仓 Sheet「平台分摊」：按「商品ID + 销售平台」汇总「海外仓仓租费」
  2. 销售平台 → 订单统计「平台」（DB market_region→market_code；未命中回填原文）
  3. 站点分摊（阶梯 D，费用尽量留在该销售平台内）：
       L1 有「商品ID+平台」站点销量 → 按该商品各站点销量占比
       L2 否则按「平台内各站点总销量」加权
       L3 否则该平台下站点均摊
       L4 否则落到默认主站（PLATFORM_TO_SITE）
       L5 仍无法落点 →「无平台-仓租费用」
  4. 两仓原「无平台-仓租费用」+ L5 差额，写在结果第 1 行
  5. 写出前去掉「海外仓仓租费」为 0 的行

输出：
  …\\仓租\\(平台分摊)所有-海外仓-仓租明细.xlsx
    Sheet1 平台分摊：SKU / 商品ID / 平台 / 站点 / 站点商品ID识别码 / 平台商品ID识别码
                     / 海外仓仓租费 / 无平台-仓租费用
    Sheet2 无平台-仓租费用：明细 + 合计

用法：
  python modules/K_storage_fee_product_info/K1_0_HY_仓租.py
  python modules/K_storage_fee_product_info/K1_0_4PX_仓租.py
  python modules/J_amz_storage_fee/daily/J1_计算_AMZ仓租_合并_订单统计.py   # 或月报 J3
  python modules/K_storage_fee_product_info/K3_0_合并仓租.py
"""

from __future__ import annotations

import importlib.util
import warnings
from pathlib import Path

import pandas as pd

_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.cang_zu_site import PLATFORM_TO_SITE  # noqa: E402
from common.platform_shop import map_region_to_platform  # noqa: E402
from common.style import Color  # noqa: E402
from config.A0_paths import DESKTOP_ROOT  # noqa: E402
from config.A0_set_date import folder_name, shared_date  # noqa: E402

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")

SHEET_PLATFORM = "平台分摊"
SHEET_NO_PLATFORM = "无平台-仓租费用"
_REMAINDER_EPS = 1e-8
_INVALID_SITES = frozenset({"", "nan", "None", "无", "其他", "ALL"})

SOURCE_FILES: tuple[tuple[str, Path], ...] = (
    (
        "4PX",
        Path(
            fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\4PX"
            fr"\(平台分摊)4PX-仓租明细.xlsx"
        ),
    ),
    (
        "HY",
        Path(
            fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\鸿羽"
            fr"\(平台分摊)HY-仓租明细.xlsx"
        ),
    ),
)

ORDER_PATH = Path(
    fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计"
    fr"\(已完成-15)订单统计-{shared_date}.xlsx"
)

OUTPUT_PATH = Path(
    fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租"
    fr"\(平台分摊)所有-海外仓-仓租明细.xlsx"
)

SHEET1_COLUMNS = [
    "SKU",
    "商品ID",
    "平台",
    "站点",
    "站点商品ID识别码",
    "平台商品ID识别码",
    "海外仓仓租费",
    "无平台-仓租费用",
]
SHEET2_COLUMNS = [
    "来源",
    "SKU",
    "商品ID",
    "销售平台",
    "无平台-仓租费用",
    "原因",
]


def _is_blank(series: pd.Series) -> pd.Series:
    return series.isna() | series.astype(str).str.strip().isin(
        ["", "nan", "None", "NaN"]
    )


def _norm_key(series: pd.Series) -> pd.Series:
    return series.map(lambda v: "" if pd.isna(v) else str(v).strip())


def _to_excel_safe(path: Path, sheet1: pd.DataFrame, sheet2: pd.DataFrame) -> None:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            sheet1.to_excel(writer, sheet_name=SHEET_PLATFORM, index=False)
            sheet2.to_excel(writer, sheet_name=SHEET_NO_PLATFORM, index=False)
    except PermissionError as exc:
        raise PermissionError(f"文件被占用，请关闭文件后再运行：{path}") from exc


def _first_nonempty(series: pd.Series) -> str:
    for v in series:
        if pd.isna(v):
            continue
        s = str(v).strip()
        if s and s.lower() not in ("nan", "none"):
            return s
    return ""


def _map_sales_platform_to_platform(sales_platform: pd.Series) -> pd.Series:
    """
    销售平台 → 订单统计口径「平台」。
    DB（market_region→market_code）优先；未命中回填原文。
    """
    tmp = pd.DataFrame({"销售平台": sales_platform})
    tmp = map_region_to_platform(tmp, site_col="销售平台", platform_col="_db平台")
    mapped = tmp["_db平台"]
    raw = _norm_key(sales_platform)
    out = mapped.where(mapped.notna() & ~_is_blank(mapped), raw)
    out = out.mask(raw.str.upper().isin({"ALL", "无", "其他"}) | raw.eq(""), pd.NA)
    return out


def _read_platform_sheet(path: Path) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=SHEET_PLATFORM)
    except PermissionError as exc:
        raise PermissionError(f"文件被占用，请关闭 Excel 后再运行：{path}") from exc
    except ValueError as exc:
        if "Worksheet named" in str(exc) or "not found" in str(exc).lower():
            df = pd.read_excel(path, sheet_name=0)
        else:
            raise
    need = ["商品ID", "销售平台", "海外仓仓租费"]
    missing = [c for c in need if c not in df.columns]
    if missing:
        raise KeyError(f"{path.name} Sheet「{SHEET_PLATFORM}」缺少列 {missing}")
    return df


def _read_no_platform_sheet(path: Path, *, source: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=SHEET_NO_PLATFORM)
    except PermissionError as exc:
        raise PermissionError(f"文件被占用，请关闭 Excel 后再运行：{path}") from exc
    except ValueError:
        return pd.DataFrame(columns=SHEET2_COLUMNS)

    if df.empty:
        return pd.DataFrame(columns=SHEET2_COLUMNS)

    out = df.copy()
    out["来源"] = source
    if "原因" in out.columns:
        reason = out["原因"].astype(str).str.strip()
        out = out.loc[reason != "合计"].copy()
    for col in SHEET2_COLUMNS:
        if col not in out.columns:
            out[col] = pd.NA
    return out[SHEET2_COLUMNS]


def _extract_no_platform_total(platform_df: pd.DataFrame) -> float:
    if "无平台-仓租费用" not in platform_df.columns:
        return 0.0
    return float(
        pd.to_numeric(platform_df["无平台-仓租费用"], errors="coerce").fillna(0).sum()
    )


def _default_site_for_platform(platform: str) -> str:
    """L4：PLATFORM_TO_SITE 主站；未收录则恒等回填平台名。"""
    p = str(platform).strip() if platform is not None and not pd.isna(platform) else ""
    if not p or p.upper() in _INVALID_SITES:
        return ""
    site = str(PLATFORM_TO_SITE.get(p, p)).strip()
    if not site or site.upper() in _INVALID_SITES:
        return ""
    return site


def _load_order_dims(
    path: Path,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, list[str]]]:
    """
    从 (已完成-15) 提取站点分摊所需维度：
      sku_site_sales  — 商品ID+平台+站点 销量（L1）
      plat_site_sales — 平台+站点 总销量（L2）
      plat_sites      — 平台 → 站点列表（L3）
    """
    if not path.is_file():
        raise FileNotFoundError(
            f"未找到订单统计 (已完成-15)：{path}\n"
            f"请先运行 J1_计算_AMZ仓租_合并_订单统计.py（日报）"
            f"或 J3_合并_订单统计.py（月报）"
        )
    try:
        df = pd.read_excel(path)
    except PermissionError as exc:
        raise PermissionError(f"文件被占用，请关闭 Excel 后再运行：{path}") from exc

    need = ["商品ID", "平台", "站点", "销量"]
    missing = [c for c in need if c not in df.columns]
    if missing:
        raise KeyError(f"{path.name} 缺少列 {missing}")

    out = df[need].copy()
    out["商品ID"] = _norm_key(out["商品ID"])
    out["平台"] = _norm_key(out["平台"])
    out["站点"] = _norm_key(out["站点"])
    out["销量"] = pd.to_numeric(out["销量"], errors="coerce").fillna(0)

    valid_site = ~out["站点"].str.upper().isin(_INVALID_SITES) & out["站点"].ne("")
    valid_plat = ~out["平台"].str.upper().isin(_INVALID_SITES) & out["平台"].ne("")
    valid_uid = out["商品ID"].ne("")
    out = out.loc[valid_site & valid_plat & valid_uid].copy()

    sku_site_sales = (
        out.groupby(["商品ID", "平台", "站点"], as_index=False, dropna=False)["销量"]
        .sum()
    )
    sku_site_sales = sku_site_sales.loc[sku_site_sales["销量"] > 0].copy()

    plat_site_sales = (
        out.groupby(["平台", "站点"], as_index=False, dropna=False)["销量"].sum()
    )
    plat_site_sales = plat_site_sales.loc[plat_site_sales["销量"] > 0].copy()

    plat_sites: dict[str, list[str]] = {}
    for plat, grp in out.groupby("平台", dropna=False):
        sites = sorted({s for s in grp["站点"].tolist() if s})
        if sites:
            plat_sites[str(plat)] = sites

    return sku_site_sales, plat_site_sales, plat_sites


def _split_fee_by_weights(
    rent_rows: pd.DataFrame,
    weights: pd.DataFrame,
    *,
    on_keys: list[str],
    weight_col: str = "销量",
) -> pd.DataFrame:
    """按权重把海外仓仓租费拆到站点；weights 需含 on_keys + 站点 + weight_col。"""
    if rent_rows.empty or weights.empty:
        return pd.DataFrame(columns=["SKU", "商品ID", "平台", "站点", "海外仓仓租费"])

    tot = (
        weights.groupby(on_keys, as_index=False, dropna=False)[weight_col]
        .sum()
        .rename(columns={weight_col: "_w_sum"})
    )
    w = weights.merge(tot, on=on_keys, how="left")
    w["_占比"] = w[weight_col] / w["_w_sum"].replace(0, pd.NA)

    merged = rent_rows.merge(w[on_keys + ["站点", "_占比"]], on=on_keys, how="left")
    merged["_占比"] = pd.to_numeric(merged["_占比"], errors="coerce")
    ok = merged["站点"].notna() & merged["_占比"].notna()
    out = merged.loc[ok, ["SKU", "商品ID", "平台", "站点"]].copy()
    out["海外仓仓租费"] = (
        merged.loc[ok, "海外仓仓租费"].to_numpy() * merged.loc[ok, "_占比"].to_numpy()
    )
    return out


def _allocate_fee_by_site(
    rent_df: pd.DataFrame,
    sku_site_sales: pd.DataFrame,
    plat_site_sales: pd.DataFrame,
    plat_sites: dict[str, list[str]],
) -> tuple[pd.DataFrame, pd.DataFrame, float]:
    """
    阶梯 D：L1 商品站点销量 → L2 平台站点销量 → L3 平台站点均摊
           → L4 默认主站 → L5 无平台。
    返回：(站点分摊明细, 未匹配差额明细, 未匹配总额)
    """
    rent = rent_df.copy()
    rent["商品ID"] = _norm_key(rent["商品ID"])
    rent["平台"] = _norm_key(rent["平台"])
    rent["海外仓仓租费"] = pd.to_numeric(rent["海外仓仓租费"], errors="coerce").fillna(0)
    if "SKU" not in rent.columns:
        rent["SKU"] = ""
    if "销售平台" not in rent.columns:
        rent["销售平台"] = ""

    rent_g = (
        rent.groupby(["商品ID", "平台"], as_index=False, dropna=False)
        .agg(
            SKU=("SKU", _first_nonempty),
            销售平台=("销售平台", _first_nonempty),
            海外仓仓租费=("海外仓仓租费", "sum"),
        )
    )

    unmatched_parts: list[pd.DataFrame] = []
    site_parts: list[pd.DataFrame] = []

    # 无效平台 → L5
    invalid_plat = (
        rent_g["平台"].eq("")
        | rent_g["平台"].str.upper().isin(_INVALID_SITES)
        | rent_g["平台"].isna()
    )
    if invalid_plat.any():
        bad = rent_g.loc[invalid_plat].copy()
        bad["无平台-仓租费用"] = bad["海外仓仓租费"]
        bad["原因"] = "销售平台无法映射到订单平台"
        bad["来源"] = "站点分摊"
        unmatched_parts.append(
            bad[["来源", "SKU", "商品ID", "销售平台", "无平台-仓租费用", "原因"]]
        )
        rent_g = rent_g.loc[~invalid_plat].copy()

    if rent_g.empty:
        empty_site = pd.DataFrame(
            columns=["SKU", "商品ID", "平台", "站点", "海外仓仓租费"]
        )
        detail = (
            pd.concat(unmatched_parts, ignore_index=True)
            if unmatched_parts
            else pd.DataFrame(columns=SHEET2_COLUMNS)
        )
        total = (
            float(
                pd.to_numeric(detail["无平台-仓租费用"], errors="coerce")
                .fillna(0)
                .sum()
            )
            if not detail.empty
            else 0.0
        )
        return empty_site, detail, total

    # —— L1：商品ID+平台 有站点销量 ——
    sku_keys = (
        sku_site_sales.groupby(["商品ID", "平台"], as_index=False)["销量"]
        .sum()
        .loc[lambda d: d["销量"] > 0, ["商品ID", "平台"]]
        if not sku_site_sales.empty
        else pd.DataFrame(columns=["商品ID", "平台"])
    )
    rent_l1 = rent_g.merge(sku_keys, on=["商品ID", "平台"], how="inner")
    rent_rest = rent_g.merge(
        sku_keys.assign(_l1=1), on=["商品ID", "平台"], how="left"
    )
    rent_rest = rent_rest.loc[rent_rest["_l1"].isna()].drop(columns=["_l1"])

    if not rent_l1.empty:
        site_parts.append(
            _split_fee_by_weights(
                rent_l1, sku_site_sales, on_keys=["商品ID", "平台"]
            )
        )

    # —— L2：平台内各站点总销量 ——
    plat_keys = (
        plat_site_sales.groupby("平台", as_index=False)["销量"]
        .sum()
        .loc[lambda d: d["销量"] > 0, ["平台"]]
        if not plat_site_sales.empty
        else pd.DataFrame(columns=["平台"])
    )
    rent_l2 = rent_rest.merge(plat_keys, on="平台", how="inner")
    rent_rest = rent_rest.merge(
        plat_keys.assign(_l2=1), on="平台", how="left"
    )
    rent_rest = rent_rest.loc[rent_rest["_l2"].isna()].drop(columns=["_l2"])

    if not rent_l2.empty:
        site_parts.append(
            _split_fee_by_weights(rent_l2, plat_site_sales, on_keys=["平台"])
        )

    # —— L3：该平台下站点均摊 ——
    l3_rows: list[dict] = []
    l3_keep_idx: list[int] = []
    for idx, row in rent_rest.iterrows():
        sites = plat_sites.get(str(row["平台"]), [])
        if not sites:
            continue
        fee = float(row["海外仓仓租费"]) / len(sites)
        for site in sites:
            l3_rows.append(
                {
                    "SKU": row["SKU"],
                    "商品ID": row["商品ID"],
                    "平台": row["平台"],
                    "站点": site,
                    "海外仓仓租费": fee,
                }
            )
        l3_keep_idx.append(idx)
    if l3_rows:
        site_parts.append(pd.DataFrame(l3_rows))
    rent_rest = rent_rest.drop(index=l3_keep_idx, errors="ignore")

    # —— L4：默认主站 ——
    l4_rows: list[dict] = []
    l4_keep_idx: list[int] = []
    for idx, row in rent_rest.iterrows():
        site = _default_site_for_platform(row["平台"])
        if not site:
            continue
        l4_rows.append(
            {
                "SKU": row["SKU"],
                "商品ID": row["商品ID"],
                "平台": row["平台"],
                "站点": site,
                "海外仓仓租费": float(row["海外仓仓租费"]),
            }
        )
        l4_keep_idx.append(idx)
    if l4_rows:
        site_parts.append(pd.DataFrame(l4_rows))
    rent_rest = rent_rest.drop(index=l4_keep_idx, errors="ignore")

    # —— L5：仍无法落点 ——
    if not rent_rest.empty:
        miss = rent_rest.copy()
        miss["无平台-仓租费用"] = miss["海外仓仓租费"]
        miss["原因"] = "无法落点到站点(无订单站点且无默认主站)"
        miss["来源"] = "站点分摊"
        unmatched_parts.append(
            miss[["来源", "SKU", "商品ID", "销售平台", "无平台-仓租费用", "原因"]]
        )

    site_df = (
        pd.concat(site_parts, ignore_index=True)
        if site_parts
        else pd.DataFrame(columns=["SKU", "商品ID", "平台", "站点", "海外仓仓租费"])
    )
    if not site_df.empty:
        site_df["海外仓仓租费"] = pd.to_numeric(
            site_df["海外仓仓租费"], errors="coerce"
        ).fillna(0)
        site_df = (
            site_df.groupby(["商品ID", "平台", "站点"], as_index=False, dropna=False)
            .agg(
                SKU=("SKU", _first_nonempty),
                海外仓仓租费=("海外仓仓租费", "sum"),
            )
        )
        site_df = site_df.sort_values(
            by=["平台", "站点", "商品ID"], kind="mergesort"
        ).reset_index(drop=True)

    unmatched_detail = (
        pd.concat(unmatched_parts, ignore_index=True)
        if unmatched_parts
        else pd.DataFrame(columns=SHEET2_COLUMNS)
    )
    unmatched_total = (
        float(
            pd.to_numeric(unmatched_detail["无平台-仓租费用"], errors="coerce")
            .fillna(0)
            .sum()
        )
        if not unmatched_detail.empty
        else 0.0
    )
    return site_df, unmatched_detail, unmatched_total


def main() -> int:
    platform_parts: list[pd.DataFrame] = []
    no_platform_parts: list[pd.DataFrame] = []
    no_platform_totals: list[tuple[str, float]] = []

    for source, path in SOURCE_FILES:
        if not path.is_file():
            print(f"{Color.YELLOW}[跳过] 未找到 {source} 平台分摊文件：{path}{Color.RESET}")
            continue
        plat_df = _read_platform_sheet(path)
        no_total = _extract_no_platform_total(plat_df)
        no_platform_totals.append((source, no_total))
        platform_parts.append(plat_df)
        no_platform_parts.append(_read_no_platform_sheet(path, source=source))
        print(
            f"[读取] {source}：平台分摊 {len(plat_df)} 行，"
            f"无平台-仓租费用={no_total:.4f} ← {path}"
        )

    if not platform_parts:
        raise FileNotFoundError(
            "未找到任何 (平台分摊) 文件，请先运行 K1_0_HY_仓租.py / K1_0_4PX_仓租.py"
        )

    sku_site_sales, plat_site_sales, plat_sites = _load_order_dims(ORDER_PATH)
    print(
        f"[读取] 订单统计 L1={len(sku_site_sales)} 行，"
        f"L2平台站点={len(plat_site_sales)} 行，"
        f"L3平台数={len(plat_sites)} ← {ORDER_PATH}"
    )

    merged = pd.concat(platform_parts, ignore_index=True)
    merged["海外仓仓租费"] = pd.to_numeric(merged["海外仓仓租费"], errors="coerce").fillna(0)
    if "SKU" not in merged.columns:
        merged["SKU"] = ""

    # 先按 商品ID+销售平台 合并两仓
    by_sales_plat = (
        merged.groupby(["商品ID", "销售平台"], as_index=False, dropna=False)
        .agg(
            SKU=("SKU", _first_nonempty),
            海外仓仓租费=("海外仓仓租费", "sum"),
        )
    )
    by_sales_plat["平台"] = _map_sales_platform_to_platform(by_sales_plat["销售平台"])

    rent_before = float(by_sales_plat["海外仓仓租费"].sum())
    site_df, site_unmatched_detail, site_unmatched_total = _allocate_fee_by_site(
        by_sales_plat, sku_site_sales, plat_site_sales, plat_sites
    )

    warehouse_no_platform = float(sum(v for _, v in no_platform_totals))
    all_no_platform = warehouse_no_platform + site_unmatched_total

    site_df = site_df.copy()
    site_df["海外仓仓租费"] = pd.to_numeric(
        site_df["海外仓仓租费"], errors="coerce"
    ).fillna(0)
    # 去掉仓租为 0 的行（如库存标签映射出 chengyi-CD 但无实际分摊费用）
    before_n = len(site_df)
    site_df = site_df.loc[site_df["海外仓仓租费"].abs() > _REMAINDER_EPS].copy()
    site_df = site_df.reset_index(drop=True)
    dropped_zero = before_n - len(site_df)
    if dropped_zero:
        print(f"[清理] 已去掉海外仓仓租费=0 的行 {dropped_zero} 条")

    uid = _norm_key(site_df["商品ID"]) if not site_df.empty else pd.Series(dtype=str)
    site = _norm_key(site_df["站点"]) if not site_df.empty else pd.Series(dtype=str)
    plat = _norm_key(site_df["平台"]) if not site_df.empty else pd.Series(dtype=str)
    # 与旧 K1/K2 一致：站点+商品ID / 平台+商品ID；任一侧为空则识别码置空
    if not site_df.empty:
        site_df["站点商品ID识别码"] = (site + uid).where(
            site.ne("") & uid.ne(""), ""
        )
        site_df["平台商品ID识别码"] = (plat + uid).where(
            plat.ne("") & uid.ne(""), ""
        )
    site_df["无平台-仓租费用"] = None
    if len(site_df) > 0:
        site_df.at[0, "无平台-仓租费用"] = all_no_platform
    else:
        site_df = pd.DataFrame(
            [
                {
                    "SKU": "",
                    "商品ID": "",
                    "平台": "",
                    "站点": "",
                    "站点商品ID识别码": "",
                    "平台商品ID识别码": "",
                    "海外仓仓租费": 0.0,
                    "无平台-仓租费用": all_no_platform,
                }
            ]
        )
    sheet1 = site_df[SHEET1_COLUMNS]
    fee_sum = float(pd.to_numeric(sheet1["海外仓仓租费"], errors="coerce").fillna(0).sum())

    # Sheet2：原无平台明细 + 站点分摊未匹配 + 合计
    detail_parts = [
        p for p in no_platform_parts if p is not None and not p.empty
    ]
    if not site_unmatched_detail.empty:
        detail_parts.append(site_unmatched_detail.reindex(columns=SHEET2_COLUMNS))
    detail = (
        pd.concat(detail_parts, ignore_index=True)
        if detail_parts
        else pd.DataFrame(columns=SHEET2_COLUMNS)
    )
    if not detail.empty and "无平台-仓租费用" in detail.columns:
        detail["无平台-仓租费用"] = pd.to_numeric(
            detail["无平台-仓租费用"], errors="coerce"
        )

    total_row = {
        "来源": "",
        "SKU": "",
        "商品ID": "",
        "销售平台": "",
        "无平台-仓租费用": all_no_platform,
        "原因": "合计",
    }
    sheet2 = pd.concat(
        [
            pd.DataFrame([total_row], columns=SHEET2_COLUMNS),
            detail.reindex(columns=SHEET2_COLUMNS),
        ],
        ignore_index=True,
    )

    _to_excel_safe(OUTPUT_PATH, sheet1, sheet2)

    print(
        f"[合并] 仓租(商品ID+销售平台)合计={rent_before:.4f}；"
        f"站点分摊后海外仓仓租费={fee_sum:.4f}；"
        f"站点未匹配={site_unmatched_total:.4f}"
    )
    print(
        f"[无平台] 仓租原额="
        + " + ".join(f"{src} {v:.4f}" for src, v in no_platform_totals)
        + f" = {warehouse_no_platform:.4f}；"
        f"加站点未匹配后={all_no_platform:.4f}"
    )
    print(
        f"[核对] 站点分摊+无平台="
        f"{fee_sum + all_no_platform:.4f}（应≈仓租原额+仓租无平台"
        f"={rent_before + warehouse_no_platform:.4f}）"
    )
    print(f"{Color.GREEN}合并完成：{OUTPUT_PATH}{Color.RESET}")
    print(f"  Sheet：{SHEET_PLATFORM} / {SHEET_NO_PLATFORM}；站点行数={len(sheet1)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
