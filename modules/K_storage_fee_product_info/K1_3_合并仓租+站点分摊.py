"""
K1_3_合并仓租+站点分摊.py — 合并 HY / 4PX「(平台分摊)」并按订单销量拆到站点

读取：
  …\\仓租\\4PX\\(平台分摊)4PX-仓租明细.xlsx
  …\\仓租\\鸿羽\\(平台分摊)HY-仓租明细.xlsx
  …\\订单统计\\(已完成-15)订单统计-{shared_date}.xlsx   # J1 日报 / J3 月报产出
  …\\仓租\\{ku_cun_date}库存动销明细.xlsx  Sheet「各平台SKU库存周转明细」销售站点
  DB platform_shop：仅当该商品该平台在 K0 无销售站点时使用

规则：
  1. 合并两仓 Sheet「平台分摊」：按「商品ID + 销售平台 + 运营负责人」汇总仓租
  2. 销售平台 → 订单统计「平台」（DB market_region→market_code；未命中回填原文）
  3. 站点分摊（平台分摊行全部落到站点，不因站点匹配失败进无平台）：
       严禁 A 的仓租落到 B 的店铺：
       优先 K0 该商品+平台+负责人的「销售站点」（LM-BTH 覆盖 LM-ES-BTH 等）
       该平台在 K0 有销售站点时，不再回退到该平台全部店铺
       无 K0 站点（如 AMAZON）：platform_shop 平台+ops_owner；仍无则该平台全部启用站点
       运营负责人为空 / 平台无白名单：不限站点
       L1 该商品在允许站点内有销量 → 按站点销量占比
       L2 否则按「允许站点」内平台总销量加权
       L4 否则默认主站（PLATFORM_TO_SITE）且主站 ∈ 允许站点（无白名单时不校验）
       仍未落点 → 强制落到自己的允许站点 / 默认主站 / 平台名（不挑他人店铺）
  4. 「无平台-仓租费用」= 两仓 (平台分摊) 原无平台合计（不含站点分摊差额）
  5. 写出前去掉「海外仓仓租费」为 0 的行

输出：
  …\\仓租\\(平台分摊)所有-海外仓-仓租明细.xlsx
    Sheet1 平台分摊：SKU / 商品ID / 运营负责人 / 平台 / 站点 / 识别码
                     / 海外仓仓租费 / 无平台-仓租费用
    Sheet2 无平台-仓租费用：明细 + 合计

用法：
  python modules/K_storage_fee_product_info/K1_0_HY_仓租.py
  python modules/K_storage_fee_product_info/K1_0_4PX_仓租.py
  python modules/J_amz_storage_fee/daily/J1_计算_AMZ仓租_合并_订单统计.py   # 或月报 J3
  python modules/K_storage_fee_product_info/K1_3_合并仓租+站点分摊.py
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

from common.cang_zu_decimal import round_rent, round_rent_columns, round_rent_series  # noqa: E402
from common.cang_zu_site import PLATFORM_TO_SITE  # noqa: E402
from common.platform_shop import (  # noqa: E402
    fetch_owner_platform_sites,
    fetch_platform_sites,
    map_region_to_platform,
    strip_lm_region_suffix,
)
from common.style import Color  # noqa: E402
from config.A0_paths import DESKTOP_ROOT  # noqa: E402
from config.A0_set_date import folder_name, ku_cun_date, shared_date  # noqa: E402

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")

SHEET_PLATFORM = "平台分摊"
SHEET_NO_PLATFORM = "无平台-仓租费用"
OWNER_COL = "运营负责人"
KU_CUN_SHEET = "各平台SKU库存周转明细"
KU_CUN_FILE_SUFFIX = "库存动销明细.xlsx"
QTY_COL = "可售库存-可调"
SITE_COL = "销售站点"
_REMAINDER_EPS = 1e-8
_INVALID_SITES = frozenset({"", "nan", "None", "无", "其他", "ALL"})
# 订单站点 LM-{国家}-{店铺} 中的国家段；用于把 K0「LM-BTH」对上「LM-ES-BTH」
_LM_COUNTRIES = frozenset({"ES", "FR", "IT", "PL", "PT", "DE", "UK", "BE", "NL", "US"})

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

K0_PATH = Path(
    fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租"
    fr"\{ku_cun_date}{KU_CUN_FILE_SUFFIX}"
)

OUTPUT_PATH = Path(
    fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租"
    fr"\(平台分摊)所有-海外仓-仓租明细.xlsx"
)

SHEET1_COLUMNS = [
    "SKU",
    "商品ID",
    "运营负责人",
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
    "运营负责人",
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
    if OWNER_COL not in df.columns:
        print(
            f"{Color.YELLOW}[检查] {path.name} 无列「{OWNER_COL}」，"
            f"将按空负责人处理（按平台站点分摊、不进无平台）；请重跑 K1_0_HY / K1_0_4PX{Color.RESET}"
        )
        df[OWNER_COL] = ""
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
        round_rent(
            pd.to_numeric(platform_df["无平台-仓租费用"], errors="coerce").fillna(0).sum()
        )
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


def _lm_shop_key(site: str) -> str:
    """
    抽出 LM 店铺键，去掉国家段：
      LM-ES-BTH → BTH；LM-BTH → BTH；LM-ES-BC-ls → BC-ls；LM-TOTO → TOTO
    非 LM 站点原样返回。
    """
    s = str(site).strip() if site else ""
    if not s.upper().startswith("LM-"):
        return s
    rest = s[3:]
    parts = rest.split("-", 1)
    if len(parts) == 2 and parts[0].upper() in _LM_COUNTRIES:
        return parts[1]
    return rest


def _is_lm_country_site(site: str) -> bool:
    """是否已是带国家的订单站点（LM-ES-BTH），而非 K0 粗站点（LM-BTH）。"""
    s = str(site).strip() if site else ""
    if not s.upper().startswith("LM-"):
        return False
    rest = s[3:]
    parts = rest.split("-", 1)
    return len(parts) == 2 and parts[0].upper() in _LM_COUNTRIES


def _site_allowed(site: str, allowed: frozenset[str]) -> bool:
    """
    订单站点是否落在允许集合。
    - 精确匹配；兼容订单站点 -ls/-xj 对上无后缀白名单
    - K0 粗站点（LM-BTH / LM-BC-ls）覆盖对应国家订单站点
    - 粗站点无 -ls/-xj 时，覆盖同店铺带后缀站点（LM-RP → LM-ES-RP-ls）
    带国家的 platform_shop 站点不做跨国家覆盖（LM-ES-BTH 不覆盖 LM-FR-BTH）。
    """
    s = str(site).strip() if site is not None and not pd.isna(site) else ""
    if not s or s.upper() in _INVALID_SITES:
        return False
    if s in allowed:
        return True
    s_base = strip_lm_region_suffix(s)
    if s_base and s_base in allowed:
        return True
    s_key = _lm_shop_key(s)
    s_key_base = strip_lm_region_suffix(s_key)
    for raw in allowed:
        a = str(raw).strip() if raw is not None else ""
        if not a or a.upper() in _INVALID_SITES:
            continue
        if s == a or s_base == a:
            return True
        if _is_lm_country_site(a):
            continue
        a_key = _lm_shop_key(a)
        if s_key and a_key and s_key == a_key:
            return True
        a_key_base = strip_lm_region_suffix(a_key)
        if a_key and a_key == a_key_base and s_key_base == a_key_base:
            return True
    return False


def _load_k0_inv_sites(
    path: Path,
) -> tuple[dict[tuple[str, str, str], frozenset[str]], frozenset[str]]:
    """
    K0「各平台SKU库存周转明细」→ (商品ID, 平台, 运营负责人) 的销售站点。
    仅收录可售库存-可调>0 且销售站点非空的行。
    第二个返回值：K0 里出现过销售站点的平台（这些平台禁止回退到全部店铺）。
    """
    if not path.is_file():
        raise FileNotFoundError(
            f"未找到 K0 库存动销明细：{path}\n"
            f"请先运行 modules/K_storage_fee_product_info/K0_库存周转.py"
        )
    try:
        df = pd.read_excel(path, sheet_name=KU_CUN_SHEET)
    except PermissionError as exc:
        raise PermissionError(f"文件被占用，请关闭 Excel 后再运行：{path}") from exc
    except ValueError as exc:
        raise ValueError(f"{path.name} 缺少 Sheet「{KU_CUN_SHEET}」") from exc

    need = ["商品ID", "销售平台", SITE_COL, QTY_COL]
    missing = [c for c in need if c not in df.columns]
    if missing:
        raise KeyError(f"K0「{KU_CUN_SHEET}」缺少列 {missing}：{path}")
    if OWNER_COL not in df.columns:
        df[OWNER_COL] = ""

    out = df[need + [OWNER_COL]].copy()
    out["商品ID"] = _norm_key(out["商品ID"])
    out["销售平台"] = _norm_key(out["销售平台"])
    out[OWNER_COL] = _norm_key(out[OWNER_COL])
    out[SITE_COL] = _norm_key(out[SITE_COL])
    out["平台"] = _map_sales_platform_to_platform(out["销售平台"])
    out[QTY_COL] = pd.to_numeric(out[QTY_COL], errors="coerce").fillna(0)

    plat = _norm_key(out["平台"])
    valid = (
        out["商品ID"].ne("")
        & plat.ne("")
        & ~plat.str.upper().isin(_INVALID_SITES)
        & out[SITE_COL].ne("")
        & ~out[SITE_COL].str.upper().isin(_INVALID_SITES)
        & (out[QTY_COL] > 0)
    )
    sub = out.loc[valid]
    buckets: dict[tuple[str, str, str], set[str]] = {}
    plats: set[str] = set()
    for uid, p, owner, site in zip(
        sub["商品ID"], plat.loc[sub.index], sub[OWNER_COL], sub[SITE_COL], strict=True
    ):
        key = (str(uid), str(p), str(owner))
        buckets.setdefault(key, set()).add(str(site))
        plats.add(str(p))
    return {k: frozenset(v) for k, v in buckets.items()}, frozenset(plats)


def _filter_sales_by_allowed_sites(
    sales_df: pd.DataFrame,
    allowed: frozenset[str] | None,
) -> pd.DataFrame:
    """allowed=None 表示不限制站点（运营负责人为空且平台无店铺白名单时）。"""
    if sales_df.empty:
        return sales_df.iloc[0:0].copy()
    if allowed is None:
        return sales_df.copy()
    if not allowed:
        return sales_df.iloc[0:0].copy()
    mask = sales_df["站点"].map(lambda s: _site_allowed(s, allowed))
    return sales_df.loc[mask].copy()


def _load_order_dims(
    path: Path,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    从 (已完成-15) 提取站点分摊所需维度：
      sku_site_sales  — 商品ID+平台+站点 销量（L1）
      plat_site_sales — 平台+站点 总销量（L2，再按负责人允许站点过滤）
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

    return sku_site_sales, plat_site_sales


def _split_fee_by_weights(
    rent_rows: pd.DataFrame,
    weights: pd.DataFrame,
    *,
    on_keys: list[str],
    weight_col: str = "销量",
) -> pd.DataFrame:
    """按权重把海外仓仓租费拆到站点；weights 需含 on_keys + 站点 + weight_col。"""
    out_cols = ["SKU", "商品ID", OWNER_COL, "平台", "站点", "海外仓仓租费"]
    if rent_rows.empty or weights.empty:
        return pd.DataFrame(columns=out_cols)

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
    keep = ["SKU", "商品ID", OWNER_COL, "平台", "站点"]
    for c in keep:
        if c not in merged.columns:
            merged[c] = ""
    out = merged.loc[ok, keep].copy()
    out["海外仓仓租费"] = round_rent_series(
        pd.Series(
            merged.loc[ok, "海外仓仓租费"].to_numpy()
            * merged.loc[ok, "_占比"].to_numpy(),
            index=out.index,
        )
    )
    return out[out_cols]


def _resolve_allowed_sites(
    uid: str,
    platform: str,
    owner: str,
    inv_sites: dict[tuple[str, str, str], frozenset[str]],
    inv_platforms: frozenset[str],
    owner_sites: dict[tuple[str, str], frozenset[str]],
    plat_sites: dict[str, frozenset[str]],
) -> frozenset[str] | None:
    """
    允许站点（严禁 A 落到 B 的店）：
      1. K0 该商品+平台+负责人的销售站点（优先）
      2. 该平台在 K0 有销售站点：不再回退整平台；可再用 platform_shop 平台+负责人
      3. 无 K0 站点的平台（如 AMAZON）：platform_shop；无匹配则该平台全部启用站点
      4. 负责人为空：该平台全部启用站点；平台无白名单 → None（不限制）
    """
    p = str(platform).strip() if platform is not None else ""
    o = str(owner).strip() if owner is not None else ""
    u = str(uid).strip() if uid is not None else ""
    if not p:
        return None
    k0 = inv_sites.get((u, p, o), frozenset())
    if k0:
        return k0
    shop = owner_sites.get((p, o), frozenset()) if o else frozenset()
    if p in inv_platforms:
        if shop:
            return shop
        if o:
            return frozenset()
        plat = plat_sites.get(p, frozenset())
        return plat if plat else None
    if o:
        if shop:
            return shop
        plat = plat_sites.get(p, frozenset())
        return plat if plat else None
    plat = plat_sites.get(p, frozenset())
    return plat if plat else None


def _force_site_row(row: pd.Series) -> dict:
    """站点阶梯仍无法落点时的最终兜底（不进无平台；不挑他人店铺）。"""
    plat = str(row.get("平台") or "").strip()
    sales_plat = str(row.get("销售平台") or "").strip()
    if not plat or plat.upper() in _INVALID_SITES:
        plat = sales_plat
    allowed = row.get("_allowed")
    default = _default_site_for_platform(plat) or plat or sales_plat or "未知"
    site = default
    if isinstance(allowed, (set, frozenset)):
        if allowed:
            if default and _site_allowed(default, allowed):
                site = default
            else:
                own = sorted(str(s) for s in allowed if str(s).strip())
                site = own[0] if own else (plat or sales_plat or "未知")
        else:
            site = plat or sales_plat or "未知"
    return {
        "SKU": row.get("SKU", ""),
        "商品ID": row.get("商品ID", ""),
        OWNER_COL: row.get(OWNER_COL, ""),
        "平台": plat or site,
        "站点": site,
        "海外仓仓租费": round_rent(row["海外仓仓租费"]),
    }


def _allocate_fee_by_site(
    rent_df: pd.DataFrame,
    sku_site_sales: pd.DataFrame,
    plat_site_sales: pd.DataFrame,
    inv_sites: dict[tuple[str, str, str], frozenset[str]],
    inv_platforms: frozenset[str],
    owner_sites: dict[tuple[str, str], frozenset[str]],
    plat_sites: dict[str, frozenset[str]],
) -> tuple[pd.DataFrame, float]:
    """
    阶梯：L1 商品×允许站点销量 → L2 允许站点平台销量
         → L4 默认主站 → 强制落点。
    平台分摊行全部落到站点；站点未匹配差额不进无平台。
    返回：(站点分摊明细, 强制落点金额)
    """
    empty_site = pd.DataFrame(
        columns=["SKU", "商品ID", OWNER_COL, "平台", "站点", "海外仓仓租费"]
    )
    rent = rent_df.copy()
    rent["商品ID"] = _norm_key(rent["商品ID"])
    rent["平台"] = _norm_key(rent["平台"])
    rent[OWNER_COL] = _norm_key(rent[OWNER_COL]) if OWNER_COL in rent.columns else ""
    rent["海外仓仓租费"] = round_rent_series(rent["海外仓仓租费"]).fillna(0)
    if "SKU" not in rent.columns:
        rent["SKU"] = ""
    if "销售平台" not in rent.columns:
        rent["销售平台"] = ""
    rent["销售平台"] = _norm_key(rent["销售平台"])

    # 无效/空平台：用销售平台回填，继续分摊（不进无平台）
    invalid_plat = (
        rent["平台"].eq("")
        | rent["平台"].str.upper().isin(_INVALID_SITES)
        | rent["平台"].isna()
    )
    if invalid_plat.any():
        rent.loc[invalid_plat, "平台"] = rent.loc[invalid_plat, "销售平台"]

    rent_g = (
        rent.groupby(["商品ID", "平台", OWNER_COL], as_index=False, dropna=False)
        .agg(
            SKU=("SKU", _first_nonempty),
            销售平台=("销售平台", _first_nonempty),
            海外仓仓租费=("海外仓仓租费", "sum"),
        )
    )
    rent_g["海外仓仓租费"] = round_rent_series(rent_g["海外仓仓租费"])

    site_parts: list[pd.DataFrame] = []
    if rent_g.empty:
        return empty_site, 0.0

    rent_g["_allowed"] = [
        _resolve_allowed_sites(
            uid, p, o, inv_sites, inv_platforms, owner_sites, plat_sites
        )
        for uid, p, o in zip(rent_g["商品ID"], rent_g["平台"], rent_g[OWNER_COL], strict=True)
    ]

    # —— L1：商品ID+平台，且站点 ∈ 允许站点 ——
    l1_parts: list[pd.DataFrame] = []
    l1_idx: list[int] = []
    for idx, row in rent_g.iterrows():
        allowed = row["_allowed"]
        uid, plat = str(row["商品ID"]), str(row["平台"])
        if not plat or plat.upper() in _INVALID_SITES:
            continue
        sku_w = sku_site_sales.loc[
            (sku_site_sales["商品ID"] == uid) & (sku_site_sales["平台"] == plat)
        ]
        sku_w = _filter_sales_by_allowed_sites(sku_w, allowed)
        if sku_w.empty or float(sku_w["销量"].sum()) <= 0:
            continue
        one = rent_g.loc[[idx]].drop(columns=["_allowed"])
        l1_parts.append(
            _split_fee_by_weights(one, sku_w, on_keys=["商品ID", "平台"])
        )
        l1_idx.append(idx)
    if l1_parts:
        site_parts.extend(l1_parts)
    rent_rest = rent_g.drop(index=l1_idx, errors="ignore")

    # —— L2：同平台、允许站点内的总销量 ——
    l2_parts: list[pd.DataFrame] = []
    l2_idx: list[int] = []
    for idx, row in rent_rest.iterrows():
        allowed = row["_allowed"]
        plat = str(row["平台"])
        if not plat or plat.upper() in _INVALID_SITES:
            continue
        plat_w = plat_site_sales.loc[plat_site_sales["平台"] == plat]
        plat_w = _filter_sales_by_allowed_sites(plat_w, allowed)
        if plat_w.empty or float(plat_w["销量"].sum()) <= 0:
            continue
        one = rent_rest.loc[[idx]].drop(columns=["_allowed"])
        l2_parts.append(_split_fee_by_weights(one, plat_w, on_keys=["平台"]))
        l2_idx.append(idx)
    if l2_parts:
        site_parts.extend(l2_parts)
    rent_rest = rent_rest.drop(index=l2_idx, errors="ignore")

    # —— L4：默认主站 ∈ 允许站点（allowed=None 时不校验白名单）——
    l4_rows: list[dict] = []
    l4_keep_idx: list[int] = []
    for idx, row in rent_rest.iterrows():
        allowed = row["_allowed"]
        site = _default_site_for_platform(row["平台"])
        if not site:
            continue
        if allowed is not None and not _site_allowed(site, allowed):
            continue
        l4_rows.append(
            {
                "SKU": row["SKU"],
                "商品ID": row["商品ID"],
                OWNER_COL: row[OWNER_COL],
                "平台": row["平台"],
                "站点": site,
                "海外仓仓租费": round_rent(row["海外仓仓租费"]),
            }
        )
        l4_keep_idx.append(idx)
    if l4_rows:
        site_parts.append(pd.DataFrame(l4_rows))
    rent_rest = rent_rest.drop(index=l4_keep_idx, errors="ignore")

    # —— 强制落点：剩余全部落到默认主站/平台名（不进无平台）——
    forced_total = 0.0
    if not rent_rest.empty:
        force_rows = [_force_site_row(row) for _, row in rent_rest.iterrows()]
        forced_total = float(
            round_rent(sum(float(r["海外仓仓租费"]) for r in force_rows))
        )
        site_parts.append(pd.DataFrame(force_rows))

    site_df = (
        pd.concat(site_parts, ignore_index=True)
        if site_parts
        else empty_site.copy()
    )
    if not site_df.empty:
        site_df["海外仓仓租费"] = round_rent_series(site_df["海外仓仓租费"]).fillna(0)
        if OWNER_COL not in site_df.columns:
            site_df[OWNER_COL] = ""
        site_df = (
            site_df.groupby(
                ["商品ID", OWNER_COL, "平台", "站点"], as_index=False, dropna=False
            )
            .agg(
                SKU=("SKU", _first_nonempty),
                海外仓仓租费=("海外仓仓租费", "sum"),
            )
        )
        site_df["海外仓仓租费"] = round_rent_series(site_df["海外仓仓租费"])
        site_df = site_df.sort_values(
            by=["平台", OWNER_COL, "站点", "商品ID"], kind="mergesort"
        ).reset_index(drop=True)

    return site_df, forced_total


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

    owner_sites = fetch_owner_platform_sites()
    plat_sites = fetch_platform_sites()
    inv_sites, inv_platforms = _load_k0_inv_sites(K0_PATH)
    print(
        f"[读取] platform_shop 平台×负责人站点组={len(owner_sites)} 组，"
        f"平台站点组={len(plat_sites)} 组"
    )
    print(
        f"[读取] K0 销售站点 {len(inv_sites)} 组（商品ID+平台+负责人），"
        f"有销售站点的平台={sorted(inv_platforms)} ← {K0_PATH}"
    )

    sku_site_sales, plat_site_sales = _load_order_dims(ORDER_PATH)
    print(
        f"[读取] 订单统计 L1={len(sku_site_sales)} 行，"
        f"L2平台站点={len(plat_site_sales)} 行 ← {ORDER_PATH}"
    )

    merged = pd.concat(platform_parts, ignore_index=True)
    merged["海外仓仓租费"] = round_rent_series(merged["海外仓仓租费"]).fillna(0)
    if "SKU" not in merged.columns:
        merged["SKU"] = ""
    if OWNER_COL not in merged.columns:
        merged[OWNER_COL] = ""
    merged[OWNER_COL] = _norm_key(merged[OWNER_COL])

    # 按 商品ID+销售平台+运营负责人 合并两仓
    by_sales_plat = (
        merged.groupby(["商品ID", "销售平台", OWNER_COL], as_index=False, dropna=False)
        .agg(
            SKU=("SKU", _first_nonempty),
            海外仓仓租费=("海外仓仓租费", "sum"),
        )
    )
    by_sales_plat["海外仓仓租费"] = round_rent_series(by_sales_plat["海外仓仓租费"])
    by_sales_plat["平台"] = _map_sales_platform_to_platform(by_sales_plat["销售平台"])

    rent_before = float(round_rent(by_sales_plat["海外仓仓租费"].sum()))
    site_df, forced_site_total = _allocate_fee_by_site(
        by_sales_plat,
        sku_site_sales,
        plat_site_sales,
        inv_sites,
        inv_platforms,
        owner_sites,
        plat_sites,
    )

    # 无平台仅等于两仓 (平台分摊) 原无平台合计，站点匹配失败不并入
    warehouse_no_platform = float(
        round_rent(sum(v for _, v in no_platform_totals))
    )
    all_no_platform = warehouse_no_platform

    site_df = site_df.copy()
    site_df["海外仓仓租费"] = round_rent_series(site_df["海外仓仓租费"]).fillna(0)
    before_n = len(site_df)
    site_df = site_df.loc[site_df["海外仓仓租费"].abs() > _REMAINDER_EPS].copy()
    site_df = site_df.reset_index(drop=True)
    dropped_zero = before_n - len(site_df)
    if dropped_zero:
        print(f"[清理] 已去掉海外仓仓租费=0 的行 {dropped_zero} 条")

    uid = _norm_key(site_df["商品ID"]) if not site_df.empty else pd.Series(dtype=str)
    site = _norm_key(site_df["站点"]) if not site_df.empty else pd.Series(dtype=str)
    plat = _norm_key(site_df["平台"]) if not site_df.empty else pd.Series(dtype=str)
    if OWNER_COL not in site_df.columns:
        site_df[OWNER_COL] = ""
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
                    OWNER_COL: "",
                    "平台": "",
                    "站点": "",
                    "站点商品ID识别码": "",
                    "平台商品ID识别码": "",
                    "海外仓仓租费": 0.0,
                    "无平台-仓租费用": all_no_platform,
                }
            ]
        )
    sheet1 = round_rent_columns(site_df[SHEET1_COLUMNS], ["海外仓仓租费", "无平台-仓租费用"])
    fee_sum = float(round_rent(sheet1["海外仓仓租费"].fillna(0).sum()))

    detail_parts = [
        p for p in no_platform_parts if p is not None and not p.empty
    ]
    detail = (
        pd.concat(detail_parts, ignore_index=True)
        if detail_parts
        else pd.DataFrame(columns=SHEET2_COLUMNS)
    )
    if not detail.empty and "无平台-仓租费用" in detail.columns:
        detail = round_rent_columns(detail, ["无平台-仓租费用"])

    total_row = {
        "来源": "",
        "SKU": "",
        "商品ID": "",
        "销售平台": "",
        OWNER_COL: "",
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
        f"[合并] 仓租(商品ID+销售平台+运营负责人)合计={rent_before:.4f}；"
        f"站点分摊后海外仓仓租费={fee_sum:.4f}；"
        f"强制落点={forced_site_total:.4f}"
    )
    print(
        f"[无平台] 两仓(平台分摊)原额="
        + " + ".join(f"{src} {v:.4f}" for src, v in no_platform_totals)
        + f" = {all_no_platform:.4f}（站点未匹配不并入）"
    )
    print(
        f"[核对] 站点分摊={fee_sum:.4f}（应≈平台分摊仓租={rent_before:.4f}）；"
        f"站点分摊+无平台={fee_sum + all_no_platform:.4f}"
        f"（应≈仓租原额+仓租无平台={rent_before + warehouse_no_platform:.4f}）"
    )
    print(f"{Color.GREEN}合并完成：{OUTPUT_PATH}{Color.RESET}")
    print(f"  Sheet：{SHEET_PLATFORM} / {SHEET_NO_PLATFORM}；站点行数={len(sheet1)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
