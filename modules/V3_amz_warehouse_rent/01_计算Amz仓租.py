"""计算 AMZ 仓租（功能同 J1_计算_AMZ仓租）。

数据源：``\\\\Betohow\\数据报表\\报表自动化下载\\仓租下载\\每月\\FBA仓租\\FBA仓租明细{fba_date}.xlsx``

用法::

    python modules/V3_amz_warehouse_rent/01_计算Amz仓租.py
    python modules/V3_amz_warehouse_rent/01_计算Amz仓租.py --month 2026-05
"""
from __future__ import annotations

import argparse
import importlib.util
import warnings
from calendar import monthrange
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import pymysql.cursors

# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
assert _spec.loader is not None
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.platform_shop import map_shop_platform_region  # noqa: E402
from common.style import Color  # noqa: E402
from config.A0_paths import DESKTOP_ROOT  # noqa: E402
from config.A0_set_date import (  # noqa: E402
    fba_date as default_fba_date,
    folder_name,
    report_date,
    shared_date,
)
from database.db_connection import get_db_manager  # noqa: E402

PRODUCT_SKU_TABLE = "product_sku"
PLATFORM_SHOP_TABLE = "platform_shop"
_KEY_CHUNK = 200

# 月度 FBA 仓租明细（自动化下载目录）
FBA_RENT_SOURCE_DIR = Path(
    r"\\Betohow\数据报表\报表自动化下载\仓租下载\每月\FBA仓租"
)

# 额外排除店铺（按「店铺」列精确匹配，后续直接往列表加即可）
EXCLUDE_SHOPS: list[str] = [
    "yiqianshangmao_DE",
]

_INACTIVE_SHOP_SQL = f"""
    SELECT DISTINCT TRIM(shop_name_en) AS shop_name_en
    FROM `{PLATFORM_SHOP_TABLE}`
    WHERE shop_status = 0
      AND TRIM(shop_name_en) <> ''
"""

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


def parse_month(month: str) -> tuple[int, int]:
    """解析 ``YYYY-MM``，返回 (year, month)。"""
    try:
        year_s, month_s = month.strip().split("-", 1)
        year, mon = int(year_s), int(month_s)
        if not (1 <= mon <= 12):
            raise ValueError
    except ValueError as exc:
        raise ValueError(f"无效 --month：{month!r}，期望 YYYY-MM") from exc
    return year, mon


def month_end(year: int, month: int) -> date:
    return date(year, month, monthrange(year, month)[1])


def format_fba_date_label(year: int, month: int) -> str:
    """与 A0_set_date.fba_date 一致：``m.1-m.d``（不补零）。"""
    last_day = monthrange(year, month)[1]
    return f"{month}.1-{month}.{last_day}"


def resolve_fba_snapshot_date(month: str | None = None) -> tuple[date, str]:
    """返回 (snapshot_date, fba_date_label)。

    - 指定 ``--month YYYY-MM``：snapshot_date = 该月最后一天
    - 默认：对齐 A0_set_date 的 fba_date 月份（月报往前 1 月 / 日报往前 2 月）
    """
    if month:
        year, mon = parse_month(month)
        snap = month_end(year, mon)
        return snap, format_fba_date_label(year, mon)

    months_ago = 2 if folder_name == "日报" else 1
    start_m = int(shared_date.split("-")[0].split(".")[0])
    target = date(report_date.year, start_m, 1)
    for _ in range(months_ago):
        target = (target.replace(day=1) - timedelta(days=1)).replace(day=1)
    snap = month_end(target.year, target.month)
    return snap, default_fba_date


def input_excel_path(fba_label: str) -> Path:
    return FBA_RENT_SOURCE_DIR / f"FBA仓租明细{fba_label}.xlsx"


def _detect_header_row(path: Path, *, max_scan: int = 10) -> int:
    """自动化下载的 FBA 明细前几行常为日期/币种元数据，定位含 sellerSku 的表头行。"""
    preview = pd.read_excel(path, header=None, nrows=max_scan)
    for i, row in preview.iterrows():
        vals = {str(v).strip() for v in row.tolist() if pd.notna(v)}
        if "sellerSku" in vals:
            return int(i)
    raise ValueError(f"未在前 {max_scan} 行找到表头 sellerSku：{path}")


def load_fba_excel(fba_label: str) -> pd.DataFrame:
    path = input_excel_path(fba_label)
    if not path.is_file():
        raise FileNotFoundError(
            f"未找到 FBA 仓租明细：{path}\n"
            f"请确认目录 {FBA_RENT_SOURCE_DIR} 下存在 FBA仓租明细{fba_label}.xlsx"
        )
    header_row = _detect_header_row(path)
    df = pd.read_excel(path, header=header_row)
    required = ("sellerSku", "站点", "仓库sku", "店铺", "仓储费用（已分摊）", "长期仓储费（已分摊）")
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f"FBA 仓租明细缺少列 {missing}，当前列: {list(df.columns)[:20]}...")
    print(f"[Excel] 读取 {len(df)} 行（header={header_row}）：{path}")
    return df


def extract_values(s):
    if pd.isna(s):
        return None
    if "amzn.gr." in s:
        return s.split(r"amzn.gr.")[-1].split("-")[0].split("_")[0]
    return s.split("#")[0].split("BCFBAFL")[0].split("FBFBAFL")[0]


def _fetch_product_uid_map(skus: list[str]) -> dict[str, str]:
    """从 product_sku 按 product_sku 查 product_uid（商品ID）。"""
    skus = sorted({str(x).strip() for x in skus if x and str(x).strip()})
    if not skus:
        return {}

    mapping: dict[str, str] = {}
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            for i in range(0, len(skus), _KEY_CHUNK):
                chunk = skus[i : i + _KEY_CHUNK]
                placeholders = ", ".join(["%s"] * len(chunk))
                sql = f"""
                    SELECT product_sku, product_uid
                    FROM `{PRODUCT_SKU_TABLE}`
                    WHERE product_sku IN ({placeholders})
                      AND is_deleted = 0
                      AND product_uid IS NOT NULL
                      AND TRIM(product_uid) <> ''
                """
                cur.execute(sql, chunk)
                for row in cur.fetchall():
                    sku = str(row.get("product_sku") or "").strip()
                    uid = str(row.get("product_uid") or "").strip()
                    if sku and uid:
                        mapping[sku] = uid
    finally:
        conn.close()
    return mapping


def map_sku_to_product_uid(main_df: pd.DataFrame, main_sku: str = "SKU") -> pd.DataFrame:
    """SKU → 商品ID（product_uid）；未命中置空。原 SKU 带 -NW 时，商品ID 缀回 -NW。"""
    out = main_df.copy()
    if main_sku not in out.columns:
        raise KeyError(f"主表缺少列 {main_sku!r}，当前列: {list(out.columns)}")

    series = out[main_sku].astype(str).str.strip()
    invalid = series.isin(("", "nan", "None", "NaN")) | out[main_sku].isna()
    nw_mask = series.str.endswith("-NW", na=False) & ~invalid
    series_no_nw = series.mask(nw_mask, series.str.replace(r"-NW$", "", regex=True))

    uid_map = _fetch_product_uid_map(series_no_nw[~invalid].tolist())
    print(f"[DB] product_sku 命中 {len(uid_map)} 条 product_sku → product_uid")

    mapped = series_no_nw.map(uid_map)
    mapped = mapped.mask(nw_mask & mapped.notna(), mapped.astype(str) + "-NW")
    mapped = mapped.mask(invalid, pd.NA)

    insert_pos = out.columns.get_loc(main_sku) + 1
    if "商品ID" in out.columns:
        out = out.drop(columns=["商品ID"])
    out.insert(insert_pos, "商品ID", mapped)
    return out


def _blank_mask(series: pd.Series) -> pd.Series:
    as_str = series.astype(object).map(lambda v: "" if pd.isna(v) else str(v).strip())
    return as_str.eq("") | as_str.isin(("nan", "None", "NaN"))


def warn_blank_product_uid(df: pd.DataFrame) -> None:
    if "商品ID" not in df.columns:
        return
    blank = _blank_mask(df["商品ID"])
    n = int(blank.sum())
    if n == 0:
        return
    preview_cols = [c for c in ("SKU", "sellerSku", "商品ID", "FBA仓租费") if c in df.columns]
    preview = df.loc[blank, preview_cols].head(10)
    print(
        f"{Color.YELLOW}[检查] 商品ID 有 {n} 行空值"
        f"（未映射到 product_uid），请核对：{Color.RESET}"
    )
    print(preview.to_string(index=False))


def fetch_inactive_shop_names() -> set[str]:
    """platform_shop.shop_status = 0 的停用店铺（shop_name_en）。"""
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            cur.execute(_INACTIVE_SHOP_SQL)
            rows = cur.fetchall()
    finally:
        conn.close()
    return {
        str(r.get("shop_name_en") or "").strip()
        for r in rows
        if str(r.get("shop_name_en") or "").strip()
    }


def drop_inactive_shops(df: pd.DataFrame, shop_col: str = "店铺") -> pd.DataFrame:
    """删除停用店铺（shop_status=0）及 EXCLUDE_SHOPS 列表中的店铺行。"""
    if shop_col not in df.columns:
        raise KeyError(f"主表缺少列 {shop_col!r}，当前列: {list(df.columns)}")

    inactive = fetch_inactive_shop_names()
    exclude = {str(s).strip() for s in EXCLUDE_SHOPS if str(s).strip()}
    drop_shops = inactive | exclude
    if not drop_shops:
        print("[过滤] 无停用/排除店铺，跳过")
        return df

    shops = df[shop_col].astype(object).map(
        lambda v: "" if pd.isna(v) else str(v).strip()
    )
    mask = shops.isin(drop_shops)
    n = int(mask.sum())
    hit_shops = sorted({s for s in shops[mask].tolist() if s})
    preview = ", ".join(hit_shops[:10])
    if len(hit_shops) > 10:
        preview += "…"
    print(
        f"[过滤] 删除 {n} 行"
        f"（shop_status=0: {len(inactive)} 个 / EXCLUDE_SHOPS: {len(exclude)} 个；"
        f"命中 {len(hit_shops)}"
        + (f"：{preview}" if preview else "")
        + "）"
    )
    return df.loc[~mask].copy() 


FEE_COLS = ("仓储费用（已分摊）", "长期仓储费（已分摊）")


def _shop_key(v) -> str:
    return "" if pd.isna(v) else str(v).strip()


def _sku_key(v) -> str | None:
    if pd.isna(v):
        return None
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "<na>"):
        return None
    return s


def _de_shop_candidates(shop: str) -> list[str]:
    """同店优先；若店铺名以 _UK 结尾则再试 _DE。"""
    keys = [shop] if shop else [""]
    if shop.endswith("_UK"):
        keys.append(shop[: -len("_UK")] + "_DE")
    return keys


def transfer_uk_fees_to_de(df: pd.DataFrame) -> pd.DataFrame:
    """UK 站点费用并入对应 DE 站点行，不再删除 UK 费用。

    匹配顺序（同 sellerSku）：
      1) 同店铺且站点=DE
      2) 店铺名 ``*_UK`` → ``*_DE`` 且站点=DE
      3) 任意店铺、站点=DE
    仍无匹配：将该行「站点」改为 DE 保留。
    """
    out = df.copy()
    for c in FEE_COLS:
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0.0)

    site = out["站点"].map(_shop_key)
    uk_mask = site.eq("UK")
    uk_n = int(uk_mask.sum())
    if uk_n == 0:
        print("[UK→DE] 无 UK 站点行")
        return out

    de_idx_by_sku_shop: dict[tuple[str | None, str], list] = {}
    de_idx_by_sku: dict[str | None, list] = {}
    for i, row in out.loc[site.eq("DE")].iterrows():
        sk = _sku_key(row.get("sellerSku"))
        shop = _shop_key(row.get("店铺"))
        de_idx_by_sku_shop.setdefault((sk, shop), []).append(i)
        de_idx_by_sku.setdefault(sk, []).append(i)

    drop_idx: list = []
    merged = 0
    converted = 0
    for i, row in out.loc[uk_mask].iterrows():
        sk = _sku_key(row.get("sellerSku"))
        shop = _shop_key(row.get("店铺"))
        target = None
        for cand_shop in _de_shop_candidates(shop):
            hits = de_idx_by_sku_shop.get((sk, cand_shop)) or []
            hits = [h for h in hits if h not in drop_idx]
            if hits:
                target = hits[0]
                break
        if target is None and sk is not None:
            hits = [h for h in (de_idx_by_sku.get(sk) or []) if h not in drop_idx]
            if hits:
                target = hits[0]

        if target is not None:
            for c in FEE_COLS:
                out.at[target, c] = float(out.at[target, c]) + float(out.at[i, c])
            drop_idx.append(i)
            merged += 1
        else:
            out.at[i, "站点"] = "DE"
            # 转为 DE 后可供后续同批 UK 行匹配
            de_idx_by_sku_shop.setdefault((sk, shop), []).append(i)
            de_idx_by_sku.setdefault(sk, []).append(i)
            converted += 1

    if drop_idx:
        out = out.drop(index=drop_idx)
    print(
        f"[UK→DE] UK行 {uk_n}：并入已有 DE {merged} 行，"
        f"无匹配改站点为 DE {converted} 行，剩余 {len(out)} 行"
    )
    return out.reset_index(drop=True)


def allocate_empty_sellersku_fee(df: pd.DataFrame) -> pd.DataFrame:
    """将 sellerSku 为空行的 FBA仓租费总额，等额分摊到所有非空行，再删除空行。"""
    out = df.copy()
    out["FBA仓租费"] = pd.to_numeric(out["FBA仓租费"], errors="coerce").fillna(0.0)
    empty = _blank_mask(out["sellerSku"]) | out["sellerSku"].isna()
    empty_fee = float(out.loc[empty, "FBA仓租费"].sum())
    empty_n = int(empty.sum())
    targets = out.loc[~empty].copy()
    n = len(targets)
    if empty_n == 0:
        print("[分摊] 无 sellerSku 空行")
        return targets
    if n == 0:
        print(
            f"{Color.YELLOW}[分摊] sellerSku 空行费用 {empty_fee}EUR，"
            f"但无有效行可分摊，费用丢弃{Color.RESET}"
        )
        return targets
    share = empty_fee / n
    targets["FBA仓租费"] = targets["FBA仓租费"] + share
    print(
        f"[分摊] sellerSku 空行 {empty_n} 行费用 {empty_fee:.6f}EUR "
        f"→ 等额摊到 {n} 行（每行 +{share:.6f}）"
    )
    return targets.reset_index(drop=True)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="从共享盘 FBA仓租明细 计算 AMZ 仓租（同 J1）"
    )
    parser.add_argument(
        "--month",
        default=None,
        help="归属月 YYYY-MM（默认按 A0_set_date 的 fba_date；用于定位 FBA仓租明细{fba_date}.xlsx）",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    snapshot_date, fba_label = resolve_fba_snapshot_date(args.month)
    print(f"fba_date={fba_label}, snapshot_date={snapshot_date}")

    main_file_df = load_fba_excel(fba_label)

    # UK 站点费用并入对应 DE（不再删除 UK 费用）
    main_file_df = transfer_uk_fees_to_de(main_file_df)

    # 删除 platform_shop.shop_status=0 的停用店铺行
    main_file_df = drop_inactive_shops(main_file_df, shop_col="店铺")

    # 使用 sellerSku 列的数据填充 仓库sku 列的空值
    main_file_df["仓库sku"] = main_file_df["仓库sku"].fillna(main_file_df["sellerSku"])

    # 清洗 仓库sku
    main_file_df["仓库sku"] = main_file_df["仓库sku"].apply(extract_values)
    main_file_df = main_file_df.rename(columns={"仓库sku": "SKU"})

    # 映射 商品ID：product_sku.product_sku → product_uid
    main_file_df_1 = map_sku_to_product_uid(main_file_df, main_sku="SKU")
    warn_blank_product_uid(main_file_df_1)

    # 映射站点 / 映射平台（数据源：platform_shop）
    main_file_df_3 = map_shop_platform_region(main_file_df_1, shop_col="店铺", site_col=None)

    # 在 映射站点 后插入新列 站点商品ID识别码
    new_column_data = main_file_df_3["映射站点"] + main_file_df_3["商品ID"]
    insert_position = main_file_df_3.columns.get_loc("映射站点") + 1
    main_file_df_3.insert(insert_position, "站点商品ID识别码", new_column_data)

    # 在 站点商品ID识别码后插入 平台商品ID识别码
    new_column_data = main_file_df_3["映射平台"] + main_file_df_3["商品ID"]
    insert_position = main_file_df_3.columns.get_loc("站点商品ID识别码") + 1
    main_file_df_3.insert(insert_position, "平台商品ID识别码", new_column_data)

    # FBA仓租费 = 仓储费用（已分摊）+ 长期仓储费（已分摊）
    main_file_df_3["FBA仓租费"] = (
        main_file_df_3["仓储费用（已分摊）"] + main_file_df_3["长期仓储费（已分摊）"]
    )
    main_file_df_3["FBA仓租费"] = main_file_df_3["FBA仓租费"].fillna(0).abs()

    # sellerSku 为空的 FBA仓租费等额分摊到所有有效行，再去掉空行
    empty_fee_before = float(
        main_file_df_3.loc[
            _blank_mask(main_file_df_3["sellerSku"]) | main_file_df_3["sellerSku"].isna(),
            "FBA仓租费",
        ].sum()
    )
    main_file_df_4 = allocate_empty_sellersku_fee(main_file_df_3)
    main_file_df_4 = main_file_df_4[main_file_df_4["FBA仓租费"] != 0]

    main_file_df_4 = main_file_df_4[
        [
            "sellerSku",
            "ASIN",
            "产品信息",
            "SKU",
            "商品ID",
            "店铺",
            "映射站点",
            "映射平台",
            "站点商品ID识别码",
            "平台商品ID识别码",
            "仓储费用（已分摊）",
            "长期仓储费（已分摊）",
            "FBA仓租费",
        ]
    ]

    out_dir = Path(DESKTOP_ROOT) / f"{folder_name}{shared_date}" / "仓租" / "FBA仓租"
    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = out_dir / f"(已完成-1)FBA仓租明细{fba_label}.xlsx"
    main_file_df_4.to_excel(output_path, index=False)
    print(f"处理完成，文件另存为：{output_path}")
    print(
        f"\n---------------------sellerSku为空的FBA仓租费（已分摊）是："
        f"{empty_fee_before}EUR------------------------------"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
