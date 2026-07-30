"""
K1_0_4PX_仓租.py — 4PX 仓租分摊（DB / K0 版）

对应 K2_4PX_仓租.py 的改版，映射/库存逻辑对齐 K1_0_HY_仓租.py：

  1. 读 4PX 海外仓仓租；忽略「应收金额」为 0 的行，按 SKU 汇总 → 总仓租
  2. product_sku_mapping（4PX / warehouse）：warehouse_sku → product_sku
     - 未命中：插入待更新行（is_active=0），提示人工补全后重跑（可先跑 upProductMapping）
  3. product_sku → product_uid（商品ID）
  4. 用 K0「各平台商品ID周转明细」，按「销售平台 + 商品ID + 可售库存-可调」占比
     分摊「海外仓仓租费」到销售平台（站点级分摊后续再做）

用法：
  python modules/K_storage_fee_product_info/K0_库存周转.py
  python modules/K_storage_fee_product_info/upProductMapping.py   # 可选
  python modules/K_storage_fee_product_info/K1_0_4PX_仓租.py
"""

from __future__ import annotations

import hashlib
import importlib.util
import os
import warnings
from pathlib import Path

import pandas as pd
import pymysql.cursors

_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.style import Color  # noqa: E402
from config.A0_paths import DESKTOP_ROOT  # noqa: E402
from config.A0_set_date import folder_name, ku_cun_date, shared_date  # noqa: E402
from database.db_connection import get_db_manager  # noqa: E402

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")

PRODUCT_SKU_TABLE = "product_sku"
PSM_TABLE = "product_sku_mapping"
PARTNER_CODE_4PX = "4PX"
PARTNER_TYPE_WH = "warehouse"
PARTNER_NAME_4PX = "4PX海外仓"
# 4PX 仓租表无仓库代码列：source_type 固定用此值
DEFAULT_SOURCE_TYPE = "4PX"
# K0 sheet2：销售平台 + 商品ID 汇总
KU_CUN_SHEET = "各平台商品ID周转明细"
KU_CUN_FILE_SUFFIX = "库存动销明细.xlsx"
PLATFORM_COL = "销售平台"
QTY_COL = "可售库存-可调"
# 不参与有效分摊的销售平台：库存不进分母，仓租不进 Sheet1；
# 仅当商品无其它有效平台库存时，整笔进入 Sheet2「无平台-仓租费用」
NO_SITE_PLATFORMS = frozenset({"无", "其他", "ALL"})
_REMAINDER_EPS = 1e-8
SHEET_PLATFORM = "平台分摊"
SHEET_NO_PLATFORM = "无平台-仓租费用"
PX4_SHEET1_COLUMNS = [
    "SKU",
    "商品ID",
    "销售平台",
    "可售库存-可调",
    "海外仓仓租费",
    "无平台-仓租费用",
]
PX4_SHEET2_COLUMNS = [
    "SKU",
    "商品ID",
    "销售平台",
    "可售库存-可调",
    "无平台-仓租费用",
    "原因",
]
_SHEET2_WAREHOUSE_KEY_CANDIDATES = ("仓库SKU", "产品代码（SKU）")
_KEY_CHUNK = 500
_SOURCE_TYPE_MAX_LEN = 24

_PSM_HASH_FIELDS = (
    "partner_code",
    "partner_type",
    "shop_hash",
    "seller_sku",
    "warehouse_sku",
)


def _psm_line_hash(
    *,
    warehouse_sku: str,
    product_sku: str,
    seller_sku: str = "",
    mapping_type: str = "single",
) -> str:
    record = {
        "partner_code": PARTNER_CODE_4PX,
        "partner_type": PARTNER_TYPE_WH,
        "partner_name": PARTNER_NAME_4PX,
        "shop_hash": "",
        "seller_sku": seller_sku,
        "warehouse_sku": warehouse_sku,
        "mapping_type": mapping_type,
        "product_sku": product_sku,
    }
    parts = [str(record.get(k) or "").strip() for k in _PSM_HASH_FIELDS]
    return hashlib.sha256("\x1f".join(parts).encode("utf-8")).hexdigest()


def _fetch_4px_warehouse_sku_map(warehouse_skus: list[str]) -> dict[str, str]:
    """有效映射：4PX/warehouse，is_active=1 且 product_sku 非空。"""
    warehouse_skus = sorted(
        {str(x).strip() for x in warehouse_skus if x and str(x).strip()}
    )
    if not warehouse_skus:
        return {}

    mapping: dict[str, str] = {}
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            for i in range(0, len(warehouse_skus), _KEY_CHUNK):
                chunk = warehouse_skus[i : i + _KEY_CHUNK]
                placeholders = ", ".join(["%s"] * len(chunk))
                sql = f"""
                    SELECT warehouse_sku, product_sku, updated_at
                    FROM `{PSM_TABLE}`
                    WHERE partner_code = %s
                      AND partner_type = %s
                      AND mapping_type = 'single'
                      AND is_active = 1
                      AND warehouse_sku IN ({placeholders})
                      AND product_sku IS NOT NULL
                      AND TRIM(product_sku) <> ''
                    ORDER BY updated_at DESC
                """
                cur.execute(sql, (PARTNER_CODE_4PX, PARTNER_TYPE_WH, *chunk))
                for row in cur.fetchall():
                    wh = str(row.get("warehouse_sku") or "").strip()
                    sku = str(row.get("product_sku") or "").strip()
                    if wh and sku and wh not in mapping:
                        mapping[wh] = sku
    finally:
        conn.close()
    return mapping


def _fetch_existing_4px_warehouse_skus(warehouse_skus: list[str]) -> set[str]:
    warehouse_skus = sorted(
        {str(x).strip() for x in warehouse_skus if x and str(x).strip()}
    )
    if not warehouse_skus:
        return set()

    existing: set[str] = set()
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            for i in range(0, len(warehouse_skus), _KEY_CHUNK):
                chunk = warehouse_skus[i : i + _KEY_CHUNK]
                placeholders = ", ".join(["%s"] * len(chunk))
                sql = f"""
                    SELECT DISTINCT warehouse_sku
                    FROM `{PSM_TABLE}`
                    WHERE partner_code = %s
                      AND partner_type = %s
                      AND warehouse_sku IN ({placeholders})
                """
                cur.execute(sql, (PARTNER_CODE_4PX, PARTNER_TYPE_WH, *chunk))
                for row in cur.fetchall():
                    wh = str(row.get("warehouse_sku") or "").strip()
                    if wh:
                        existing.add(wh)
    finally:
        conn.close()
    return existing


def _fetch_product_sku_uid_map(skus: list[str]) -> dict[str, str]:
    """product_sku → product_uid（is_deleted=0）；无 uid 时值为 ''。"""
    skus = sorted({str(x).strip() for x in skus if x and str(x).strip()})
    if not skus:
        return {}

    found: dict[str, str] = {}
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
                """
                cur.execute(sql, chunk)
                for row in cur.fetchall():
                    sku = str(row.get("product_sku") or "").strip()
                    if sku:
                        found[sku] = str(row.get("product_uid") or "").strip()
    finally:
        conn.close()
    return found


def _insert_pending_4px_mappings(
    warehouse_skus: list[str],
) -> tuple[int, int, dict[str, str]]:
    """
    写入 product_sku_mapping（4PX/warehouse）。
    命中 product_sku 表 → 填 product_sku / seller_ean，is_active=1；
    未命中 → product_sku 空，is_active=0。
    source_type 固定 DEFAULT_SOURCE_TYPE；seller_sku 空串。
    """
    skus = sorted({str(x).strip() for x in warehouse_skus if x and str(x).strip()})
    if not skus:
        return 0, 0, {}

    sku_uid_map = _fetch_product_sku_uid_map(skus)
    source_type = DEFAULT_SOURCE_TYPE[:_SOURCE_TYPE_MAX_LEN]

    sql = f"""
        INSERT INTO `{PSM_TABLE}` (
            line_hash, partner_code, partner_type, partner_name, shop_hash,
            seller_ean, seller_sku, warehouse_sku, mapping_type, product_sku,
            component_info, source_type, is_active
        ) VALUES (
            %s, %s, %s, %s, '',
            %s, '', %s, %s, %s,
            %s, %s, %s
        )
    """
    n_auto = 0
    n_pending = 0
    auto_map: dict[str, str] = {}
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor() as cur:
            for wh in skus:
                if wh in sku_uid_map:
                    mapping_type = "single"
                    product_sku = wh
                    seller_ean = sku_uid_map[wh] or ""
                    component_info = None
                    is_active = 1
                    auto_map[wh] = wh
                else:
                    mapping_type = "single"
                    product_sku = ""
                    seller_ean = ""
                    component_info = None
                    is_active = 0

                line_hash = _psm_line_hash(
                    warehouse_sku=wh,
                    product_sku=product_sku,
                    seller_sku="",
                    mapping_type=mapping_type,
                )
                try:
                    cur.execute(
                        sql,
                        (
                            line_hash,
                            PARTNER_CODE_4PX,
                            PARTNER_TYPE_WH,
                            PARTNER_NAME_4PX,
                            seller_ean,
                            wh,
                            mapping_type,
                            product_sku,
                            component_info,
                            source_type,
                            is_active,
                        ),
                    )
                    if cur.rowcount:
                        if is_active:
                            n_auto += 1
                        else:
                            n_pending += 1
                except (pymysql.err.IntegrityError, pymysql.err.OperationalError) as exc:
                    print(
                        f"{Color.RED}[写入失败] warehouse_sku={wh} "
                        f"product_sku={product_sku!r} → {exc}{Color.RESET}"
                    )
                    err = str(exc)
                    if "chk_psm_mapping_payload" in err:
                        print(
                            f"{Color.YELLOW}[提示] 请先执行 "
                            f"database/alter/alter_product_sku_mapping_allow_empty_product_sku.sql{Color.RESET}"
                        )
        conn.commit()
    finally:
        conn.close()
    return n_auto, n_pending, auto_map


def map_4px_warehouse_to_product_sku(
    main_df: pd.DataFrame, main_sku: str = "仓库SKU"
) -> tuple[pd.DataFrame, list[str]]:
    """product_sku_mapping(4PX/warehouse) → 标准 SKU；未命中 SKU 置空。"""
    out = main_df.copy()
    if main_sku not in out.columns:
        raise KeyError(f"主表缺少列 {main_sku!r}，当前列: {list(out.columns)}")

    series = out[main_sku].astype(str).str.strip()
    invalid = series.isin(("", "nan", "None", "NaN")) | out[main_sku].isna()

    sku_map = _fetch_4px_warehouse_sku_map(series[~invalid].tolist())
    print(
        f"[DB] product_sku_mapping(4PX/warehouse) 命中 "
        f"{len(sku_map)} 条 warehouse_sku → product_sku"
    )

    mapped = series.map(sku_map)
    mapped = mapped.mask(invalid, pd.NA)
    missing = sorted(
        {
            str(s).strip()
            for s in series[~invalid]
            if str(s).strip() and str(s).strip() not in sku_map
        }
    )

    insert_pos = out.columns.get_loc(main_sku) + 1
    if "SKU" in out.columns:
        out = out.drop(columns=["SKU"])
    out.insert(insert_pos, "SKU", mapped)
    return out, missing


def _apply_sku_map(
    df: pd.DataFrame, sku_map: dict[str, str], main_sku: str = "仓库SKU"
) -> pd.DataFrame:
    if not sku_map or "SKU" not in df.columns or main_sku not in df.columns:
        return df
    out = df.copy()
    key = out[main_sku].astype(str).str.strip()
    fill = key.map(sku_map)
    sku_str = out["SKU"].astype(object).map(lambda v: "" if pd.isna(v) else str(v).strip())
    blank = sku_str.eq("") | sku_str.isin(("nan", "None", "NaN"))
    out.loc[blank & fill.notna(), "SKU"] = fill[blank & fill.notna()]
    return out


def _fetch_product_uid_map(skus: list[str]) -> dict[str, str]:
    """仅返回有非空 product_uid 的映射。"""
    raw = _fetch_product_sku_uid_map(skus)
    return {k: v for k, v in raw.items() if v}


def map_sku_to_product_uid(main_df: pd.DataFrame, main_sku: str = "SKU") -> pd.DataFrame:
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
    preview_cols = [c for c in ("仓库SKU", "SKU", "商品ID", "总仓租") if c in df.columns]
    preview = df.loc[blank, preview_cols].head(10)
    print(
        f"{Color.YELLOW}[检查] 商品ID 有 {n} 行空值"
        f"（未映射到 product_uid），请核对：{Color.RESET}"
    )
    print(preview.to_string(index=False))


def _to_excel_safe(df: pd.DataFrame, path: str, **kwargs) -> None:
    try:
        df.to_excel(path, index=False, **kwargs)
    except PermissionError as exc:
        raise PermissionError(f"文件被占用，请关闭文件后再运行：{path}") from exc


def _prepare_platform_qty_ratio(ku_cun_df: pd.DataFrame) -> pd.DataFrame:
    """
    按商品ID 计算有效销售平台的库存占比。
    「无」「其他」「ALL」库存不进分母，避免稀释有效平台分摊。
    有效库存合计为 0 时占比为 NaN（整笔进无平台）。
    """
    out = ku_cun_df.copy()
    out[QTY_COL] = pd.to_numeric(out[QTY_COL], errors="coerce").fillna(0)
    out = (
        out.groupby(["商品ID", PLATFORM_COL], as_index=False, dropna=False)[QTY_COL]
        .sum()
    )
    is_no_site = out[PLATFORM_COL].isin(NO_SITE_PLATFORMS)
    out["_有效库存"] = out[QTY_COL].where(~is_no_site, 0.0)
    effective_total = out.groupby("商品ID")["_有效库存"].transform("sum")
    out["数量占比"] = out["_有效库存"] / effective_total.replace(0, pd.NA)
    out.loc[is_no_site, "数量占比"] = 0.0
    return out.drop(columns=["_有效库存"])


def _remainder_group_key(df: pd.DataFrame) -> str:
    """差额按「仓租账单行」汇总，避免多条仓库SKU映射到同一SKU时总仓租只取首行。"""
    for col in _SHEET2_WAREHOUSE_KEY_CANDIDATES:
        if col in df.columns:
            return col
    return "SKU"


def _sku_no_platform_reason(
    platforms: pd.Series,
    *,
    rem: float,
    allocated: float,
) -> str:
    plats = {
        str(p).strip()
        for p in platforms.dropna().tolist()
        if str(p).strip() and str(p).strip().lower() not in ("nan", "none")
    }
    if rem < -_REMAINDER_EPS:
        return "分摊超额(请检查重复映射)"
    if not plats:
        return "无库存匹配"
    if plats <= NO_SITE_PLATFORMS:
        return "仅无/其他/ALL库存"
    if allocated <= _REMAINDER_EPS:
        return "有效平台可售库存合计为0"
    return "部分未分摊"


def _build_no_platform_detail(df: pd.DataFrame) -> pd.DataFrame:
    """
    按仓租账单行（仓库SKU）汇总「总仓租 − 已分摊到有效平台」。
    不可按映射后 SKU 汇总：多条仓库SKU 共一个 SKU 时会把多笔仓租的已分摊
    扣在单笔总仓租上，产生负数。
    """
    empty = pd.DataFrame(columns=PX4_SHEET2_COLUMNS)
    if df.empty:
        return empty

    work = df.copy()
    work["_仓租"] = pd.to_numeric(work.get("仓租"), errors="coerce").fillna(0)
    work["_总仓租"] = pd.to_numeric(work["总仓租"], errors="coerce").fillna(0)
    work[QTY_COL] = pd.to_numeric(work.get(QTY_COL), errors="coerce").fillna(0)
    key_col = _remainder_group_key(work)

    rows: list[dict] = []
    for key, g in work.groupby(key_col, dropna=False, sort=False):
        total = float(g["_总仓租"].iloc[0])
        allocated = float(g["_仓租"].sum())
        rem = total - allocated
        if abs(rem) <= _REMAINDER_EPS:
            continue
        no_site_qty = float(
            g.loc[g[PLATFORM_COL].isin(NO_SITE_PLATFORMS), QTY_COL].sum()
        )
        no_site_plats = sorted(
            {
                str(p).strip()
                for p in g[PLATFORM_COL].dropna()
                if str(p).strip() in NO_SITE_PLATFORMS
            }
        )
        head = g.iloc[0]
        show_sku = head["SKU"] if "SKU" in g.columns else key
        rows.append(
            {
                "SKU": show_sku,
                "商品ID": head.get("商品ID"),
                PLATFORM_COL: ",".join(no_site_plats),
                QTY_COL: no_site_qty
                if no_site_plats
                else float(g[QTY_COL].fillna(0).sum()),
                "无平台-仓租费用": rem,
                "原因": _sku_no_platform_reason(
                    g[PLATFORM_COL], rem=rem, allocated=allocated
                ),
            }
        )

    if not rows:
        return empty
    return pd.DataFrame(rows)[PX4_SHEET2_COLUMNS]


def _with_no_platform_total(
    detail: pd.DataFrame,
    total: float,
    columns: list[str],
) -> pd.DataFrame:
    """Sheet2 首行写合计；避免空表/全 NA 列 concat 触发 FutureWarning。"""
    head = pd.DataFrame(
        [
            {
                "SKU": "",
                "商品ID": "",
                PLATFORM_COL: "",
                QTY_COL: 0,
                "无平台-仓租费用": total,
                "原因": "合计",
            }
        ],
        columns=columns,
    )
    if detail is None or detail.empty:
        return head
    body = detail.reindex(columns=columns)
    return pd.concat([head, body], ignore_index=True)


def _save_alloc_workbook(
    sheet1: pd.DataFrame,
    sheet2: pd.DataFrame,
    path: str,
) -> None:
    """Sheet1=平台分摊，Sheet2=无平台-仓租费用明细。"""
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            sheet1.to_excel(writer, sheet_name=SHEET_PLATFORM, index=False)
            sheet2.to_excel(writer, sheet_name=SHEET_NO_PLATFORM, index=False)
    except PermissionError as exc:
        raise PermissionError(f"文件被占用，请关闭文件后再运行：{path}") from exc


def _save_step1(df: pd.DataFrame, out_dir: str) -> str:
    out = df
    if "总仓租" in df.columns:
        fee = pd.to_numeric(df["总仓租"], errors="coerce").fillna(0)
        out = df.loc[fee != 0]
    cols = [c for c in ("仓库SKU", "SKU", "商品ID", "总仓租") if c in out.columns]
    cols += [c for c in out.columns if c not in cols]
    path = os.path.join(out_dir, "(商品ID)4PX-仓租明细.xlsx")
    _to_excel_safe(out[cols], path)
    return path


def _save_empty_final(out_dir: str) -> str:
    path = os.path.join(out_dir, "(平台分摊)4PX-仓租明细.xlsx")
    _save_alloc_workbook(
        pd.DataFrame(columns=PX4_SHEET1_COLUMNS),
        pd.DataFrame(columns=PX4_SHEET2_COLUMNS),
        path,
    )
    return path

# ------------------------------------------------------------------------------------------#
# ------------------------------------------------------------------------------------------#
# ------------------------------------------------------------------------------------------#
# ---------- 1. 读取并汇总 4PX 仓租 ----------
px4_path = (
    fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\4PX"
    fr"\4PX法国仓-仓租明细-{shared_date}.xlsx"
)
output_dir = os.path.dirname(px4_path)

if not os.path.isfile(px4_path):
    raise FileNotFoundError(f"未找到 4PX 仓租文件：{px4_path}")

try:
    px4_df = pd.read_excel(px4_path)
except PermissionError as exc:
    raise PermissionError(f"文件被占用，请关闭 Excel 后再运行：{px4_path}") from exc

row_count = len(px4_df)
if row_count == 0:
    print(f"提示：4PX 仓租表无数据行，继续后续流程：{px4_path}")
    px4_df = pd.DataFrame(columns=["SKU", "应收金额"])
elif row_count < 2:
    print(f"提示：4PX 仓租表仅 {row_count} 行，继续处理：{px4_path}")

AMOUNT_COL = "应收金额"
if px4_df.empty:
    px4_df = pd.DataFrame(columns=["SKU", AMOUNT_COL])
else:
    missing_cols = [c for c in ("SKU", AMOUNT_COL) if c not in px4_df.columns]
    if missing_cols:
        raise ValueError(f"4PX 仓租表缺少列 {missing_cols}：{px4_path}")
    px4_df["SKU"] = px4_df["SKU"].astype(str).str.strip()
    # 应收金额为 0（含空值）不影响总仓租，跳过以免无费用 SKU 触发映射待更新
    px4_df[AMOUNT_COL] = pd.to_numeric(px4_df[AMOUNT_COL], errors="coerce").fillna(0)
    _n_before = len(px4_df)
    px4_df = px4_df.loc[px4_df[AMOUNT_COL] != 0].copy()
    _n_skipped = _n_before - len(px4_df)
    if _n_skipped:
        print(f"[过滤] 忽略 {AMOUNT_COL}=0 的行：{_n_skipped} 条，剩余 {len(px4_df)} 条")
    if not px4_df.empty:
        px4_df = px4_df.groupby("SKU", as_index=False)[AMOUNT_COL].sum()
        # 汇总后仍为 0 的（正负相抵）同样忽略
        px4_df = px4_df.loc[
            pd.to_numeric(px4_df[AMOUNT_COL], errors="coerce").fillna(0) != 0
        ].copy()

px4_df = px4_df.rename(columns={AMOUNT_COL: "总仓租", "SKU": "仓库SKU"})
px4_all_cang_zu = float(pd.to_numeric(px4_df["总仓租"], errors="coerce").fillna(0).sum())

# ---------- 2. product_sku_mapping：仓库 SKU → 标准 SKU ----------
if px4_df.empty:
    print(f"{Color.RED}提示：4PX 无 SKU 明细，生成分摊结果空表并结束{Color.RESET}")
    out_path = _save_empty_final(output_dir)
    print(f"平台分摊，结果已保存到{out_path}")
    raise SystemExit(0)

px4_df_1, missing_wh = map_4px_warehouse_to_product_sku(px4_df, main_sku="仓库SKU")

if missing_wh:
    already = _fetch_existing_4px_warehouse_skus(missing_wh)
    to_insert = [s for s in missing_wh if s not in already]
    pending_only = [s for s in missing_wh if s in already]

    n_auto, n_pending, auto_map = (
        _insert_pending_4px_mappings(to_insert) if to_insert else (0, 0, {})
    )
    if auto_map:
        px4_df_1 = _apply_sku_map(px4_df_1, auto_map, main_sku="仓库SKU")
        print(
            f"{Color.GREEN}[DB] 已用 product_sku 表自动写入 mapping 并回填 SKU："
            f"{len(auto_map)} 条{Color.RESET}"
        )

    still_missing = [s for s in missing_wh if s not in auto_map]
    if not still_missing:
        print("[继续] 缺失映射均已通过 product_sku 表自动补全")
    else:
        step1_path = _save_step1(px4_df_1, output_dir)
        print(f"中间结果已保存：{step1_path}")
        print(
            f"{Color.YELLOW}[待更新] product_sku_mapping 仍有 {len(still_missing)} 个 "
            f"warehouse_sku 无标准 SKU"
            f"（新建待填 {n_pending} 条；库中已有待更新 {len(pending_only)} 条；"
            f"自动生效 {n_auto} 条）{Color.RESET}"
        )
        preview = still_missing[:20]
        print("  样例：" + ", ".join(preview) + (" ..." if len(still_missing) > 20 else ""))
        print(
            f"{Color.CYAN}请手动更新 `{PSM_TABLE}` 或运行 upProductMapping.py：\n"
            f"  - partner_code={PARTNER_CODE_4PX}, partner_type={PARTNER_TYPE_WH}\n"
            f"  - 填写正确 product_sku，并设 is_active=1\n"
            f"  - 完成后重新执行本脚本{Color.RESET}"
        )
        raise SystemExit(2)

# ---------- 3. 映射 商品ID ----------
px4_df_2 = map_sku_to_product_uid(px4_df_1, main_sku="SKU")
warn_blank_product_uid(px4_df_2)

step1_path = _save_step1(px4_df_2, output_dir)
print(f"平台分摊，结果已保存到{step1_path}")

# ---------- 4. K0「各平台商品ID周转明细」按销售平台分摊 ----------
# 口径：同一商品ID 下，仅对有效销售平台（排除 无/其他/ALL）按「可售库存-可调」占比分摊；
#   海外仓仓租费 = 总仓租 × (该平台有效库存 / 该商品ID有效库存合计)
# 无有效库存 / 无匹配 → 整笔进 Sheet2「无平台-仓租费用」（总额守恒：Sheet1+Sheet2=总仓租）
ku_cun_path = (
    fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\{ku_cun_date}{KU_CUN_FILE_SUFFIX}"
)
if not os.path.isfile(ku_cun_path):
    raise FileNotFoundError(
        f"未找到 K0 库存动销明细：{ku_cun_path}\n"
        f"请先运行 modules/K_storage_fee_product_info/K0_库存周转.py"
    )

ku_cun_df = pd.read_excel(ku_cun_path, sheet_name=KU_CUN_SHEET)
need_cols = ["商品ID", PLATFORM_COL, QTY_COL]
missing = [c for c in need_cols if c not in ku_cun_df.columns]
if missing:
    raise KeyError(f"「{KU_CUN_SHEET}」缺少列 {missing}：{ku_cun_path}")

ku_cun_df = _prepare_platform_qty_ratio(ku_cun_df[need_cols])

DF = pd.merge(px4_df_2, ku_cun_df, on="商品ID", how="left")
DF["数量占比"] = pd.to_numeric(DF["数量占比"], errors="coerce")
DF["仓租"] = DF["总仓租"] * DF["数量占比"]
DF.loc[DF[PLATFORM_COL].isin(NO_SITE_PLATFORMS), "仓租"] = 0
DF["仓租"] = pd.to_numeric(DF["仓租"], errors="coerce")

result_DF = DF.loc[
    DF[PLATFORM_COL].notna() & ~DF[PLATFORM_COL].isin(NO_SITE_PLATFORMS),
    ["SKU", "商品ID", PLATFORM_COL, QTY_COL, "仓租"],
].copy()
result_DF = result_DF.rename(columns={"仓租": "海外仓仓租费"})

px4_have_site_cang_zu = float(result_DF["海外仓仓租费"].fillna(0).sum())
px4_no_site_fen_tan = px4_all_cang_zu - px4_have_site_cang_zu
print(
    f"[分摊] 维度=有效销售平台+商品ID+{QTY_COL}（不含无/其他/ALL）；"
    f"总仓租={px4_all_cang_zu:.4f}，"
    f"已分摊到销售平台={px4_have_site_cang_zu:.4f}，"
    f"无平台-仓租费用={px4_no_site_fen_tan:.4f}，"
    f"核对合计={px4_have_site_cang_zu + px4_no_site_fen_tan:.4f}"
)

result_DF["无平台-仓租费用"] = None
if len(result_DF) > 0:
    result_DF.at[result_DF.index[0], "无平台-仓租费用"] = px4_no_site_fen_tan
result_DF = result_DF[PX4_SHEET1_COLUMNS]

no_platform_DF = _with_no_platform_total(
    _build_no_platform_detail(DF),
    px4_no_site_fen_tan,
    PX4_SHEET2_COLUMNS,
)

output_file_path = os.path.join(output_dir, "(平台分摊)4PX-仓租明细.xlsx")
_save_alloc_workbook(result_DF, no_platform_DF, output_file_path)
print(
    f"平台分摊，结果已保存到{output_file_path}"
    f"（{SHEET_PLATFORM} / {SHEET_NO_PLATFORM}）"
)
