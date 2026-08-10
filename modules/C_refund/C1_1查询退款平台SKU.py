"""C1_2：RMA 退款清洗 + 从 DB 映射「平台sku」。

用 sales_order_shipped（order_no + warehouse_sku → platform_sku）批量补平台 SKU；
重点店铺 LM_BC_FR / LM_RP_FR 会打印映射完整度检查。

输入：{DESKTOP_ROOT}/{folder_name}{shared_date}/RMA/RMA-{shared_date}.xlsx
输出：同目录 (已完成-1)RMA-{shared_date}.xlsx → 供 C1_3 使用

处理步骤：
1. 自动定位 RMA 表头，去掉作废单、误操作订单号黑名单
2. 非作废行若「退款金额」为 0 → 报错（多为未审核/漏标记）
3. 批量查 sales_order_shipped，按 退款原订单号||RMA产品 写入「平台sku」
   （同键多条不同 platform_sku 时取 ship_time 较新，其次 id）
4. 输出指定列；对 LM_BC_FR / LM_RP_FR 打印未映射样例
"""

from __future__ import annotations

import importlib.util
import sys
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd
import pymysql.cursors

# 须在 import config/common 之前：加载项目根到 sys.path
_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_PROJECT_ROOT = _epr_mod.bootstrap(__file__)

from common.style import Color
from config.A0_paths import DESKTOP_ROOT
from config.A0_set_date import folder_name, shared_date

from database.db_connection import get_db_manager  # noqa: E402

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------

SHIPPED_TABLE = "sales_order_shipped"
_KEY_CHUNK = 200

# RMA 表必须列（自动找表头时校验）
_RMA_REQUIRED_COLS = (
    "平台",
    "店铺英文名",
    "订单目的国家",
    "退款原订单号",
    "RMA产品",
    "RMA产品数量",
    "退款金额",
    "退款状态",
)

# 写出列（仓库名称供 C1_3 判断「分销」；缺失则补空列）
_OUTPUT_COLS = [
    "平台",
    "店铺英文名",
    "仓库名称",
    "订单目的国家",
    "退款原订单号",
    "平台sku",
    "RMA产品",
    "RMA产品数量",
    "退款金额",
    "退款状态",
]

# 误操作 / 历史脏单：不参与退款统计（按「退款原订单号」排除）
_IGNORE_ORDER_NOS: frozenset[str] = frozenset(
    {
        # "real-MZC3LC9",
        # "real-MYCAAC9",
        # "real-M5X6CE9",
        # "real-MBW8TE9",
        # "real-M5W8TE9",
        # "real-MAPC169",
        # "real-M6H1569",
        # "real-M61PHE9",
        # "real-MSSEC69",
        # "real-MQBES79",
        # "real-MDZUBF9",
        # "real-MP9R66Q",
        # "real-MGBE87Q",
        # "WEC0612408260019",
        # "WEC0362410010060",
        # "WEC0142409230035",
        # "WEC0272407190079",
        # "WEC0792409080009",
        # "WEC0552412030033",
        # "WEC0452501010040",
        # "WEC0642501100024",
        # "WEC0272502050137",
        # "WEC0492407190078",
        # "WEC0482505110030",
        # "WEC0312407170018",
        # "303-2724093-0301127",
        # "404-0584444-5506747",
    }
)

# 映射完整度重点检查店铺（避免全量刷屏）
_FOCUS_SHOPS = frozenset({"LM_BC_FR", "LM_RP_FR"})


# ---------------------------------------------------------------------------
# 读写 / 清洗
# ---------------------------------------------------------------------------

def _strip_df_strings(df: pd.DataFrame) -> pd.DataFrame:
    """对 object 列去首尾空格（就地语义：返回处理后的副本）。"""
    out = df.copy()
    for col in out.columns:
        if out[col].dtype == object:
            out[col] = out[col].map(lambda x: x.strip() if isinstance(x, str) else x)
    return out


def _read_rma_excel(path: Path | str) -> pd.DataFrame:
    """
    读取 ERP 导出的 RMA 表。
    表头前常有 2~3 行说明：优先扫描前 8 行，定位含「退款状态」的表头行；
    失败再回退 skiprows=3 / 2。
    """
    path = Path(path)
    preview = pd.read_excel(path, header=None, nrows=8)
    for i, row in preview.iterrows():
        cells = {str(v).strip() for v in row if pd.notna(v)}
        if "退款状态" not in cells:
            continue
        df = pd.read_excel(path, header=int(i))
        missing = [c for c in _RMA_REQUIRED_COLS if c not in df.columns]
        if not missing:
            return df

    for skip in (3, 2):
        df = pd.read_excel(path, skiprows=skip)
        missing = [c for c in _RMA_REQUIRED_COLS if c not in df.columns]
        if not missing:
            return df

    cols = list(pd.read_excel(path, skiprows=2, nrows=0).columns)
    raise KeyError(
        f"RMA 文件缺少必要列（含「退款状态」），请检查表头。"
        f"路径: {path}；skiprows=2 时列名: {cols[:10]}..."
    )


def _assert_no_zero_refund(df: pd.DataFrame) -> None:
    """非作废行不允许「退款金额」为 0（多为 TEMU 等未审核漏标）。"""
    amt = pd.to_numeric(df["退款金额"], errors="coerce")
    zero_mask = amt.eq(0)
    if not zero_mask.any():
        return
    bad_orders = df.loc[zero_mask, "退款原订单号"].astype(str).tolist()
    raise ValueError(
        f'退款金额列存在 0 值，对应「退款原订单号」：{bad_orders}，'
        f"询问相应运营，是否忘记标记、审核退款订单"
        f"（TEMU 退款：一帆录入 / 一帆审核）！！！"
    )


# ---------------------------------------------------------------------------
# DB：sales_order_shipped → 平台 SKU
# ---------------------------------------------------------------------------

def _chunked_in_query(
    cur,
    sql_template: str,
    keys: list[str],
    extra_params: tuple = (),
) -> list[dict]:
    """按批次执行 IN 查询，返回合并后的 dict 行列表。"""
    if not keys:
        return []
    results: list[dict] = []
    for i in range(0, len(keys), _KEY_CHUNK):
        chunk = keys[i : i + _KEY_CHUNK]
        placeholders = ",".join(["%s"] * len(chunk))
        cur.execute(sql_template.format(placeholders=placeholders), extra_params + tuple(chunk))
        results.extend(cur.fetchall())
    return results


def _ship_time_ts(row: dict) -> float:
    """ship_time → unix 秒；无法解析或为空 → 0（视为更旧）。"""
    st = row.get("ship_time")
    if st is None or (isinstance(st, float) and pd.isna(st)):
        return 0.0
    if isinstance(st, datetime):
        return st.timestamp()
    ts = pd.to_datetime(st, errors="coerce")
    if pd.isna(ts):
        return 0.0
    try:
        return float(ts.timestamp())
    except (OSError, OverflowError, ValueError):
        return 0.0


def _row_id(row: dict) -> int:
    try:
        return int(row.get("id") or 0)
    except (TypeError, ValueError):
        return 0


def _is_newer(a: dict, b: dict) -> bool:
    """True 表示 a 比 b 更新：优先 ship_time，相同则用 id。"""
    ta, tb = _ship_time_ts(a), _ship_time_ts(b)
    if ta != tb:
        return ta > tb
    return _row_id(a) > _row_id(b)


def fetch_platform_sku_map_from_db(order_nos: list[str]) -> dict[str, str]:
    """
    从 sales_order_shipped 批量查询平台 SKU。

    关联键：
      退款原订单号 = order_no
      RMA产品      = warehouse_sku

    返回：{"order_no||warehouse_sku": platform_sku}
    同键多条且 platform_sku 不同 → 取 ship_time/id 较新的一条。
    """
    order_nos = sorted({str(x).strip() for x in order_nos if x and str(x).strip()})
    if not order_nos:
        return {}

    sql = f"""
        SELECT id, ship_time, order_no, warehouse_sku, platform_sku
        FROM `{SHIPPED_TABLE}`
        WHERE order_no IN ({{placeholders}})
          AND platform_sku IS NOT NULL
          AND TRIM(platform_sku) <> ''
          AND warehouse_sku IS NOT NULL
          AND TRIM(warehouse_sku) <> ''
        ORDER BY id ASC
    """
    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        rows = _chunked_in_query(cur, sql, order_nos)
    finally:
        cur.close()
        conn.close()

    chosen_row: dict[str, dict] = {}
    dup_keys: set[str] = set()
    for row in rows:
        order_no = str(row.get("order_no") or "").strip()
        wh_sku = str(row.get("warehouse_sku") or "").strip()
        platform_sku = str(row.get("platform_sku") or "").strip()
        if not order_no or not wh_sku or not platform_sku:
            continue
        key = f"{order_no}||{wh_sku}"
        prev = chosen_row.get(key)
        if prev is None:
            chosen_row[key] = row
            continue
        prev_sku = str(prev.get("platform_sku") or "").strip()
        if prev_sku != platform_sku:
            dup_keys.add(key)
            if _is_newer(row, prev):
                chosen_row[key] = row

    if dup_keys:
        print(
            f"{Color.CYAN}[DB][警告] `{SHIPPED_TABLE}` 中存在 "
            f"「order_no+warehouse_sku」对应多个不同 platform_sku，"
            f"已按 ship_time/id 选择较新记录（冲突键 {len(dup_keys)} 个）。{Color.RESET}"
        )

    return {k: str(r.get("platform_sku") or "").strip() for k, r in chosen_row.items()}


def _apply_platform_sku(df: pd.DataFrame, sku_map: dict[str, str]) -> pd.DataFrame:
    """按 退款原订单号||RMA产品 写入「平台sku」。"""
    out = df.copy()
    out["平台sku"] = ""
    if not sku_map:
        print(f"{Color.YELLOW}[DB] sales_order_shipped 未查到可用平台 SKU 映射{Color.RESET}")
        return out

    map_key = out["退款原订单号"].astype(str).str.strip() + "||" + out["RMA产品"].astype(str).str.strip()
    matched = map_key.isin(sku_map)
    out.loc[matched, "平台sku"] = map_key[matched].map(sku_map)
    hit = int(matched.sum())
    print(
        f"{Color.GREEN}[DB] `{SHIPPED_TABLE}`：映射表 {len(sku_map)} 条键，"
        f"命中 RMA 行 {hit}/{len(out)}{Color.RESET}"
    )
    return out


def _print_focus_shop_mapping_report(df: pd.DataFrame) -> None:
    """对 LM_BC_FR / LM_RP_FR 打印平台 SKU 映射完整度与未映射样例。"""
    present = sorted(
        {
            str(s).strip()
            for s in df.loc[df["店铺英文名"].isin(_FOCUS_SHOPS), "店铺英文名"].dropna()
        }
    )
    for shop in present:
        shop_df = df.loc[df["店铺英文名"].astype(str).str.strip() == shop]
        total = len(shop_df)
        if total == 0:
            print(f"{Color.CYAN}[检查] {shop}：无退款行{Color.RESET}")
            continue

        mapped_mask = shop_df["平台sku"].notna() & (shop_df["平台sku"].astype(str).str.strip() != "")
        mapped_cnt = int(mapped_mask.sum())
        unmapped_cnt = total - mapped_cnt

        if unmapped_cnt == 0:
            print(
                f"{Color.GREEN}[检查] {shop}：已全部映射平台SKU（{mapped_cnt}/{total}）{Color.RESET}"
            )
            continue

        print(
            f"{Color.YELLOW}[检查] {shop}：未完全映射"
            f"（已映射 {mapped_cnt}/{total}，未映射 {unmapped_cnt}）{Color.RESET}"
        )
        cols = [
            c
            for c in ("退款原订单号", "RMA产品", "RMA产品数量", "退款金额", "退款状态")
            if c in shop_df.columns
        ]
        sample = shop_df.loc[~mapped_mask, cols].head(20)
        if not sample.empty:
            print(f"{Color.YELLOW}[检查] 未映射样例（前20条）：{Color.RESET}")
            print(sample.to_string(index=False))


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

# TODO 文件路径！！！依赖 A0 日期 / 桌面目录
rma_dir = Path(DESKTOP_ROOT) / f"{folder_name}{shared_date}" / "RMA"
rma_path = rma_dir / f"RMA-{shared_date}.xlsx"

rma_df = _strip_df_strings(_read_rma_excel(rma_path))

# 去掉作废；再排除误操作订单号黑名单
rma_df = rma_df.loc[rma_df["退款状态"] != "作废"].copy()
rma_df = rma_df.loc[~rma_df["退款原订单号"].astype(str).str.strip().isin(_IGNORE_ORDER_NOS)].copy()

_assert_no_zero_refund(rma_df)

# 全店铺一次查库：退款原订单号 → (order_no||warehouse_sku → platform_sku)
order_nos = (
    rma_df["退款原订单号"].dropna().astype(str).str.strip().unique().tolist()
)
sku_map = fetch_platform_sku_map_from_db(order_nos)
rma_df = _apply_platform_sku(rma_df, sku_map)

# C1_3 需要「仓库名称」；旧导出若无该列则补空，避免下游 KeyError
if "仓库名称" not in rma_df.columns:
    rma_df["仓库名称"] = ""

rma_out = rma_df[[c for c in _OUTPUT_COLS if c in rma_df.columns]].copy()

output_path = rma_dir / f"(已完成-1)RMA-{shared_date}.xlsx"
rma_out.to_excel(output_path, index=False)
print(f"处理完成，文件另存为：{output_path}")

_print_focus_shop_mapping_report(rma_out)
