import importlib.util
import warnings
import sys
from pathlib import Path

# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

import pandas as pd
import pymysql.cursors
from datetime import datetime
from A_报表.Z_method.style import Color
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")

_RMA_REQUIRED_COLS = (
    "平台", "店铺英文名", "订单目的国家", "退款原订单号",
    "RMA产品", "RMA产品数量", "退款金额", "退款状态",
)


def _read_rma_excel(path: str) -> pd.DataFrame:
    """读取 ERP 导出的 RMA 表（表头前可能有 2~3 行说明，自动定位含「退款状态」的表头行）。"""
    preview = pd.read_excel(path, header=None, nrows=8)
    for i, row in preview.iterrows():
        cells = {str(v).strip() for v in row if pd.notna(v)}
        if "退款状态" in cells:
            df = pd.read_excel(path, header=i)
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


# TODO 文件路径！！！
RMA_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\RMA\RMA-{shared_date}.xlsx"
RMA_file_df = _read_rma_excel(RMA_file_path)
# 去除 整张表 的前后空格
for col in RMA_file_df.columns:
    RMA_file_df[col] = RMA_file_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
# 检查 是否有 "退款状态" != "作废"，且 退款金额 == 0
# 筛选出“退款状态”列中不 等于 “作废”的行
RMA_file_df_1 = RMA_file_df[RMA_file_df["退款状态"] != "作废"].copy()

# 去除 误操作的订单，退款原订单号 在列表
ignoreList = [
'real-MZC3LC9',
'real-MYCAAC9',
'real-M5X6CE9',
'real-MBW8TE9',
'real-M5W8TE9',
'real-MAPC169',
'real-M6H1569',
'real-M61PHE9',
'real-MSSEC69',
'real-MQBES79',
'real-MDZUBF9',
'real-MP9R66Q',
'real-MGBE87Q',
'WEC0612408260019',
'WEC0362410010060',
'WEC0142409230035',
'WEC0272407190079',
'WEC0792409080009',
'WEC0552412030033',
'WEC0452501010040',
'WEC0642501100024',
'WEC0272502050137',
'WEC0492407190078',
'WEC0482505110030',
'WEC0312407170018',
'303-2724093-0301127',
'404-0584444-5506747',
'manomano-M240773382753',
'manomano-M240369950127',
'manomano-M2501504225715',
'manomano-M2510505226785'
]
RMA_file_df_1 = RMA_file_df_1[~RMA_file_df_1['退款原订单号'].isin(ignoreList)]
# ========================================================================= #

# 1. 找出所有退款金额为 0 的行
zero_mask = RMA_file_df_1['退款金额'] == 0
# 2. 如果有 0 值，主动报错并返回对应的“退款原退款原订单号”
if zero_mask.any():
    bad_orders = RMA_file_df_1.loc[zero_mask, '退款原订单号'].tolist()
    raise ValueError(
        f'退款金额列存在 0 值，对应"退款原订单号"：{bad_orders}，询问相应运营，是否忘记标记、审核，退款订单(TEMU退款，一帆录入，一帆审核) ！！！')

# LM_BC、LM_RP 的  退款订单  的 平台SKU 映射
_REPORT_PRA_ROOT = next(
    (p / "reportPRA" for p in Path(__file__).resolve().parents if (p / "reportPRA").is_dir()),
    None,
)
if _REPORT_PRA_ROOT and str(_REPORT_PRA_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPORT_PRA_ROOT))

from database.db_connection import get_db_manager  # noqa: E402  # pyright: ignore[reportMissingImports]

SHIPPED_TABLE = "sales_order_shipped"
_KEY_CHUNK = 200


def _chunked_in_query(cur, sql_template: str, keys: list[str], extra_params: tuple = ()) -> list[dict]:
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


def fetch_platform_sku_map_from_db(order_nos: list[str]) -> dict[str, str]:
    """
    从 sales_order_shipped 批量查询平台 SKU 映射。

    关联键：
      - 退款原订单参考号 = sales_order_shipped.order_no
      - RMA产品 = sales_order_shipped.warehouse_sku

    返回： "order_no||warehouse_sku" -> "platform_sku"
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

    def _ts(row: dict) -> float:
        """
        ship_time 可能是 datetime / pandas Timestamp / 字符串 / NULL。
        规则：能解析就用其时间戳；解析失败或为空则视为 0（更旧）。
        """
        st = row.get("ship_time")
        if isinstance(st, datetime):
            return st.timestamp()

        # pandas.Timestamp 等
        if st is not None and hasattr(st, "timestamp"):
            try:
                return float(st.timestamp())
            except Exception:
                pass

        # MySQL 有时会返回字符串
        if isinstance(st, str):
            s = st.strip()
            if not s:
                return 0.0
            # 兼容 "YYYY-mm-dd HH:MM:SS" / "YYYY-mm-ddTHH:MM:SS" / 带毫秒
            try:
                return datetime.fromisoformat(s.replace(" ", "T")).timestamp()
            except Exception:
                return 0.0

        return 0.0

    def _rid(row: dict) -> int:
        try:
            return int(row.get("id") or 0)
        except (TypeError, ValueError):
            return 0

    def _is_newer(a: dict, b: dict) -> bool:
        """
        True 表示 a 比 b 更新。
        规则：优先 ship_time 更新；若 ship_time 相同或都不可用，再用 id 兜底。
        """
        ta, tb = _ts(a), _ts(b)
        if ta != tb:
            return ta > tb
        return _rid(a) > _rid(b)

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
            # 发生冲突时：选 ship_time 最新的一条；若无法比较时间则用 id 兜底
            if _is_newer(row, prev):
                chosen_row[key] = row

    if dup_keys:
        print(
            f"{Color.CYAN}[DB][警告] sales_order_shipped 表中存在「order_no+warehouse_sku」对应多个不同 platform_sku，"
            f"已按 ship_time/id 选择较新记录。{Color.RESET}"
        )

    mapping: dict[str, str] = {}
    for k, r in chosen_row.items():
        mapping[k] = str(r.get("platform_sku") or "").strip()
    return mapping


_FOCUS_SHOPS = {"LM_BC_FR", "LM_RP_FR"}
RMA_file_df_1["平台sku"] = ""

# 预先构造 RMA 侧复合键（退款原订单号||RMA产品）
rma_map_key_all = RMA_file_df_1["退款原订单号"].astype(str) + "||" + RMA_file_df_1["RMA产品"].astype(str)

# 一次性查询：覆盖所有店铺的退款原订单号
all_order_nos = (
    RMA_file_df_1["退款原订单号"].dropna().astype(str).str.strip().unique().tolist()
)
sku_map_all = fetch_platform_sku_map_from_db(all_order_nos)

if sku_map_all:
    print(
        f"{Color.GREEN}[DB] 全部店铺：从 sales_order_shipped 查到 {len(sku_map_all)} 条 "
        f"order_no+warehouse_sku 平台SKU 映射{Color.RESET}"
    )

matched = rma_map_key_all.isin(sku_map_all)
RMA_file_df_1.loc[matched, "平台sku"] = rma_map_key_all[matched].map(sku_map_all)
# RMA_file_df_1.loc[matched, "退款原订单号"] += '——已映射"' + RMA_file_df_1.loc[matched, "平台sku"] + '"'

# 仍只提示重点店铺检查（避免刷屏）
have_order_shop_list: list[str] = sorted(
    set(RMA_file_df_1.loc[RMA_file_df_1["店铺英文名"].isin(_FOCUS_SHOPS), "店铺英文名"].dropna().tolist())
)

# 保留指定列
RMA_file_df_1 = RMA_file_df_1[
    ['平台', '店铺英文名', '仓库名称','订单目的国家', '退款原订单号', '平台sku', 'RMA产品', 'RMA产品数量', '退款金额', '退款状态']]

# 保存结果
output_path = RMA_file_path.rsplit('\\', 1)[0] + '\\(已完成-1)' + RMA_file_path.rsplit('\\', 1)[-1]
RMA_file_df_1.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
for have_order_shop in have_order_shop_list:
    shop_mask = RMA_file_df_1["店铺英文名"].astype(str).str.strip() == have_order_shop
    shop_df = RMA_file_df_1.loc[shop_mask].copy()
    total = len(shop_df)
    if total == 0:
        print(f'{Color.CYAN}[检查] {have_order_shop}：无退款行{Color.RESET}')
        continue

    mapped_mask = shop_df["平台sku"].notna() & (shop_df["平台sku"].astype(str).str.strip() != "")
    mapped_cnt = int(mapped_mask.sum())
    unmapped_cnt = int(total - mapped_cnt)

    if unmapped_cnt == 0:
        print(f'{Color.GREEN}[检查] {have_order_shop}：已全部映射平台SKU（{mapped_cnt}/{total}）{Color.RESET}')
        continue

    print(f'{Color.YELLOW}[检查] {have_order_shop}：未完全映射（已映射 {mapped_cnt}/{total}，未映射 {unmapped_cnt}）{Color.RESET}')
    cols = [c for c in ["退款原订单号", "RMA产品", "RMA产品数量", "退款金额", "退款状态"] if c in shop_df.columns]
    sample = shop_df.loc[~mapped_mask, cols].head(20)
    if not sample.empty:
        print(f"{Color.YELLOW}[检查] 未映射样例（前20条）：{Color.RESET}")
        print(sample.to_string(index=False))
