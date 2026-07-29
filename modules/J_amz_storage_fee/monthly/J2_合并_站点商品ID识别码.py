import warnings
import pandas as pd
import pymysql.cursors
import importlib.util
from pathlib import Path
# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.style import Color
from config.A0_set_date import shared_date, folder_name, fba_date
from config.A0_paths import DESKTOP_ROOT
from database.db_connection import get_db_manager

PRODUCT_SKU_TABLE = "product_sku"
_KEY_CHUNK = 200

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# TODO 文件路径！！！   上月的 利润报表
Amazon_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\FBA仓租\(已完成-1)FBA仓租明细{fba_date}.xlsx"
Amazon_file_df = pd.read_excel(Amazon_file_path)

# 重命名
Amazon_file_df = Amazon_file_df.rename(columns={'映射站点': '站点'})
Amazon_file_df = Amazon_file_df.rename(columns={'映射平台': '平台'})
# 删除列 'FBA仓租费' 中值为 0 的行
Amazon_file_df = Amazon_file_df[Amazon_file_df['FBA仓租费'] != 0]

# 按照 '站点商品ID识别码' 列进行分组汇总
result_df = Amazon_file_df.groupby('站点商品ID识别码').agg({
    'SKU': 'first',
    '商品ID': 'first',
    '站点': 'first',
    '平台': 'first',
    '平台商品ID识别码': 'first',
    'FBA仓租费': 'sum'
}).reset_index()


#替换-产品信息库2025.xlsx
# 映射 商品ID → 主产品编码（product_uid → 首个 product_sku）
def _fetch_first_sku_by_uid(uids: list[str]) -> dict[str, str]:
    """
    product_uid → 第一个 product_sku（按 id 升序，对齐原 Excel keep='first'）。
    """
    uids = sorted({str(x).strip() for x in uids if x and str(x).strip()})
    if not uids:
        return {}

    mapping: dict[str, str] = {}
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor(pymysql.cursors.DictCursor) as cur:
            for i in range(0, len(uids), _KEY_CHUNK):
                chunk = uids[i : i + _KEY_CHUNK]
                placeholders = ", ".join(["%s"] * len(chunk))
                sql = f"""
                    SELECT product_uid, product_sku
                    FROM `{PRODUCT_SKU_TABLE}`
                    WHERE product_uid IN ({placeholders})
                      AND is_deleted = 0
                      AND product_sku IS NOT NULL
                      AND TRIM(product_sku) <> ''
                    ORDER BY id ASC
                """
                cur.execute(sql, chunk)
                for row in cur.fetchall():
                    uid = str(row.get("product_uid") or "").strip()
                    sku = str(row.get("product_sku") or "").strip()
                    if uid and sku and uid not in mapping:
                        mapping[uid] = sku
    finally:
        conn.close()
    return mapping


def map_product_uid_to_sku(main_df: pd.DataFrame, main_uid: str = "商品ID") -> pd.DataFrame:
    """
    商品ID（product_uid）→ 主产品编码（product_sku）。
    未命中保留原 SKU；原 商品ID 带 -NW 时，映射 SKU 缀回 -NW。
    """
    out = main_df.copy()
    if main_uid not in out.columns:
        raise KeyError(f"主表缺少列 {main_uid!r}，当前列: {list(out.columns)}")
    if "SKU" not in out.columns:
        raise KeyError(f"主表缺少列 'SKU'，当前列: {list(out.columns)}")

    series = out[main_uid].astype(str).str.strip()
    invalid = series.isin(("", "nan", "None", "NaN")) | out[main_uid].isna()
    nw_mask = series.str.endswith("-NW", na=False) & ~invalid
    series_no_nw = series.mask(nw_mask, series.str.replace(r"-NW$", "", regex=True))

    uid_sku_map = _fetch_first_sku_by_uid(series_no_nw[~invalid].tolist())
    print(f"[DB] product_sku 命中 {len(uid_sku_map)} 条 product_uid → 首个 product_sku")

    mapped = series_no_nw.map(uid_sku_map)
    miss = (~invalid) & mapped.isna()
    mapped = mapped.mask(nw_mask & mapped.notna(), mapped.astype(str) + "-NW")

    # 未命中：保留分组后的原 SKU
    orig_sku = out["SKU"]
    mapped = mapped.where(mapped.notna(), orig_sku)
    mapped = mapped.mask(invalid, orig_sku)

    out = out.rename(columns={"SKU": "原-SKU"})
    insert_pos = out.columns.get_loc(main_uid) + 1
    out.insert(insert_pos, "SKU", mapped)

    n_miss = int(miss.sum())
    if n_miss:
        preview_cols = [c for c in ("原-SKU", "SKU", "商品ID", "站点", "FBA仓租费") if c in out.columns]
        preview = out.loc[miss, preview_cols].head(10)
        print(
            f"{Color.YELLOW}[检查] 商品ID 有 {n_miss} 行未命中 product_sku"
            f"（已保留原 SKU），请核对：{Color.RESET}"
        )
        print(preview.to_string(index=False))
    return out


# 商品ID → 主产品编码（product_uid → 首个 product_sku）
result_df_1 = map_product_uid_to_sku(result_df, main_uid="商品ID")

result_df_1 = result_df_1[['SKU', '商品ID', '站点', '平台', '站点商品ID识别码', '平台商品ID识别码', 'FBA仓租费']]

# 保存修改
output_path = Amazon_file_path.replace('已完成-1', '处理完成')
result_df_1.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
