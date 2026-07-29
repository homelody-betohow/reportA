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

from common.platform_shop import map_shop_platform_region
from common.style import Color
from config.A0_set_date import shared_date, folder_name, fba_date
from config.A0_paths import DESKTOP_ROOT
from database.db_connection import get_db_manager

PRODUCT_SKU_TABLE = "product_sku"
_KEY_CHUNK = 200

# 忽略特定的 UserWarning
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
print(f"{fba_date}")
# TODO 文件路径！！！   上月的 利润报表
# main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\FBA仓租\FBA仓租明细{fba_date}.xlsx"
# FBA仓租 引用的是 SellerSku利润报表， 将 对应日期文件 重命名 即可
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\FBA仓租\FBA仓租明细{fba_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)

# 删除 UK站点 的数据行
main_file_df = main_file_df[main_file_df['站点'] != 'UK']

# 使用 sellerSku 列的数据填充 仓库sku 列的空值
main_file_df['仓库sku'] = main_file_df['仓库sku'].fillna(main_file_df['sellerSku'])

def extract_values(s):
    if pd.isna(s):  # 检查是否为 NaN
        return None  # 如果是 NaN，返回 None 或其他默认值
    if 'amzn.gr.' in s:
        return s.split(r'amzn.gr.')[-1].split('-')[0].split('_')[0]
    else:
        return s.split('#')[0].split('BCFBAFL')[0].split('FBFBAFL')[0]


# 应用提取规则，清洗 仓库sku
main_file_df['仓库sku'] = main_file_df['仓库sku'].apply(extract_values)
main_file_df = main_file_df.rename(columns={'仓库sku': 'SKU'})

# 映射 商品ID → 主产品编码（product_uid → 首个 product_sku）
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


# 映射 商品ID：product_sku.product_sku → product_uid
main_file_df_1 = map_sku_to_product_uid(main_file_df, main_sku="SKU")
warn_blank_product_uid(main_file_df_1)

# 映射站点 / 映射平台（数据源：platform_shop）
main_file_df_3 = map_shop_platform_region(main_file_df_1, shop_col='店铺', site_col=None)

# 在 映射站点 后插入新列 站点商品ID识别码
new_column_name = "站点商品ID识别码"
new_column_data = main_file_df_3["映射站点"] + main_file_df_3["商品ID"]
target_column = "映射站点"
insert_position = main_file_df_3.columns.get_loc(target_column) + 1
main_file_df_3.insert(insert_position, new_column_name, new_column_data)

# 在 站点商品ID识别码后插入 平台商品ID识别码
new_column_name = "平台商品ID识别码"
new_column_data = main_file_df_3["映射平台"] + main_file_df_3["商品ID"]
target_column = "站点商品ID识别码"
insert_position = main_file_df_3.columns.get_loc(target_column) + 1
main_file_df_3.insert(insert_position, new_column_name, new_column_data)

# 得到上月的  FBA仓租费
main_file_df_3['FBA仓租费'] = main_file_df_3['仓储费用（已分摊）'] + main_file_df_3['长期仓储费（已分摊）']
main_file_df_3['FBA仓租费'] = main_file_df_3['FBA仓租费'].fillna(0).abs()  # 去掉 负号

# 删除"sellerSku"为空的行
main_file_df_4 = main_file_df_3.dropna(subset=['sellerSku'])
# 删除"FBA仓租费"为0的行
main_file_df_4 = main_file_df_4[main_file_df_4['FBA仓租费'] != 0]

main_file_df_4 = main_file_df_4[
    ['sellerSku', 'ASIN', '产品信息', 'SKU','商品ID', '店铺', '映射站点', '映射平台', '站点商品ID识别码',
     '平台商品ID识别码', '仓储费用（已分摊）', '长期仓储费（已分摊）', 'FBA仓租费']]
# 保存修改
output_path = main_file_path.rsplit('\\', 1)[0] + '\\(已完成-1)' + main_file_path.rsplit('\\', 1)[1]
main_file_df_4.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')

# 获取sellerSku为空的数据
empty_sellerSku_df = main_file_df_3[main_file_df_3['sellerSku'].isna()]
empty_sellerSku_FBA = empty_sellerSku_df['FBA仓租费'].sum()
print(f'\n---------------------sellerSku为空的FBA仓租费是：{empty_sellerSku_FBA}EUR------------------------------')
