import csv
import chardet
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

from common.sku_mapping import sku_mappings
from common.platform_shop import map_region_to_platform
from common.style import Color
from config.A0_set_date import shared_date, folder_name
from common.split_rows_data_SKU import split_one_rows_data
from config.A0_paths import DESKTOP_ROOT
from database.db_connection import get_db_manager

PRODUCT_SKU_TABLE = "product_sku"
_KEY_CHUNK = 200

# TODO 文件路径！！！
otto_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\广告\OTTO\OTTO-广告数据-{shared_date}.csv"
with open(otto_file_path, 'rb') as file:
    raw_data = file.read()
    result = chardet.detect(raw_data)
# 检测文件编码
encoding = result['encoding']
print(f"文件的编码是: {encoding}")

# 读取并处理文件内容
with open(otto_file_path, encoding=encoding) as f:
    # 跳过前两行
    for _ in range(2):
        next(f)
    # 创建TSV阅读器             分割规则：分号
    reader = csv.reader(f, delimiter=';')
    new_data = []
    # 逐行读取
    for row in reader:
        # 这里row是一个列表，包含该行的所有列
        new_row = []
        for cell in row:
            cell = cell.strip()
            try:
                new_row.append(int(cell))
            except ValueError:
                try:
                    new_row.append(float(cell))
                except ValueError:
                    new_row.append(cell)
        new_data.append(new_row)

    # 创建DataFrame
    # new_data[1:] 表示跳过第一行，取出了从第二行开始的所有行作为 DataFrame 的数据部分。
    # columns 参数用于指定 DataFrame 的列名。new_data[0]：new_data 的第一行作为列名
    otto_file_df = pd.DataFrame(new_data[1:], columns=new_data[0])

# 先去除欧元符号和前后空格                     空值，替换成：0
otto_file_df['Ausgaben'] = otto_file_df['Ausgaben'].apply(
    lambda x: float(x.replace('€', '').replace(',', '.').strip()) if isinstance(x, str) and x.strip() else 0
)


# 筛选出 'Ausgaben' 列 中值不等于 0 的行（相对应删掉=0的行）
otto_file_df = otto_file_df[otto_file_df['Ausgaben'] != 0]

# 去除 整张表 的前后空格
for col in otto_file_df.columns:
    otto_file_df[col] = otto_file_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

# # 保存原始文件
# original_file_oath = otto_file_path.rsplit('\\', 1)[0] + '\\(original)' + otto_file_path.rsplit('\\', 1)[-1]
# otto_file_df.to_csv(original_file_oath, index=False)

# OTTO 源表部分行 SKU 为空，用「广告-SKU关系对应」按货号(Artikelnummer)补仓库SKU
product_map_sku_path = fr"{DESKTOP_ROOT}\广告-SKU关系对应.xlsx"
otto_file_df = sku_mappings(
    main_df=otto_file_df,
    main_sku='Artikelnummer',
    map_sku_path=product_map_sku_path,
    map_old_sku='货号',
    map_new_sku='仓库SKU',
    map_sku_sheet='OTTO 货号对应表'
)
empty_sku = otto_file_df['SKU'].isna() | (otto_file_df['SKU'].astype(str).str.strip() == '')
# sku_mappings 未命中时会回填原货号，故仅在映射结果与货号不同时视为命中
mapped_hit = (
    otto_file_df['映射仓库SKU'].notna()
    & (otto_file_df['映射仓库SKU'].astype(str).str.strip()
       != otto_file_df['Artikelnummer'].astype(str).str.strip())
)
otto_file_df.loc[empty_sku & mapped_hit, 'SKU'] = otto_file_df.loc[empty_sku & mapped_hit, '映射仓库SKU']

#  拆分有“+”的sku
otto_file_df_1 = split_one_rows_data(
    input_df=otto_file_df,
    data_column='SKU',
    value_column='Ausgaben'
)
# 中间输入一份，方便核对！
output_file_path = otto_file_path.rsplit('\\', 1)[0] + '\\(已完成-1)' + otto_file_path.rsplit('\\', 1)[-1]
otto_file_df_1.to_csv(output_file_path, index=False)  # index=False表示不保存索引列
print(f"处理完成，结果已保存到{output_file_path}")

def _fetch_first_sku_by_uid(uids: list[str]) -> dict[str, str]:
    """product_uid → 第一个 product_sku（按 id 升序，对齐原 Excel keep='first'）。"""
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


def map_product_uid_sku_col(main_df: pd.DataFrame, sku_col: str = "SKU") -> pd.DataFrame:
    """
    OTTO 部分行 SKU 实为商品ID（product_uid）：映射为首个产品编码。
    未命中保留原值；带 -NW 时剥后缀查库再缀回。
    """
    out = main_df.copy()
    if sku_col not in out.columns:
        raise KeyError(f"主表缺少列 {sku_col!r}，当前列: {list(out.columns)}")

    series = out[sku_col].astype(str).str.strip()
    invalid = series.isin(("", "nan", "None", "NaN")) | out[sku_col].isna()
    nw_mask = series.str.endswith("-NW", na=False) & ~invalid
    series_no_nw = series.mask(nw_mask, series.str.replace(r"-NW$", "", regex=True))

    uid_sku_map = _fetch_first_sku_by_uid(series_no_nw[~invalid].tolist())
    print(f"[DB] product_sku 命中 {len(uid_sku_map)} 条 product_uid → 首个 product_sku")

    mapped = series_no_nw.map(uid_sku_map)
    miss = (~invalid) & mapped.isna()
    mapped = mapped.mask(nw_mask & mapped.notna(), mapped.astype(str) + "-NW")
    mapped = mapped.where(mapped.notna(), out[sku_col])
    mapped = mapped.mask(invalid, out[sku_col])

    out = out.rename(columns={sku_col: "原-SKU"})
    insert_pos = out.columns.get_loc("原-SKU") + 1
    out.insert(insert_pos, sku_col, mapped)

    n_miss = int(miss.sum())
    if n_miss:
        preview_cols = [c for c in ("原-SKU", sku_col, "Artikelnummer", "Ausgaben") if c in out.columns]
        preview = out.loc[miss, preview_cols].head(10)
        print(
            f"{Color.YELLOW}[检查] 商品ID 有 {n_miss} 行未命中 product_sku"
            f"（已保留原 SKU），请核对：{Color.RESET}"
        )
        print(preview.to_string(index=False))
    return out


# OTTO 有部分SKU是商品ID，要映射回 第一个 SKU
# SKU 是否以'25-'开头，分成两个 df
mask = otto_file_df_1['SKU'].str.startswith('25-')
df_25 = otto_file_df_1.loc[mask].copy()
df_other = otto_file_df_1.loc[~mask].copy()
# 商品ID → 主产品编码（product_uid → 首个 product_sku）
df_25_1 = map_product_uid_sku_col(df_25, sku_col="SKU")
# --- 合并 ---
otto_file_df_1 = pd.concat([df_25_1, df_other]).sort_index()

# 在 SKU 后插入新列 SKU-站点识别码
new_column_name = "SKU-站点识别码"  # 新列名
new_column_data = "OTTO-BTH" + otto_file_df_1["SKU"]  # 新列数据，OTTO平台广告花费的站点都是：OTTO-BTH
target_column = "SKU"  # 目标列名（在其后插入）
insert_position = otto_file_df_1.columns.get_loc(target_column) + 1  # 计算插入位置
otto_file_df_1.insert(insert_position, new_column_name, new_column_data)  # 插入新列
# OTTO平台广告花费的站点都是：OTTO-BTH
otto_file_df_1['站点'] = "OTTO-BTH"

# 映射 平台（数据源：platform_shop）
otto_file_df_2 = map_region_to_platform(otto_file_df_1, site_col='站点')
# 在 SKU-站点识别码 后插入 SKU-平台识别码
new_column_name = "SKU-平台识别码"  # 新列名
new_column_data = otto_file_df_2["映射平台"] + otto_file_df_2["SKU"]  # 新列数据
target_column = "SKU-站点识别码"  # 目标列名（在其后插入）
insert_position = otto_file_df_2.columns.get_loc(target_column) + 1  # 计算插入位置
otto_file_df_2.insert(insert_position, new_column_name, new_column_data)  # 插入新列
# 保存目标列
otto_file_df_2 = otto_file_df_2[
    ['Artikelnummer', 'SKU', '站点', '映射平台', 'SKU-站点识别码', 'SKU-平台识别码', 'Ausgaben']]

# 更改列名，将’Ausgaben‘  改为 ’广告费(非AMZ)‘
otto_file_df_2 = otto_file_df_2.rename(columns={'Ausgaben': '广告费(非AMZ)'})

# 将处理后的数据保存到新的Excel文件
output_file_path = otto_file_path.rsplit('\\', 1)[0] + '\\(处理完成)OTTO广告.xlsx'
otto_file_df_2.to_excel(output_file_path, index=False)  # index=False表示不保存索引列

print(f"处理完成，结果已保存到{output_file_path}")
