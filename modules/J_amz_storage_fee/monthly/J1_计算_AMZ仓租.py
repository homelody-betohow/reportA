"""
J1：计算 AMZ / FBA 仓租明细（月度）

流程概览：
  1. 读取上月「FBA仓租明细」Excel（表头可能被元数据顶到下方，需自动定位）
  2. 清洗仓库 SKU → 统一为 product_sku 形态
  3. 查库 product_sku → product_uid（商品ID），含 -NW 特殊处理
  4. 按店铺映射站点 / 平台，拼识别码
  5. 汇总 FBA仓租费，过滤无效行，输出「(已完成-1)」结果文件
"""
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

# product_sku 表：用清洗后的 SKU 反查商品ID（product_uid）
PRODUCT_SKU_TABLE = "product_sku"
# IN 查询分批大小，避免 SQL 参数过多
_KEY_CHUNK = 200

# 忽略 openpyxl 读 Excel 时的无关 UserWarning
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
print(f"{fba_date}")

# ---------------------------------------------------------------------------
# 输入文件：桌面「仓租\FBA仓租」目录下的上月 FBA 仓租明细
# 实际数据源多为 SellerSku 利润报表，按约定重命名后即可被本脚本读取
# ---------------------------------------------------------------------------
# TODO 文件路径！！！   上月的 利润报表
# main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\FBA仓租\FBA仓租明细{fba_date}.xlsx"
# FBA仓租 引用的是 SellerSku利润报表， 将 对应日期文件 重命名 即可
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\FBA仓租\FBA仓租明细{fba_date}.xlsx"


def _detect_header_row(path: str, max_scan: int = 20) -> int:
    """
    定位 Excel 真实表头行号。

    自动化下载的 FBA 明细前几行常是日期、币种等元数据，
    真正列名（含 sellerSku）可能不在第 0 行，故扫描前 max_scan 行。
    """
    preview = pd.read_excel(path, header=None, nrows=max_scan)
    for i, row in preview.iterrows():
        vals = {str(v).strip() for v in row.tolist() if pd.notna(v)}
        if "sellerSku" in vals:
            return int(i)
    raise ValueError(f"未在前 {max_scan} 行找到表头 sellerSku：{path}")


# 读入主表（按检测到的表头行）
_header_row = _detect_header_row(main_file_path)
main_file_df = pd.read_excel(main_file_path, header=_header_row)
print(f"[Excel] 读取 {len(main_file_df)} 行（header={_header_row}）")

# UK 站点仓租不纳入本流程，整行剔除
main_file_df = main_file_df[main_file_df['站点'] != 'UK']

# 仓库sku 为空时用 sellerSku 兜底，保证后续清洗与映射有可用键
main_file_df['仓库sku'] = main_file_df['仓库sku'].fillna(main_file_df['sellerSku'])


def extract_values(s):
    """
    从原始「仓库sku」字符串中抽出可用于对齐 product_sku 的编码。

    规则：
      - 含 amzn.gr.：取其后第一段（再按 - / _ 截断）
      - 否则：按 #、BCFBAFL、FBFBAFL 等后缀标记截断，取前缀
    """
    if pd.isna(s):  # 检查是否为 NaN
        return None  # 如果是 NaN，返回 None 或其他默认值
    if 'amzn.gr.' in s:
        return s.split(r'amzn.gr.')[-1].split('-')[0].split('_')[0]
    else:
        return s.split('#')[0].split('BCFBAFL')[0].split('FBFBAFL')[0]


# 清洗仓库sku，并改名为 SKU（后续与 product_sku 表对齐）
main_file_df['仓库sku'] = main_file_df['仓库sku'].apply(extract_values)
main_file_df = main_file_df.rename(columns={'仓库sku': 'SKU'})


def _fetch_product_uid_map(skus: list[str]) -> dict[str, str]:
    """
    批量查询：product_sku → product_uid（商品ID）。

    仅取未删除、且 product_uid 非空的记录；按 _KEY_CHUNK 分批 IN 查询。
    """
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
    """
    主表 SKU → 商品ID（product_uid）。

    - 查库前去掉末尾 -NW，用无 -NW 的 SKU 去匹配
    - 原 SKU 带 -NW 且映射成功时，商品ID 再缀回 -NW（区分 NW 变体）
    - 无效 / 未命中的商品ID 置空
    - 「商品ID」插在 SKU 列右侧
    """
    out = main_df.copy()
    if main_sku not in out.columns:
        raise KeyError(f"主表缺少列 {main_sku!r}，当前列: {list(out.columns)}")

    series = out[main_sku].astype(str).str.strip()
    # 空串 / 字面量 nan 等视为无效 SKU，不参与查库
    invalid = series.isin(("", "nan", "None", "NaN")) | out[main_sku].isna()
    # 带 -NW 后缀的行：查库用去掉 -NW 的键，映射后再把 -NW 写回商品ID
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
    """将 NA / 空白 / 字面量 nan、None、NaN 统一视为「空」。"""
    as_str = series.astype(object).map(lambda v: "" if pd.isna(v) else str(v).strip())
    return as_str.eq("") | as_str.isin(("nan", "None", "NaN"))


def warn_blank_product_uid(df: pd.DataFrame) -> None:
    """打印商品ID 为空的行数及最多 10 行样例，便于人工核对未映射 SKU。"""
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


# ---------------------------------------------------------------------------
# 主流程：映射商品ID → 站点/平台 → 识别码 → 仓租合计 → 导出
# ---------------------------------------------------------------------------

# SKU → 商品ID（product_sku.product_sku → product_uid）
main_file_df_1 = map_sku_to_product_uid(main_file_df, main_sku="SKU")
warn_blank_product_uid(main_file_df_1)

# 按「店铺」映射「映射站点」「映射平台」（数据源：platform_shop）
main_file_df_3 = map_shop_platform_region(main_file_df_1, shop_col='店铺', site_col=None)

# 站点商品ID识别码 = 映射站点 + 商品ID（紧挨「映射站点」列插入）
new_column_name = "站点商品ID识别码"
new_column_data = main_file_df_3["映射站点"] + main_file_df_3["商品ID"]
target_column = "映射站点"
insert_position = main_file_df_3.columns.get_loc(target_column) + 1
main_file_df_3.insert(insert_position, new_column_name, new_column_data)

# 平台商品ID识别码 = 映射平台 + 商品ID（紧挨「站点商品ID识别码」列插入）
new_column_name = "平台商品ID识别码"
new_column_data = main_file_df_3["映射平台"] + main_file_df_3["商品ID"]
target_column = "站点商品ID识别码"
insert_position = main_file_df_3.columns.get_loc(target_column) + 1
main_file_df_3.insert(insert_position, new_column_name, new_column_data)

# FBA仓租费 = 仓储费用（已分摊）+ 长期仓储费（已分摊）；空填 0，取绝对值去掉负号
main_file_df_3['FBA仓租费'] = main_file_df_3['仓储费用（已分摊）'] + main_file_df_3['长期仓储费（已分摊）']
main_file_df_3['FBA仓租费'] = main_file_df_3['FBA仓租费'].fillna(0).abs()  # 去掉 负号

# 输出前过滤：无 sellerSku、或仓租费为 0 的行不进入结果表
main_file_df_4 = main_file_df_3.dropna(subset=['sellerSku'])
main_file_df_4 = main_file_df_4[main_file_df_4['FBA仓租费'] != 0]

# 只保留下游需要的列
main_file_df_4 = main_file_df_4[
    ['sellerSku', 'ASIN', '产品信息', 'SKU','商品ID', '店铺', '映射站点', '映射平台', '站点商品ID识别码',
     '平台商品ID识别码', '仓储费用（已分摊）', '长期仓储费（已分摊）', 'FBA仓租费']]

# 同目录输出，文件名前缀「(已完成-1)」标记本步骤已完成
output_path = main_file_path.rsplit('\\', 1)[0] + '\\(已完成-1)' + main_file_path.rsplit('\\', 1)[1]
main_file_df_4.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')

# 汇总被过滤掉的「sellerSku 为空」行的仓租合计，便于对账核对
empty_sellerSku_df = main_file_df_3[main_file_df_3['sellerSku'].isna()]
empty_sellerSku_FBA = empty_sellerSku_df['FBA仓租费'].sum()
print(f'\n---------------------sellerSku为空的FBA仓租费是：{empty_sellerSku_FBA}EUR------------------------------')
