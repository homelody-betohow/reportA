import pandas as pd
import importlib.util
from pathlib import Path
# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.Z_method.sku_映射 import sku_mappings
from A_报表.Z_method.platform_shop import map_region_to_platform, strip_lm_region_suffix
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT


def _is_blank(series: pd.Series) -> pd.Series:
    """空值判定：NaN / 空串 / 'nan' / 'None'。"""
    return series.isna() | series.astype(str).str.strip().isin(['', 'nan', 'None'])


# 仓租侧站点 / 库存「原-平台」标签 → 订单统计口径的标准平台（DB 未命中时的本地兜底）
# 与 K4 replace_site_with_rent 使用的平台名一致：MANO-EU / REAL / LM / AMAZON-EU …
_LOCAL_TO_PLATFORM: dict[str, str] = {
    "AMAZON-EU": "AMAZON-EU",
    "AMAZON-DE": "AMAZON-EU",
    "AMAZON-US": "AMAZON-EU",
    "MANO-EU": "MANO-EU",
    "MANO-FR": "MANO-EU",
    "REAL": "REAL",
    "REAL-FB": "REAL",
    "LM": "LM",
    "LM-BTH": "LM",
    "LM-TOTO": "LM",
    "LM-FR": "LM",
    "LM-BC": "LM",
    "LM-BC-ls": "LM",
    "LM-BC-xj": "LM",
    "OTTO": "OTTO",
    "OTTO-BTH": "OTTO",
    "DLZ-EU": "DLZ",
    "DLZ-DE": "DLZ",
    "castorama": "castorama",
    "CD": "chengyi-CD",
    "chengyi-CD": "chengyi-CD",
    "SHEIN": "SHEIN",
    "TEMU-AIH": "TEMU-AIH",
    "TEMU-BV": "TEMU-BV",
    "TEMU-HM": "TEMU-HM",
    "TEMU-AL": "TEMU-AL",
    "TEMU-KR-A": "TEMU-KR-A",
    "TEMU-KR-B": "TEMU-KR-B",
    "TEMU-KR-C": "TEMU-KR-C",
    "TEMU-HJ-A": "TEMU-HJ-A",
    "TEMU-HJ-B": "TEMU-HJ-B",
    "TEMU-HJ-C": "TEMU-HJ-C",
    "TEMU-NF-A": "TEMU-NF-A",
    "TEMU-NF-B": "TEMU-NF-B",
    "TEMU-NF-C": "TEMU-NF-C",
    "TEMU-BZ": "TEMU-BZ",
    "TEMU-AQ": "TEMU-AQ",
}


def _local_platform_lookup(keys: pd.Series) -> pd.Series:
    """本地字典查平台；LM 站点会先去掉 -ls/-xj 再查。"""
    normalized = keys.astype(str).str.strip().map(strip_lm_region_suffix)
    return normalized.map(_LOCAL_TO_PLATFORM)


def _fill_blank_platform(df: pd.DataFrame, fill_series: pd.Series) -> int:
    """只填「平台」空值，返回本次新填行数。"""
    empty = _is_blank(df['平台'])
    usable = empty & ~_is_blank(fill_series)
    df.loc[usable, '平台'] = fill_series.loc[usable]
    return int(usable.sum())

# 2个平台处理好的仓租用文件路径
# TODO 文件路径！！！
file_paths = [
    fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\4PX\(处理完成)4PX-仓租明细.xlsx',
    fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\鸿羽\(处理完成)hy-仓租明细.xlsx'
]
all_data = []
# 遍历文件路径列表
for file_path in file_paths:
    df = pd.read_excel(file_path)
    # 将读取的数据添加到列表中
    all_data.append(df)
    print(f"成功读取文件：{file_path}")

# 合并所有数据
merged_df = pd.concat(all_data, ignore_index=True)
# 重命名列
merged_df = merged_df.rename(columns={'SKU': '原-SKU'})
# 计算 LM-BC 的海外仓仓租费总值
if '原-平台' not in merged_df.columns:
    merged_df['原-平台'] = pd.NA
lm_bc_sum = merged_df[merged_df['原-平台'] == 'LM-BC']['海外仓仓租费'].sum()
merged_df['LM-BC的仓租'] = float('nan')  # 新建一列，默认填充 NaN
merged_df.loc[1, 'LM-BC的仓租'] = lm_bc_sum  # 在第二行（索引为1）填入总值

# 计算-一共需要分摊的仓租
all_fen_tan_cang_zu = merged_df['无平台-需要分摊的费用'].sum()  # 所有仓库 需要分摊的仓租

# 按照 '站点商品ID识别码' 列进行分组，并对 '仓租' 列进行汇总
# 空的'站点商品ID识别码'的数据会丢失，原-平台 == LM-BC 的数据会丢失
# 保留「原-平台」供后续平台兜底（K1/K2 库存标签，平台映射失败时仍有值）
result_df = merged_df.groupby('站点商品ID识别码').agg({
    '海外仓仓租费': 'sum',  # 求和
    '无平台-需要分摊的费用': 'sum',  # 求和
    '原-SKU': 'first',  # 保留每组的第一行数据
    '商品ID': 'first',  # 保留每组的第一行数据
    '站点': 'first',  # 保留每组的第一行数据
    '平台': 'first',  # 保留每组的第一行数据
    '原-平台': 'first',  # 库存侧平台标签，用于平台空值兜底
    '平台商品ID识别码': 'first'  # 保留每组的第一行数据
}).reset_index()

#  商品ID 去映射 产品信息库 的 第一个 产品编码（SKU）
product_map_sku_path = r"\\Betohow\数据报表\数据库\产品信息库2025.xlsx"  # 改成对应的映射表
result_df_1 = sku_mappings(
    main_df=result_df,
    main_sku='商品ID',
    map_sku_path=product_map_sku_path,
    map_old_sku="商品ID",
    map_new_sku="产品编码",
    map_sku_sheet='产品信息表'
)
# 重命名列
result_df_1 = result_df_1.rename(columns={'映射产品编码': 'SKU'})

result_df_1 = result_df_1[
    ['SKU', '商品ID', '站点', '平台', '原-平台', '站点商品ID识别码', '平台商品ID识别码', '海外仓仓租费']]

# ---------------------------------------------------------------------------
# 补全「平台」「平台商品ID识别码」
#   0) 若「平台」写成了站点名/库存标签（如 AMAZON-DE、MANO-FR、LM-BTH），
#      先标准化为订单统计口径（AMAZON-EU / MANO-EU / LM …）
#   兜底顺序（只填空）：
#     1) 站点 → platform_shop（DB）
#     2) 原-平台 → platform_shop（DB；K1/K2 库存标签）
#     3) 站点 / 原-平台 → 本地标准平台字典（DB 未配置时）
#     4) 仍空则回填原-平台原文（最后兜底，避免下游全空）
#   平台商品ID识别码 = 平台 + 商品ID（只补空）
# ---------------------------------------------------------------------------
# 0) 标准化：站点名误写入「平台」列
_std_from_platform = _local_platform_lookup(result_df_1['平台'])
_need_norm = (
    ~_is_blank(result_df_1['平台'])
    & _std_from_platform.notna()
    & (result_df_1['平台'].astype(str).str.strip() != _std_from_platform)
)
result_df_1.loc[_need_norm, '平台'] = _std_from_platform.loc[_need_norm]
print(f"平台标准化（站点名→标准平台）：{int(_need_norm.sum())} 行")

# 1) 站点 → DB
result_df_1 = map_region_to_platform(result_df_1, site_col='站点', platform_col='_映射平台_临时')
n1 = _fill_blank_platform(result_df_1, result_df_1['_映射平台_临时'])
result_df_1 = result_df_1.drop(columns=['_映射平台_临时'])

# 2) 原-平台 → DB
result_df_1 = map_region_to_platform(result_df_1, site_col='原-平台', platform_col='_映射平台_临时')
n2 = _fill_blank_platform(result_df_1, result_df_1['_映射平台_临时'])
result_df_1 = result_df_1.drop(columns=['_映射平台_临时'])

# 3) 本地字典：先试站点，再试原-平台
n3a = _fill_blank_platform(result_df_1, _local_platform_lookup(result_df_1['站点']))
n3b = _fill_blank_platform(result_df_1, _local_platform_lookup(result_df_1['原-平台']))

# 4) 仍空：用原-平台原文兜底
n4 = _fill_blank_platform(result_df_1, result_df_1['原-平台'])

# 5) 补全后再标准化一次（DB 可能返回站点名作 market_code）
_std_after = _local_platform_lookup(result_df_1['平台'])
_need_norm2 = (
    ~_is_blank(result_df_1['平台'])
    & _std_after.notna()
    & (result_df_1['平台'].astype(str).str.strip() != _std_after)
)
result_df_1.loc[_need_norm2, '平台'] = _std_after.loc[_need_norm2]

print(
    f"平台补全：站点→DB {n1} 行，原-平台→DB {n2} 行，"
    f"本地(站点) {n3a} 行，本地(原-平台) {n3b} 行，原-平台原文 {n4} 行；"
    f"再标准化 {int(_need_norm2.sum())} 行；"
    f"仍空 {int(_is_blank(result_df_1['平台']).sum())} 行"
)

# 输出不再保留「原-平台」
result_df_1 = result_df_1.drop(columns=['原-平台'])

# 有平台时统一重建识别码，保证与最终「平台」一致（含标准化后的行）
_can_build_plat_id = ~_is_blank(result_df_1['平台']) & result_df_1['商品ID'].notna()
result_df_1.loc[_can_build_plat_id, '平台商品ID识别码'] = (
    result_df_1.loc[_can_build_plat_id, '平台'].astype(str)
    + result_df_1.loc[_can_build_plat_id, '商品ID'].astype(str)
)
print(f"平台商品ID识别码重建：{int(_can_build_plat_id.sum())} 行")

# 筛选“站点”列不为空、不为空字符串、不等于“无”、不等于“其它”，海外仓仓租费 不等于 0
filtered_df = result_df_1[result_df_1['站点'].notna() & (result_df_1['站点'] != '') & (result_df_1['站点'] != '无') & (
        result_df_1['站点'] != '其他') & (result_df_1['海外仓仓租费'] != 0)]

filtered_df = filtered_df.copy()
# 若筛选后无数据（如短周期日报无符合条件的海外仓记录），补一行空行，
# 用于承载下面的两个汇总值，避免 iloc[0] 因空表越界报错
if filtered_df.empty:
    filtered_df = pd.DataFrame({col: [pd.NA] for col in filtered_df.columns})
# 新增空列：所有仓库-无平台-需要分摊的费用
filtered_df['所有仓库-无平台-需要分摊的费用'] = pd.NA
# 将 求和结果，放在新增列的第一个单元格
filtered_df.iloc[0, filtered_df.columns.get_loc('所有仓库-无平台-需要分摊的费用')] = all_fen_tan_cang_zu

# 新增空列：所有仓库-无平台-需要分摊的费用
filtered_df['LM-BC的仓租'] = pd.NA
# 将 求和结果，放在新增列的第一个单元格
filtered_df.iloc[0, filtered_df.columns.get_loc('LM-BC的仓租')] = lm_bc_sum

# 将所有结果写入
output_path = fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租' + '\\(处理完成)所有-海外仓-仓租明细.xlsx'
filtered_df.to_excel(output_path, index=False)
print(f'所有，平台仓租费用，文件合并成功，path：{output_path}')
