import warnings
import pandas as pd
import importlib.util
from pathlib import Path
# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.sku_mapping import sku_mappings
from common.cang_zu_site import map_platform_to_site
from common.platform_shop import map_region_to_platform
from config.A0_set_date import shared_date, folder_name, ku_cun_date
from config.A0_paths import DESKTOP_ROOT
from common.style import Color

# 忽略特定的警告
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")

# TODO 文件路径！！！
px4_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\4PX\4PX法国仓-仓租明细-{shared_date}.xlsx"
output_dir = px4_path.rsplit("\\", 1)[0]
px4_df = pd.read_excel(px4_path)

# 无明细或仓租为 0 时继续后续流程（不再因行数不足中断）
row_count = len(px4_df)
if row_count == 0:
    print(f"提示：4PX 仓租表无数据行（仓租为 0 或未产生明细），继续后续流程：{px4_path}")
    px4_df = pd.DataFrame(columns=["SKU", "应收金额"])
elif row_count < 2:
    print(f"提示：4PX 仓租表仅 {row_count} 行，继续处理：{px4_path}")

# 按 SKU 汇总应收金额
if px4_df.empty:
    px4_df = pd.DataFrame(columns=["SKU", "应收金额"])
else:
    missing_cols = [c for c in ("SKU", "应收金额") if c not in px4_df.columns]
    if missing_cols:
        raise ValueError(f"4PX 仓租表缺少列 {missing_cols}：{px4_path}")
    px4_df = px4_df.groupby("SKU", as_index=False)["应收金额"].sum()
px4_df = px4_df.rename(columns={'应收金额': '总仓租'})
px4_df['总仓租'] = px4_df['总仓租']
px4_all_cang_zu = px4_df['总仓租'].sum()  # HY 总的仓租费用

px4_df = px4_df.rename(columns={'SKU': 'SKU编码'})  # 为了后面的SKU列进行合并

# 映射 商品ID
product_map_sku_path = r"\\Betohow\数据报表\数据库\产品信息库2025.xlsx"  # 改成对应的映射表
px4_df_1 = sku_mappings(
    main_df=px4_df,
    main_sku='SKU编码',
    map_sku_path=product_map_sku_path,
    map_old_sku="产品编码",
    map_new_sku="商品ID",
    map_sku_sheet='产品信息表'
).copy()  # 显式创建副本

px4_df_1 = px4_df_1.rename(columns={'映射商品ID': '商品ID'})  # 为了后续合并

# 保存文件（步骤 1）
output_file_path = output_dir + "\\(已完成-1)4PX-仓租明细.xlsx"
px4_df_1.to_excel(output_file_path, index=False)
print(f"处理完成，结果已保存到{output_file_path}")

# 无仓租明细时，输出空表供 K3 合并，跳过分摊逻辑
PX4_FINAL_COLUMNS = [
    "SKU",
    "商品ID",
    "平台",
    "海外仓仓租费",
    "无平台-需要分摊的费用",
    "站点",
    "站点商品ID识别码",
    "原-平台",
    "平台商品ID识别码",
]
if px4_df_1.empty:
    print(f"{Color.RED}提示：4PX 无 SKU 明细，生成分摊结果空表并结束本脚本{Color.RESET}")
    result_DF_1 = pd.DataFrame(columns=PX4_FINAL_COLUMNS)
    output_file_path = output_dir + "\\(处理完成)4PX-仓租明细.xlsx"
    result_DF_1.to_excel(output_file_path, index=False)
    print(f"处理完成，结果已保存到{output_file_path}")
    raise SystemExit(0)

# 读取库存周转明细
# TODO 文件路径！！！
ku_cun_path = fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租\{ku_cun_date}库存周转明细.xlsx'
ku_cun_df = pd.read_excel(ku_cun_path, sheet_name='各平台SKU库存周转明细')

# 筛选 平台 列中只包含 'LM' 和 'MANO-EU' 的行
ku_cun_filtered_df = ku_cun_df[
    ku_cun_df['平台'].isin(['LM-BTH', 'LM-TOTO', 'LM-BC-ls', 'LM-BC-xj', 'MANO-EU'])].copy()  # 显式创建副本
# 计算每个 商品ID 的总数量
sku_total_quantity = ku_cun_filtered_df.groupby('商品ID')['在库（可调拨）'].transform('sum')
# 计算每个SKU在每个平台的数量占比
ku_cun_filtered_df['数量占比'] = ku_cun_filtered_df['在库（可调拨）'] / sku_total_quantity

# 合并px4_df_1和ku_cun_filtered_df，根据SKU进行合并
DF = pd.merge(px4_df_1, ku_cun_filtered_df, on='商品ID', how='left')
# 计算每个SKU在每个平台的仓租
DF['仓租'] = DF['总仓租'] * DF['数量占比']

# 选择需要的列生成新的DataFrame
result_DF = DF[['SKU编码', '商品ID', '平台', '仓租']].copy()  # 显式创建副本

px4_have_site_cang_zu = result_DF["仓租"].fillna(0).sum()  # 4px 有平台（站点）的仓租
# 需要分摊没有平台（站点）的仓租
px4_no_site_fen_tan = px4_all_cang_zu - px4_have_site_cang_zu
result_DF["无平台-需要分摊的费用"] = None
if len(result_DF) > 0:
    result_DF.at[0, "无平台-需要分摊的费用"] = px4_no_site_fen_tan

# 映射 站点（原桌面「仓租-站点映射.xlsx」→ cang_zu_site.PLATFORM_TO_SITE）
result_DF_1 = map_platform_to_site(result_DF, platform_col='平台')
# 在 映射站点 后插入新列 站点商品ID识别码
new_column_name = "站点商品ID识别码"  # 新列名
new_column_data = result_DF_1["映射站点"] + result_DF_1["商品ID"]  # 新列数据
target_column = "映射站点"  # 目标列名（在其后插入）
insert_position = result_DF_1.columns.get_loc(target_column) + 1  # 计算插入位置
result_DF_1.insert(insert_position, new_column_name, new_column_data)  # 插入新列

# 映射 平台 （原表的 平台列 里面有站点！数据源：platform_shop）
result_DF_1 = map_region_to_platform(result_DF_1, site_col='平台')
# 重命名
result_DF_1 = result_DF_1.rename(columns={'平台': '原-平台'})
result_DF_1 = result_DF_1.rename(columns={'映射平台': '平台'})
# 在 站点商品ID识别码 后插入 平台商品ID识别码
new_column_name = "平台商品ID识别码"  # 新列名
new_column_data = result_DF_1["平台"] + result_DF_1["商品ID"]  # 新列数据
target_column = "站点商品ID识别码"  # 目标列名（在其后插入）
insert_position = result_DF_1.columns.get_loc(target_column) + 1  # 计算插入位置
result_DF_1.insert(insert_position, new_column_name, new_column_data)  # 插入新列

result_DF_1 = result_DF_1.rename(columns={'映射站点': '站点'})
result_DF_1 = result_DF_1.rename(columns={'仓租': '海外仓仓租费'})
result_DF_1 = result_DF_1.rename(columns={'SKU编码': 'SKU'})
# 保存文件
output_file_path = output_dir + "\\(处理完成)4PX-仓租明细.xlsx"
result_DF_1.to_excel(output_file_path, index=False)
print(f"处理完成，结果已保存到{output_file_path}")
