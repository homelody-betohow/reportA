import csv
import chardet
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
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.Z_method.split_rows_data_拆分SKU_1个加号_逗号 import split_one_rows_data
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

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

product_map_sku_path = fr"{DESKTOP_ROOT}\广告-SKU关系对应.xlsx"  # 改成对应的映射表

# 先去除欧元符号和前后空格                     空值，替换成：0
otto_file_df['Ausgaben'] = otto_file_df['Ausgaben'].apply(
    lambda x: float(x.replace('€', '').replace(',', '.').strip()) if isinstance(x, str) and x.strip() else 0
)


# 筛选出 'Ausgaben' 列 中值不等于 0 的行（相对应删掉=0的行）
otto_file_df = otto_file_df[otto_file_df['Ausgaben'] != 0]

# 去除 整张表 的前后空格
for col in otto_file_df.columns:
    otto_file_df[col] = otto_file_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

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

# OTTO 有部分SKU是商品ID，要映射回 第一个 SKU
# SKU 是否以'25-'开头，分成两个 df
mask = otto_file_df_1['SKU'].str.startswith('25-')
df_25 = otto_file_df_1.loc[mask].copy()
df_other = otto_file_df_1.loc[~mask].copy()
# 商品ID 去映射 产品信息库 的 第一个 产品编码（SKU）
product_map_sku_path = r"\\Betohow\数据报表\数据库\产品信息库2025.xlsx"  # 改成对应的映射表
df_25_1 = sku_mappings(
    main_df=df_25,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="商品ID",
    map_new_sku="产品编码",
    map_sku_sheet='产品信息表'
)
# 重命名列
df_25_1 = df_25_1.rename(columns={'SKU': '原-SKU'})
df_25_1 = df_25_1.rename(columns={'映射产品编码': 'SKU'})
# --- 合并 ---
otto_file_df_1 = pd.concat([df_25_1, df_other]).sort_index()

# 在 SKU 后插入新列 儿子-站点识别码
new_column_name = "儿子-站点识别码"  # 新列名
new_column_data = "OTTO-BTH" + otto_file_df_1["SKU"]  # 新列数据，OTTO平台广告花费的站点都是：OTTO-BTH
target_column = "SKU"  # 目标列名（在其后插入）
insert_position = otto_file_df_1.columns.get_loc(target_column) + 1  # 计算插入位置
otto_file_df_1.insert(insert_position, new_column_name, new_column_data)  # 插入新列
# OTTO平台广告花费的站点都是：OTTO-BTH
otto_file_df_1['站点'] = "OTTO-BTH"

# 映射 平台
product_map_sku_path = fr"{DESKTOP_ROOT}\站点-匹配表.xlsx"  # 改成对应的映射表
otto_file_df_2 = sku_mappings(
    main_df=otto_file_df_1,
    main_sku='站点',
    map_sku_path=product_map_sku_path,
    map_old_sku="站点",
    map_new_sku="平台",
    map_sku_sheet='站点匹配'
)
# 在 儿子-站点识别码 后插入 儿子-平台识别码
new_column_name = "儿子-平台识别码"  # 新列名
new_column_data = otto_file_df_2["映射平台"] + otto_file_df_2["SKU"]  # 新列数据
target_column = "儿子-站点识别码"  # 目标列名（在其后插入）
insert_position = otto_file_df_2.columns.get_loc(target_column) + 1  # 计算插入位置
otto_file_df_2.insert(insert_position, new_column_name, new_column_data)  # 插入新列
# 保存目标列
otto_file_df_2 = otto_file_df_2[
    ['Artikelnummer', 'SKU', '站点', '映射平台', '儿子-站点识别码', '儿子-平台识别码', 'Ausgaben']]

# 更改列名，将’Ausgaben‘  改为 ’广告费(非AMZ)‘
otto_file_df_2 = otto_file_df_2.rename(columns={'Ausgaben': '广告费(非AMZ)'})

# 将处理后的数据保存到新的Excel文件
output_file_path = otto_file_path.rsplit('\\', 1)[0] + '\\(处理完成)OTTO广告.xlsx'
otto_file_df_2.to_excel(output_file_path, index=False)  # index=False表示不保存索引列

print(f"处理完成，结果已保存到{output_file_path}")
