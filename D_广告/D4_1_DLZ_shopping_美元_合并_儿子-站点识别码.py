import os
import csv
import glob
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

from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.Z_method.split_rows_data_拆分SKU_1个加号_逗号 import split_one_rows_data
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT


def dlz_csv_to_df(file_path):
    with open(file_path, 'rb') as file:
        raw_data = file.read()
        result = chardet.detect(raw_data)
    # 检测文件编码
    encoding = result['encoding']
    print(f"文件的编码是: {encoding}")

    # 读取并处理文件内容
    with open(file_path, encoding=encoding) as f:
        # 跳过前两行
        for _ in range(2):
            next(f)
        # 创建TSV阅读器          分割规则：\t（制表符）
        reader = csv.reader(f, delimiter='\t')
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
        file_df = pd.DataFrame(new_data[1:], columns=new_data[0])

        site = ''
        if 'de-shopping广告' in file_path.rsplit('\\', 1)[-1]:
            site = 'DLZ-DE'
        elif 'es-shopping广告' in file_path.rsplit('\\', 1)[-1]:
            site = 'DLZ-ES'

        if site:
            file_df['站点'] = site
            return file_df
        else:
            print(f'无法获取到对应的站点，请检查文件名，file_path：{file_path}，程序终止！！！')
            exit()


# TODO 文件夹路径！！！
dlz_folder_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\广告\DLZ"  # 替换为你的文件夹路径
# 获取文件夹中的所有 ‘shopping广告-当前日期’ 文件
dlz_file_paths = glob.glob(os.path.join(dlz_folder_path, f'*shopping广告-{shared_date}*'))
# 读取所有文件并自动设置站点
dlz_df_list = [dlz_csv_to_df(DLZ_file_path) for DLZ_file_path in dlz_file_paths]

# 合并所有 DataFrame
all_dlz_df = pd.concat(dlz_df_list, ignore_index=True)

# 筛选 "费用" 列既不为空也不为 0 的行
all_dlz_df = all_dlz_df[all_dlz_df["费用"].notna() & (all_dlz_df["费用"] != 0)]
# print("all_dlz_df 的列名：", all_dlz_df.columns.tolist()) # 打印表头

# 将 '产品 ID' 列中的所有值转换为大写
all_dlz_df['产品 ID'] = all_dlz_df['产品 ID'].str.upper()
# 将处理后的数据保存到新的Excel文件
output_file_path = dlz_file_paths[0].rsplit('\\', 1)[0] + '\\(已完成-1)DLZ广告-shopping.xlsx'
all_dlz_df.to_excel(output_file_path, index=False)  # index=False表示不保存索引列
print(f"处理完成，结果已保存到{output_file_path}")

#  拆分有“+”的sku
all_dlz_df_1 = split_one_rows_data(
    input_df=all_dlz_df,
    data_column='产品 ID',
    value_column='费用'
)

# 将处理后的数据保存到新的Excel文件
output_file_path = dlz_file_paths[0].rsplit('\\', 1)[0] + '\\(已完成-2)DLZ广告-shopping.xlsx'
all_dlz_df_1.to_excel(output_file_path, index=False)  # index=False表示不保存索引列
print(f"处理完成，结果已保存到{output_file_path}")

all_dlz_df_1 = all_dlz_df_1.rename(columns={'产品 ID': 'SKU'})

# 去除 整张表 的前后空格
for col in all_dlz_df_1.columns:
    all_dlz_df_1[col] = all_dlz_df_1[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

new_column_name = "儿子-站点识别码"  # 新列名
new_column_data = all_dlz_df_1["站点"] + all_dlz_df_1["SKU"]  # 新列数据
target_column = "SKU"  # 目标列名（在其后插入）
insert_position = all_dlz_df_1.columns.get_loc(target_column) + 1  # 计算插入位置
all_dlz_df_1.insert(insert_position, new_column_name, new_column_data)  # 插入新列

# 按照 '儿子-站点识别码' 列进行分组，并对 '费用' 列进行汇总
grouped_df = all_dlz_df_1.groupby('儿子-站点识别码').agg({
    '费用': 'sum',  # 汇总 费用
    'SKU': 'first',  # 保留每组的第一行数据
    '站点': 'first',  # 保留每组的第一行数据
}).reset_index()

# 保存目标列（shopping 报表无「商家 ID」列）
grouped_df = grouped_df[['SKU', '儿子-站点识别码', '费用', '站点']]

# 将处理后的数据保存到新的Excel文件
output_file_path = dlz_file_paths[0].rsplit('\\', 1)[0] + '\\(处理完成)DLZ广告-shopping-美元.xlsx'
grouped_df.to_excel(output_file_path, index=False)  # index=False表示不保存索引列

print(f"处理完成，结果已保存到{output_file_path}")
