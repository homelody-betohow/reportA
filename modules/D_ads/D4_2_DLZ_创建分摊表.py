import os
import csv
import glob
import chardet
import pandas as pd
from openpyxl import Workbook
import importlib.util
from pathlib import Path
# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from config.A0_set_date import shared_date, folder_name
from config.A0_paths import DESKTOP_ROOT


def csv_to_site_df(file_path):
    with open(file_path, 'rb') as file:
        raw_data = file.read()
        result = chardet.detect(raw_data)
    # 检测文件编码
    # encoding = result['encoding']
    encoding = 'utf-16le'   # 暂时 指定编码！
    print(f"文件的编码是: {encoding}")

    # 读取并处理文件内容
    with open(file_path, encoding=encoding) as f:
        # 跳过前两行
        for _ in range(2):
            next(f)
        # 创建TSV阅读器          分割规则：\t（制表符）
        reader = csv.reader(f, delimiter='\t')
        lines = list(reader)
        index_cost = lines[0].index("费用")
        index_out = lines[0].index("转化价值")
        # 站点 总的广告花费
        site_all_cost = [line[index_cost].replace(',', '') for line in lines if '总计：账号' in line][0]
        # 站点 总的产出
        site_all_out = [line[index_out].replace(',', '') for line in lines if '总计：账号' in line][0]

    file_name_to_site_dict = {
        "de-所有广告": "DLZ-DE",
        "es-所有广告": "DLZ-ES",
        "fr-所有广告": "DLZ-FR",
    }
    site = ''
    site_list = []
    for key in file_name_to_site_dict.keys():
        if key in file_path.rsplit('\\', 1)[-1]:
            site = file_name_to_site_dict[key]
            ROI = round(float(site_all_out) / float(site_all_cost), 2)  # 保留2位小数
            site_list = [site, site_all_cost, site_all_out, ROI]

    if site and site_list:
        return site_list
    else:
        print(f'无法获取到对应的站点，请检查文件名，file_path：{file_path}，程序终止！！！')
        exit()

# TODO 文件路径！！！
shopping_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\广告\DLZ\(处理完成)DLZ广告-shopping-美元.xlsx"
# TODO 文件夹路径！！！
dlz_folder_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\广告\DLZ"  # 替换为你的文件夹路径
# 获取文件夹中的所有 ‘所有广告’ 文件
dlz_file_paths = glob.glob(os.path.join(dlz_folder_path, '*所有广告*'))
site_data_list = [csv_to_site_df(dlz_file_path) for dlz_file_path in dlz_file_paths]
print(f'总的：{site_data_list}')

# 分摊
shopping_file_df = pd.read_excel(shopping_file_path)
all_cost = 0
for site_data in site_data_list:
    dlz_site = site_data[0]
    dlz_site_all_cost = float(site_data[1])
    all_cost += dlz_site_all_cost
    current_site_df = shopping_file_df[shopping_file_df['站点'] == dlz_site]
    # 计算当前站点指定列的总和
    shopping_site_cost = float(current_site_df['费用'].sum())
    print(f'dlz_site：{dlz_site}')
    print(f'dlz_site_all_cost：{dlz_site_all_cost}')
    print(f'shopping_site_cost：{shopping_site_cost}')
    print(f'{'-' * 50}')
    dlz_site_last_cost = round(dlz_site_all_cost - shopping_site_cost, 2)  # 保留2位小数，计算分摊费用
    site_data[1] = dlz_site_last_cost  # 将索引1的总花费替换成 剩余分摊的费用
print(f'分摊的：{site_data_list}')

# 总计
all_last_cost = 0
all_out = 0
for site_data in site_data_list:
    last_site_cost = float(site_data[1])
    all_last_cost += last_site_cost
    site_out = float(site_data[2])
    all_out += site_out
all_out = round(all_out, 2)  # 保留2位小数
all_ROI = round(all_out / all_last_cost, 2)
site_data_list.append(['总计：', all_last_cost, all_out, all_ROI])
# 实际（总的广告花费-美元）
shi_ji_ROI = round(all_out / all_cost, 2)
site_data_list.append(['实际（总的广告花费-美元）：', all_cost, all_out, shi_ji_ROI])

# 插入表头
head_list = ['站点', '需要摊分花费（美元）', '产出（美元）', 'ROI']
site_data_list.insert(0, head_list)
print(f'最终的：{site_data_list}')

# 将数据写excel表格 .xlsx
# 创建一个新的工作簿
wb = Workbook()

# 移除默认创建的sheet
wb.remove(wb.active)

# 创建一个新的sheet，并指定名字
sheet_name = "独立站-总的-广告数据"
ws = wb.create_sheet(sheet_name)

# 将二维列表写入sheet
for row in site_data_list:
    ws.append(row)

# 保存工作簿到文件
output_file_path = dlz_file_paths[0].rsplit('\\', 1)[0] + '\\(处理完成)DLZ-总的-广告数据.xlsx'
wb.save(output_file_path)

print(f"数据已成功写入成功， file_path：{output_file_path}")
