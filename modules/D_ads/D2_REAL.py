import os
import glob
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
from common.platform_shop import map_region_to_platform
from config.A0_set_date import shared_date, folder_name, kc_to_EUR
from common.split_rows_data_SKU import split_one_rows_data
from config.A0_paths import DESKTOP_ROOT


def _read_csv_lines(file_path):
    """按常见编码读取 REAL 导出 csv，避免 Kč 等表头因编码错误变成乱码。"""
    raw = Path(file_path).read_bytes()
    fallback = None
    for encoding in ('utf-8-sig', 'utf-8', 'cp1250', 'cp1252', 'latin-1'):
        try:
            text = raw.decode(encoding)
        except UnicodeDecodeError:
            continue
        lines = text.splitlines()
        if lines and 'Cost' in lines[0]:
            return lines
        if fallback is None:
            fallback = lines
    return fallback or []


def csv_to_df(file_path):
    # 同样的文件，编码要一致，不然不能正确合并
    lines = _read_csv_lines(file_path)
    new_data = []
    for line in lines:
        cells = line.strip().split(';')
        new_row = []
        for cell in cells:
            cell = cell.strip().replace('"', '').replace('\ufeff', '').replace(',', '.')
            try:
                new_row.append(int(cell))
            except ValueError:
                new_row.append(cell)
        new_data.append(new_row)

    site = ''
    file_name = os.path.basename(file_path)
    if 'REAL-DE-FB' in file_name:
        site = 'REAL-DE-FB'
    elif 'REAL-IT-FB' in file_name:
        site = 'REAL-IT-FB'
    elif 'REAL-CZ-FB' in file_name:
        site = 'REAL-CZ-FB'
    elif 'REAL-BTH' in file_name:
        site = 'REAL-BTH'

    real_file_df = pd.DataFrame(new_data[1:], columns=new_data[0])
    if site:
        real_file_df['站点'] = site
    else:
        print(f'无法获取到对应的站点，请检查文件名，程序终止！！！')
        exit()
    return real_file_df


def _find_cz_cost_col(columns):
    """定位 CZ 花费列：优先 Cost (Kč)，兼容编码差异或已是欧元的表头。"""
    cols = list(columns)
    if 'Cost (Kč)' in cols:
        return 'Cost (Kč)', 'kc'
    if 'Cost (€)' in cols:
        return 'Cost (€)', 'eur'
    for col in cols:
        col_str = str(col)
        lower = col_str.lower()
        if 'cost' in lower and ('kč' in lower or 'kc' in lower or 'czk' in lower):
            return col, 'kc'
        if 'cost' in lower and '€' in col_str:
            return col, 'eur'
    # 编码乱码时：仍按 Cost (...) 识别；含 euro/eur 视为欧元，否则按克朗换算
    for col in cols:
        col_str = str(col)
        if col_str.startswith('Cost (') and col_str.endswith(')'):
            inner = col_str[6:-1].lower()
            if 'eur' in inner or '€' in col_str:
                return col, 'eur'
            return col, 'kc'
    return None, None


# TODO 文件夹路径！！！
real_folder_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\广告\REAL"
# 获取文件夹中的所有的  当前日期.csv 文件
real_file_paths = glob.glob(os.path.join(real_folder_path, f'*{shared_date}.csv'))
# 过滤掉以 REAL-CZ-FB 开头的文件
real_file_paths = [f for f in real_file_paths if not os.path.basename(f).startswith('REAL-CZ-FB')]
print(real_file_paths)
if not real_file_paths:
    raise FileNotFoundError(f'未找到 REAL 广告 csv：{real_folder_path}\\*{shared_date}.csv')
# 读取并处理每个文件               合并 DataFrame
all_real_df_no_cz = pd.concat([csv_to_df(real_file_path) for real_file_path in real_file_paths], ignore_index=True)

# TODO 文件路径！！！
# DE、IT的广告花费是：欧元，CZ的广告花费是：捷克克朗（先转成RMB，再 / 7.3 转 欧元）；PL没有广告投入
# TODO 每月1号，手动更新 捷克克朗 转 RMB 的汇率！！！
real_cz_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\广告\REAL\REAL-CZ-FB-广告数据-{shared_date}.csv"
if os.path.isfile(real_cz_path):
    real_cz_df = csv_to_df(real_cz_path)
    cz_cost_col, cz_currency = _find_cz_cost_col(real_cz_df.columns)
    if cz_cost_col is None:
        raise KeyError(
            f"REAL-CZ-FB 未找到花费列（期望 Cost (Kč) 或 Cost (€)），实际列名：{real_cz_df.columns.tolist()}"
        )
    real_cz_df[cz_cost_col] = real_cz_df[cz_cost_col].astype(float)
    if cz_currency == 'kc':
        real_cz_df[cz_cost_col] = real_cz_df[cz_cost_col] * kc_to_EUR
        real_cz_df = real_cz_df.rename(columns={cz_cost_col: 'Cost (€)'})
        # 其余表头中的 Kč 一并换成 €，便于与其它站点对齐
        real_cz_df.columns = real_cz_df.columns.str.replace('Kč', '€', regex=False)
        real_cz_df.columns = real_cz_df.columns.str.replace('Kc', '€', regex=False)
    elif cz_cost_col != 'Cost (€)':
        real_cz_df = real_cz_df.rename(columns={cz_cost_col: 'Cost (€)'})
    all_real_df = pd.concat([all_real_df_no_cz, real_cz_df], ignore_index=True)
else:
    print(f'未找到 CZ 广告文件，跳过：{real_cz_path}')
    all_real_df = all_real_df_no_cz
# 去除 整张表 的前后空格
for col in all_real_df.columns:
    all_real_df[col] = all_real_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

# 删除"Cost (€)"列为  0.00、0 的行
all_real_df = all_real_df[~all_real_df['Cost (€)'].isin(['0.00', '0'])]
# 将列 'Cost (€)' 转换为 float 类型
all_real_df['Cost (€)'] = all_real_df['Cost (€)'].astype(float)

# print("all_real_df 的列名：", all_real_df.columns.tolist()) # 打印表头

product_map_sku_path = fr"{DESKTOP_ROOT}\广告-SKU关系对应.xlsx"  # 改成对应的映射表
#  映射sku（儿子）
# REAL的EAN的对应关系不用分站点，不同站点的相同EAN对应的儿子可能不同，但是爸爸相同
all_real_df_1 = sku_mappings(
    main_df=all_real_df,
    main_sku='EAN',
    map_sku_path=product_map_sku_path,
    map_old_sku="EAN",
    map_new_sku="仓库sku",
    map_sku_sheet='REAL ENA对应表'
)

all_real_df_1 = all_real_df_1.rename(columns={'映射仓库sku': 'SKU'})

output_file_path = real_file_paths[0].rsplit('\\', 1)[0] + '\\(已完成-1)REAL广告.xlsx'
all_real_df_1.to_excel(output_file_path, index=False)
print(f"处理完成，结果已保存到{output_file_path}")

#  拆分有 “+” 的sku
all_real_df_2 = split_one_rows_data(
    input_df=all_real_df_1,
    data_column='SKU',
    value_column='Cost (€)'
)

# SKU-站点识别码
new_column_name = "SKU-站点识别码"  # 新列名
new_column_data = all_real_df_2["站点"] + all_real_df_2["SKU"]  # 新列数据
target_column = "SKU"  # 目标列名（在其后插入）
insert_position = all_real_df_2.columns.get_loc(target_column) + 1  # 计算插入位置
all_real_df_2.insert(insert_position, new_column_name, new_column_data)  # 插入新列
# 映射 平台（数据源：platform_shop）
all_real_df_3 = map_region_to_platform(all_real_df_2, site_col='站点')
# 在 SKU-站点识别码 后插入 SKU-平台识别码
new_column_name = "SKU-平台识别码"  # 新列名
new_column_data = all_real_df_3["映射平台"] + all_real_df_3["SKU"]  # 新列数据
target_column = "SKU-站点识别码"  # 目标列名（在其后插入）
insert_position = all_real_df_3.columns.get_loc(target_column) + 1  # 计算插入位置
all_real_df_3.insert(insert_position, new_column_name, new_column_data)  # 插入新列
# 保存目标列
all_real_df_4 = all_real_df_3[
    ['EAN', 'SKU', '站点', '映射平台', 'SKU-站点识别码', 'SKU-平台识别码', 'Cost (€)']]
# 更改列名，将’Cost (€)‘  改为 ’广告费(非AMZ)‘
all_real_df_4 = all_real_df_4.rename(columns={'Cost (€)': '广告费(非AMZ)'})

# 将处理后的数据保存到新的Excel文件
output_file_path = real_file_paths[0].rsplit('\\', 1)[0] + '\\(处理完成)REAL广告.xlsx'
all_real_df_4.to_excel(output_file_path, index=False)  # index=False表示不保存索引列
print(f"处理完成，结果已保存到{output_file_path}")
