import numpy as np
import pandas as pd
from datetime import datetime
import importlib.util
from pathlib import Path
# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.A0_设置_时间段.A0_set_date import folder_name, shared_date
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-18)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)
# 删除 仓租 中间计算用的列
main_file_df.drop(columns=['原-海外仓仓租费', '所有仓库-无平台-需要分摊的费用', '仓租分摊'], inplace=True)

yue_zu_site_path = fr"{DESKTOP_ROOT}\月租总摊分.xlsx"
yue_zu_site_df = pd.read_excel(yue_zu_site_path, sheet_name='总的-月租')
# 创建映射字典
mapping_dict = dict(zip(yue_zu_site_df['站点'], yue_zu_site_df['原-站点']))

# 映射站点到原-站点，没有映射的保留空值
main_file_df['月租-站点识别'] = main_file_df['站点'].map(mapping_dict)

# 定义月租的字典，月租是固定的，按照“月租-站点识别”去摊分月租
# 这里的是 '原-站点'
yue_zu_dict = {
    "MANO-DE-OHPAMF": 11.25,
    "MANO-DE-OHPA": 110,
    "MANO-FR-OHPAMF": 11.25,
    "MANO-FR-OHPA": 110,
    "MANO-FR-OHPA-B2B": 110,
    "MANO-FR-BTHMF": 45,
    "MANO-ES-OHPAMF": 11.25,
    "MANO-IT-OHPAMF": 11.25,
    "MANO-IT-OHPA": 110,
    "DLZ-ES": 85.5416,
    "DLZ-FR": 85.5416,
    "DLZ-DE": 162.8124,
    "DLZ-IT": 78.1908,
    "DLZ-US": 78.1908,
    "REAL-BTH": 40,
    "REAL-FB": 40,
    "LM-TOTO": 39,
    "LM-BTH": 39,
    "LM-BC": 39,
    "LM-RP": 39,
    "OTTO-BTH": 99.9,
    "castorama": 39
}

# 按“月租-站点识别”分组，计算每个站点的销量总
grouped_key = main_file_df.groupby('月租-站点识别')['销量'].sum().reset_index()
grouped_key.rename(columns={'销量': '总销量_per_月租'}, inplace=True)
# 将分组结果与原始数据合并
main_file_df = main_file_df.merge(grouped_key, on='月租-站点识别', how='left')
# 计算每行的占比（针对每个月租-站点识别）；总销量为 0 时占比为 0，避免 0/0 产生 NaN
_total_sales = main_file_df['总销量_per_月租'].fillna(0)
main_file_df['占比_per_月租'] = np.where(
    (_total_sales > 0) & main_file_df['月租-站点识别'].notna(),
    main_file_df['销量'] / _total_sales,
    0,
)
# 计算摊分的费用
main_file_df['月租'] = np.round(main_file_df.apply(
    lambda row: yue_zu_dict.get(row['月租-站点识别'], 0) * row['占比_per_月租'] if pd.notna(
        row['月租-站点识别']) else 0,
    axis=1
), 2)

# 删除中间计算用的列
main_file_df.drop(columns=['总销量_per_月租', '占比_per_月租'], inplace=True)

if folder_name == '日报':
    # 获取 当前年份
    current_year = datetime.now().year
    # 获取 日报月份
    month = int(shared_date.split('.')[0])
    # 获取 日报 所在月的天数（当前年）
    days_in_month = pd.Timestamp(year=current_year, month=month, day=1).daysinmonth
    # 获取 当前日报的天数
    d1, d2 = (int(part.split('.')[1]) for part in shared_date.split('-'))
    days = d2 - d1 + 1
    # 重命名
    main_file_df = main_file_df.rename(columns={'月租': '月租-整月'})
    # 得到 单日 的月租
    main_file_df['月租'] = np.round(
        main_file_df['月租-整月'].fillna(0) / days_in_month * days, 2
    )

# 保存结果到新的 Excel 文件
output_path = main_file_path.replace('已完成-18', '已完成-19')
main_file_df.to_excel(output_path, index=False)
print(f"结果已保存到 {output_path}")
