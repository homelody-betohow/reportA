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
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-6)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)
# 指定列名（需要筛选的列和需要提取的列）
columns_to_extract = ['SKU', '站点', '儿子-站点识别码', '销量']  # 替换为你需要提取的列名

# TODO 文件路径！！！
file_path = fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\广告\DLZ\(处理完成)DLZ-总的-广告数据.xlsx'
df = pd.read_excel(file_path, sheet_name='独立站-总的-广告数据', skipfooter=2)  # 跳过最后两行
# 获取 站点、费用
site_list = df['站点'].tolist()
site_cost_USD_list = df['需要摊分花费（美元）'].tolist()

# 创建一个空的 DataFrame，用于存储最终结果
result_df = pd.DataFrame()

for site, site_cost_USD in zip(site_list, site_cost_USD_list):
    # 在 main_file_df 中筛选出当前站点的数据
    main_site_df = main_file_df[main_file_df['站点'] == site].copy()  # 使用 copy() 确保是副本
    # 如果筛选结果为空，记录该站点信息
    if main_site_df.empty:
        print(f'站点：{site}，无sku，请检查！！！！！！！！！！！！')
        # 创建一个空的 DataFrame，记录无销量的站点信息
        empty_site_df = pd.DataFrame({
            '无销量的站点': [site],
            '需要摊分花费（美元）': [site_cost_USD]
        })
        # 将无销量站点信息追加到最终结果 DataFrame 中
        result_df = pd.concat([result_df, empty_site_df], ignore_index=True)
        continue
    # 计算当前站点指定列的总和
    total_value = main_site_df['销量'].sum()
    # 计算每行的“费用”
    main_site_df.loc[:, '费用'] = (main_site_df['销量'] / total_value) * site_cost_USD

    # 提取需要的列，并加上“费用”列
    main_site_df = main_site_df[columns_to_extract + ['费用']]
    # 对“费用”列保留两位小数
    main_site_df['费用'] = main_site_df['费用'].round(2)

    # 将筛选结果追加到最终结果 DataFrame 中
    result_df = pd.concat([result_df, main_site_df], ignore_index=True)

# 将结果保存到一个新的 Sheet 中
sheet_name = '分摊明细-美元'
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    result_df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"操作完成，file_path：{file_path}，sheet_name：{sheet_name}")