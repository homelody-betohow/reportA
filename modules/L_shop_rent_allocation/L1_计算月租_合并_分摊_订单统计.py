import numpy as np
import pandas as pd
from datetime import datetime
import importlib.util
from pathlib import Path
# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from config.A0_set_date import folder_name, shared_date
from config.A0_paths import DESKTOP_ROOT
from common.platform_shop import fetch_rent_by_region, fetch_rent_region_keys, map_site_to_rent_region
from common.style import Color

main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-18)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)
# 删除 仓租 中间计算用的列
main_file_df.drop(columns=['原-海外仓仓租费', '所有仓库-无平台-需要分摊的费用', '仓租分摊'], inplace=True)

# 月租摊分：数据源 platform_shop（market_region + store_fees，替代原桌面「月租总摊分.xlsx」）
rent_regions = fetch_rent_region_keys()
yue_zu_dict = fetch_rent_by_region()
main_file_df['月租-站点识别'] = map_site_to_rent_region(main_file_df['站点'], rent_regions)

if not yue_zu_dict:
    print(
        f"{Color.YELLOW}[WARN] platform_shop.store_fees 为空，月租将全部为 0；"
        f"请先执行 database/alter/update_platform_shop_store_fees.sql{Color.RESET}"
    )
else:
    print(
        f"{Color.CYAN}[映射] 月租摊分：platform_shop"
        f"（摊分组 {len(yue_zu_dict)} 个）{Color.RESET}"
    )

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
