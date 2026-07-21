"""
B2 映射站点 / 构建识别码（订单统计第 2 步）

作用概述：
  1. 读取 B1 输出的订单统计表
  2. 清洗仓库 SKU
  3. 从 platform_shop 映射「映射站点」「映射平台」（替代原桌面「站点-匹配表.xlsx」）
  4. LM_BC_FR / LM_RP_FR 按平台 sku 后缀特殊处理
  5. 生成「SKU-站点识别码」「SKU-平台识别码」

输入：{DESKTOP_ROOT}/{folder_name}{shared_date}/订单统计/(已完成-1-1)订单统计-{shared_date}.xlsx
输出：同目录 (已完成-2)订单统计-{shared_date}.xlsx
"""
import importlib.util
import sys
from pathlib import Path

import pandas as pd

# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.style import Color
from common.platform_shop import apply_lm_fr_region_suffix, map_shop_platform_region
from config.A0_set_date import shared_date, folder_name
from config.A0_paths import DESKTOP_ROOT

# Windows 下避免输出中文乱码
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-1-1)订单统计-{shared_date}.xlsx"
main_df = pd.read_excel(main_file_path)

# 去除整张表的前后空格
for col in main_df.columns:
    main_df[col] = main_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

# 定义例外列表：不去掉尾缀的仓库SKU
exceptions = ['XPYN2125D-1', 'EXPYN2125D-1', 'EBS8029-1']
main_df['仓库SKU'] = main_df['仓库SKU'].where(
    main_df['仓库SKU'].isin(exceptions),
    # 去掉尾缀 -1、-2、-3、-4、-5、_NEVER_USED、-AT；去掉开头的 EXM
    main_df['仓库SKU']
    .str.replace(r'(-1|-2|-3|-4|-5|_NEVER_USED|-AT)$', '', regex=True)
    .str.replace(r'^EXM', '', regex=True)
)

# 重命名
main_df = main_df.rename(columns={'仓库SKU': '原-仓库SKU'})
# 匹配 AMZN.GR. 和 _FB 之间的字符
main_df['仓库SKU'] = main_df['原-仓库SKU'].str.extract(r'AMZN\.GR\.(.*?)_FB')
main_df['仓库SKU'] = main_df['仓库SKU'].fillna(main_df['原-仓库SKU'])

# ------------------- 映射平台 / 站点（数据源：数据库 platform_shop）-------------------
main_df = map_shop_platform_region(main_df, shop_col='店铺英文名', site_col='站点')
main_df = apply_lm_fr_region_suffix(main_df)

# 重命名列
main_df = main_df.rename(columns={'仓库SKU': 'SKU'})

# 检查：映射站点、映射平台 是否有空值（在构建识别码之前拦截）
_check_cols = ['映射站点', '映射平台']
_preview_cols = ['店铺英文名', '站点', '店铺英文名-站点', '订单号', 'SKU']
_missing_info = {}
for _col in _check_cols:
    _mask = main_df[_col].isna() | (main_df[_col].astype(str).str.strip() == '')
    _cnt = int(_mask.sum())
    if _cnt > 0:
        _preview = [c for c in _preview_cols if c in main_df.columns] + [_col]
        _missing_info[_col] = {
            'count': _cnt,
            'preview': main_df.loc[_mask, _preview].drop_duplicates().head(20),
        }

if _missing_info:
    print(f"{Color.RED} --- ====== [错误]映射结果有空值，请先在数据库 platform_shop 表补齐店铺信息后再继续 ====== --- {Color.RESET}")
    for _col, _info in _missing_info.items():
        print(f"{Color.YELLOW}[缺失]{Color.RESET} 列：{_col}，空值行数：{_info['count']}")
        print(_info['preview'].to_string(index=False))
    raise SystemExit(1)

print(f"{Color.GREEN}['映射站点', '映射平台']检查通过，可进行下一步{Color.RESET}")

# 构建识别码
main_df['SKU-站点识别码'] = main_df['映射站点'] + main_df['SKU']
main_df['SKU-平台识别码'] = main_df['映射平台'] + main_df['SKU']

# 保留目标列
output_main_df_3 = main_df[
    ['平台', '店铺英文名', '站点', '店铺英文名-站点', '映射站点', '映射平台', 'SKU-站点识别码', 'SKU-平台识别码',
     '订单类型', '参考号', '订单号', '平台sku', '仓库属性', '仓库',  '运输方式','国家','邮编','SKU', '仓库SKU销量', '跟踪单号', '币种',
     '订单总金额', '平台运费', 'fba费用', '头程运费', '头程税费', '派送运费']]

output_path = main_file_path.replace('已完成-1-1', '已完成-2')
output_main_df_3.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
