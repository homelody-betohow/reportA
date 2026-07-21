"""
C1_3 映射站点 / 构建识别码（退款 RMA 第 1-3 步）

作用概述：
  1. 读取 (已完成-1)RMA 表，过滤无关店铺，清洗 RMA产品
  2. 从 platform_shop 映射「映射站点」「映射平台」（替代原桌面「站点-匹配表.xlsx」）
  3. LM_BC_FR / LM_RP_FR 按平台 sku 后缀特殊处理
  4. 生成「SKU-站点识别码」「SKU-平台识别码」
  5. 按「仓库名称」是否含「分销」生成「分销」列（是/否）

输入：{DESKTOP_ROOT}/{folder_name}{shared_date}/RMA/(已完成-1)RMA-{shared_date}.xlsx
输出：同目录 (已完成-1-1)RMA-{shared_date}.xlsx
"""
import importlib.util
import sys
import warnings
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

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# TODO 文件路径！！！
RMA_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\RMA\(已完成-1)RMA-{shared_date}.xlsx"
RMA_file_df = pd.read_excel(RMA_file_path)

# 去除整张表的前后空格
for col in RMA_file_df.columns:
    RMA_file_df[col] = RMA_file_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

# 筛选「店铺英文名」不包含 ECO、Biancca、yiqianshangmao_DE 的行
RMA_file_df_1 = RMA_file_df[
    ~RMA_file_df['店铺英文名'].str.contains('ECO|Biancca|yiqianshangmao_DE', na=False)
].copy()

# 例外列表：不去掉尾缀的仓库 SKU
exceptions = ['XPYN2125D-1', 'EXPYN2125D-1', 'EBS8029-1']
RMA_file_df_1['RMA产品'] = RMA_file_df_1['RMA产品'].where(
    RMA_file_df_1['RMA产品'].isin(exceptions),
    RMA_file_df_1['RMA产品']
    .str.replace(r'(-1|-2|-3|-4|-5|_NEVER_USED|-AT)$', '', regex=True)
    .str.replace(r'^EXM', '', regex=True)
)

# ------------------- 映射平台 / 站点（数据源：数据库 platform_shop）-------------------
# 退款表用「订单目的国家」拼「店铺英文名-站点」，对应 platform_shop.platform_site
RMA_file_df_3 = map_shop_platform_region(
    RMA_file_df_1,
    shop_col='店铺英文名',
    site_col='订单目的国家',
)
RMA_file_df_3 = apply_lm_fr_region_suffix(RMA_file_df_3)

# 重命名
RMA_file_df_3 = RMA_file_df_3.rename(columns={'RMA产品': 'SKU'})

# 检查：映射站点、映射平台 是否有空值（在构建识别码之前拦截）
_check_cols = ['映射站点', '映射平台']
_preview_cols = ['店铺英文名', '订单目的国家', '店铺英文名-站点', '平台sku', 'SKU', '退款原订单号']
_missing_info = {}
for _col in _check_cols:
    _mask = RMA_file_df_3[_col].isna() | (RMA_file_df_3[_col].astype(str).str.strip() == '')
    _cnt = int(_mask.sum())
    if _cnt > 0:
        _preview = [c for c in _preview_cols if c in RMA_file_df_3.columns] + [_col]
        _missing_info[_col] = {
            'count': _cnt,
            'preview': RMA_file_df_3.loc[_mask, _preview].drop_duplicates().head(20),
        }

if _missing_info:
    print(f"{Color.RED} --- ====== [错误]映射结果有空值，请先在数据库 platform_shop 表补齐店铺信息后再继续 ====== --- {Color.RESET}")
    for _col, _info in _missing_info.items():
        print(f"{Color.YELLOW}[缺失]{Color.RESET} 列：{_col}，空值行数：{_info['count']}")
        print(_info['preview'].to_string(index=False))
    raise SystemExit(1)

print(f"{Color.GREEN}['映射站点', '映射平台']检查通过，可进行下一步{Color.RESET}")

# 构建识别码
RMA_file_df_3['SKU-站点识别码'] = RMA_file_df_3['映射站点'] + RMA_file_df_3['SKU']
RMA_file_df_3['SKU-平台识别码'] = RMA_file_df_3['映射平台'] + RMA_file_df_3['SKU']

# 分销标识：已完成-1 中 仓库名称 含「分销」则标记为「是」，否则为「否」
RMA_file_df_3['分销'] = '否'
RMA_file_df_3.loc[
    RMA_file_df_3['仓库名称'].astype(str).str.contains('分销', na=False), '分销'
] = '是'

# 保留指定列
RMA_file_df_3 = RMA_file_df_3[
    ['平台', '店铺英文名', '订单目的国家', '店铺英文名-站点', '映射站点', '映射平台', 'SKU-站点识别码',
     'SKU-平台识别码', '退款原订单号', '平台sku', 'SKU', 'RMA产品数量', '退款金额', '退款状态', '分销']]

output_path = RMA_file_path.replace('已完成-1', '已完成-1-1')
RMA_file_df_3.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
