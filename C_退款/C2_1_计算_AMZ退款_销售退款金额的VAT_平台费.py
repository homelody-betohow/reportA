import importlib.util
import warnings
from pathlib import Path

# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

# ====================================================================================================

import numpy as np
import pandas as pd
from A_报表.Z_method.style import Color
from A_报表.Z_method.platform_shop import map_shop_platform_region, map_site_vat_commission
from A_报表.A0_设置_时间段.A0_paths import (
    SELLERSKU_PROFIT_FILE_NAME,
    SELLERSKU_PROFIT_REPORT_DIR,
)

# 忽略特定的 UserWarning
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


# 定义 sku 提取规则
def extract_values(s):
    if pd.isna(s):  # 检查是否为 NaN
        return None  # 如果是 NaN，返回 None 或其他默认值
    if 'amzn.gr.' in s:
        return s.split(r'amzn.gr.')[-1].split('-')[0].split('_')[0]
    else:
        return s.split('#')[0].split('BCFBAFL')[0]


# 新版利润报表：Sheet='SellerSku'，前2行为元信息，第3行为列头（header=2）
main_file_path = fr"{SELLERSKU_PROFIT_REPORT_DIR}\{SELLERSKU_PROFIT_FILE_NAME}"
main_file_df = pd.read_excel(main_file_path, sheet_name='SellerSku', header=2)
# main_file_df = pd.read_excel(main_file_path, sheet_name='SellerSku利润报表', header=2)

# 去除 整张表 的前后空格
for col in main_file_df.columns:
    main_file_df[col] = main_file_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

# 筛选  “店铺”不包含 ECO、Biancca、yiqianshangmao_DE 的行
main_file_df = main_file_df[~main_file_df['店铺'].str.contains('ECO|Biancca|yiqianshangmao_DE', na=False)]

# 使用 sellerSku 列的数据填充 仓库sku 列的空值
main_file_df['仓库sku'] = main_file_df['仓库sku'].fillna(main_file_df['sellerSku'])

# 应用提取规则，清洗 仓库sku
main_file_df['仓库sku'] = main_file_df['仓库sku'].apply(extract_values)
# 重命名
main_file_df = main_file_df.rename(columns={'仓库sku': 'SKU'})
# 替换操作
replacements = {
    'E02022001\nE16042004': 'E02022001',
    'E45046100\nE45047002': 'E45046100',
    'E54042001\nE54047001': 'E54042001'
}
# 统一换行符
main_file_df['SKU'] = main_file_df['SKU'].str.replace('\r\n', '\n', regex=False)
# 批量替换
for old, new in replacements.items():
    mask = main_file_df['SKU'].str.contains(old, na=False)
    main_file_df.loc[mask, 'SKU'] = new

# 映射站点 / 映射平台（数据源：platform_shop；有「站点」列时优先店铺-站点精确匹配）
_site_col = '站点' if '站点' in main_file_df.columns else None
main_file_df_2 = map_shop_platform_region(main_file_df, shop_col='店铺', site_col=_site_col)

# 构建识别码
main_file_df_2['SKU-站点识别码'] = main_file_df_2['映射站点'] + main_file_df_2['SKU']
main_file_df_2['SKU-平台识别码'] = main_file_df_2['映射平台'] + main_file_df_2['SKU']

main_file_df_2['退款额'] = main_file_df_2['销售退款金额'] + main_file_df_2['退款服务费用']
# 保留2位小数
main_file_df_2['退款额'] = np.round(main_file_df_2['退款额'], 2)
# 去掉 负号
main_file_df_2['退款额'] = main_file_df_2['退款额'].apply(
    lambda x: abs(x) if isinstance(x, (int, float)) else x)

# 筛选 退款量 列和 退款额 列均不等于 0 的行
main_file_df_2 = main_file_df_2[(main_file_df_2['退款量'] != 0) & (main_file_df_2['退款额'] != 0)]
# 保留指定列
# 新版 SellerSku 报表已无「历史ASIN」列，仅保留现有 ASIN
main_file_df_3 = main_file_df_2[
    ['sellerSku', 'ASIN', '产品信息', 'SKU', '店铺', '映射站点', '映射平台', 'SKU-站点识别码',
     'SKU-平台识别码', '退款服务费用', '销售退款金额', '退款额', '退款量']]
# 不要 sellerSku  为空 的数据
main_file_df_3 = main_file_df_3.dropna(subset=['sellerSku'])

# TODO 计算Amazon 的 销售退款金额VAT(不计算'退款服务费用')
# 映射 VAT税（DB 优先，Excel 兜底），按 映射站点 匹配 platform_shop.market_region
main_file_df_4 = map_site_vat_commission(main_df=main_file_df_3, site_col='映射站点')
# 计算VAT
main_file_df_4['销售退款金额VAT-amazon'] = main_file_df_4['销售退款金额'] * main_file_df_4['映射VAT税']
#
# ========================================================================================
# 筛选：店铺 不包含 'BinFen'，对应‘销售退款金额VAT-amazon’列的数据替换成0
condition = ~main_file_df_4['店铺'].str.contains('BinFen')
main_file_df_4.loc[condition, '销售退款金额VAT-amazon'] = 0
# 去掉 负号
main_file_df_4['销售退款金额VAT-amazon'] = main_file_df_4['销售退款金额VAT-amazon'].apply(
    lambda x: abs(x) if isinstance(x, (int, float)) else x)
# 保留2位小数
main_file_df_4['销售退款金额VAT-amazon'] = np.round(main_file_df_4['销售退款金额VAT-amazon'], 2)

# TODO 计算Amazon 的 销售退款金额的佣金
# AMAZON-EU 的 销售退款金额的佣金 = 销售退款金额 * 0.125
# AMAZON-US 的 销售退款金额的佣金 = 销售退款金额 * 0.12
main_file_df_4['销售退款金额的佣金'] = np.where(
    main_file_df_4['映射平台'] == 'AMAZON-EU',
    main_file_df_4['销售退款金额'] * 0.125,
    main_file_df_4['销售退款金额'] * 0.12
)
# 去掉 负号
main_file_df_4['销售退款金额的佣金'] = main_file_df_4['销售退款金额的佣金'].apply(
    lambda x: abs(x) if isinstance(x, (int, float)) else x)
# 保留2位小数
main_file_df_4['销售退款金额的佣金'] = np.round(main_file_df_4['销售退款金额的佣金'], 2)

# 保存修改
output_path = fr"{SELLERSKU_PROFIT_REPORT_DIR}\(退款-1){SELLERSKU_PROFIT_FILE_NAME}"
main_file_df_4.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')

# 自动检查 映射站点、映射平台 是否为空
_check_cols = ['映射站点', '映射平台']
for _col in _check_cols:
    _empty = main_file_df_4[_col].isna() | (
        main_file_df_4[_col].astype(str).str.strip().isin(['', 'nan', 'None'])
    )
    if _empty.any():
        print(f"{Color.RED}检查失败：「{_col}」存在空值，共 {_empty.sum()} 行{Color.RESET}")
        print(main_file_df_4.loc[_empty, ['店铺', 'SKU', _col]].drop_duplicates().to_string(index=False))
        raise SystemExit(1)
print(f"{Color.GREEN}一切正常，进入下一步：运行 C2_2_合并_AMZ退款_销售退款金额的VAT_平台费.py{Color.RESET}")
