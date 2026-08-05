"""
H1 映射计算 —— AMZ SellerSku 利润报表预处理（H 系列第 1 步）

作用概述：
  1. 读取原始 SellerSku 利润报表
  2. 过滤无关店铺与空 sellerSku、拆分组合 SKU（sellerSku 含 +）、汇总广告费/赔偿/其他分摊费用
  3. 清洗并标准化 SKU
  4. 从 platform_shop 将店铺映射为站点、平台
  5. 生成「SKU-站点识别码」「SKU-平台识别码」供后续 H2~H4 合并分摊

输入：{SELLERSKU_PROFIT_REPORT_DIR}/{SELLERSKU_PROFIT_FILE_NAME}
输出：{SELLERSKU_PROFIT_REPORT_DIR}/(已完成-1){SELLERSKU_PROFIT_FILE_NAME}
下游：H2_合并_儿子-站点识别码.py
"""
import warnings
import pandas as pd
import importlib.util
import sys
from pathlib import Path

# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.style import Color
from common.split_rows_data_SKU import (
    extract_internal_sku,
    split_one_rows_data,
    strip_hash_suffix,
)
from common.platform_shop import map_shop_platform_region
from config.A0_paths import SELLERSKU_PROFIT_FILE_NAME, SELLERSKU_PROFIT_REPORT_DIR

# 组合 SKU 拆分时需均摊的金额列（与下方「计算结果」汇总所用列一致）
_COMBO_FEE_COLUMNS = [
    'SD广告费', 'SP广告费', 'SB广告费', 'SBV广告费',
    '其他交易费汇总', '移除费用', '合作承运费', '合仓费', '超量费',
    '其他FBA库存和入境服务费', 'FBA退货处理费', 'coupon优惠券',
    'FBA月订阅费(平台店租)', '其他服务费', '平台其他支出汇总',
    'FBA库存赔偿汇总',
]

# openpyxl 读取 xlsx 时可能产生无害的 UserWarning，此处统一忽略
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Windows 下避免输出中文乱码（Cursor/终端捕获常见编码问题）
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# ---------------------------------------------------------------------------
# 一、读取原始利润报表
# ---------------------------------------------------------------------------
# 新版利润报表：Sheet='SellerSku'，前 2 行为元信息，第 3 行为列头（header=2）
main_file_path = fr"{SELLERSKU_PROFIT_REPORT_DIR}\{SELLERSKU_PROFIT_FILE_NAME}"
main_file_df = pd.read_excel(main_file_path, sheet_name='SellerSku', header=2)

# 去除整张表各列字符串值的前后空格，避免映射时因空格匹配失败
for col in main_file_df.columns:
    main_file_df[col] = main_file_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

# ---------------------------------------------------------------------------
# 二、筛选店铺 & 计算汇总费用列
# ---------------------------------------------------------------------------
# 排除 ECO、Biancca、yiqianshangmao_DE 等非本报表统计范围的店铺
main_file_df_1 = main_file_df[~main_file_df['店铺'].str.contains('ECO|Biancca|yiqianshangmao_DE', na=False)]

# sellerSku 为空则跳过，不写入 (已完成-1)
_empty_sku_mask = main_file_df_1['sellerSku'].isna() | (main_file_df_1['sellerSku'].astype(str).str.strip() == '')
_empty_sku_cnt = int(_empty_sku_mask.sum())
if _empty_sku_cnt > 0:
    print(f"{Color.YELLOW}[跳过]{Color.RESET} sellerSku 为空 {_empty_sku_cnt} 行，不写入 (已完成-1)")
    main_file_df_1 = main_file_df_1.loc[~_empty_sku_mask].copy()

# ---------------------------------------------------------------------------
# 2.5 拆分组合 SKU（sellerSku 含 '+' 或 ','）
# ---------------------------------------------------------------------------
# 与 D 广告 / E 秒杀 等脚本一致：按子 SKU 数量均摊费用，仓库sku 同步为子 SKU
_combo_mask = main_file_df_1['sellerSku'].astype(str).str.contains(r'[+,]', na=False)
_combo_cnt = int(_combo_mask.sum())
if _combo_cnt > 0:
    _fee_cols = [c for c in _COMBO_FEE_COLUMNS if c in main_file_df_1.columns]
    _rows_before = len(main_file_df_1)
    main_file_df_1 = split_one_rows_data(
        input_df=main_file_df_1,
        data_column='sellerSku',
        value_column=_fee_cols,
        sync_columns=['仓库sku'],
    )
    print(
        f"{Color.YELLOW}[组合SKU]{Color.RESET} 拆分 {_combo_cnt} 行（sellerSku 含 +/，），"
        f"行数 {_rows_before} → {len(main_file_df_1)}，金额按子 SKU 均摊"
    )

# 赔偿：直接沿用报表中的「FBA库存赔偿汇总」列
main_file_df_1 = main_file_df_1.rename(columns={'FBA库存赔偿汇总': '计算结果-赔偿'})

# 广告费 = 四种亚马逊广告类型之和
main_file_df_1['计算结果-广告费'] = main_file_df_1['SD广告费'] + main_file_df_1['SP广告费'] + main_file_df_1[
    'SB广告费'] + main_file_df_1['SBV广告费']

# 其他分摊费用：将多项平台/FBA 杂费合并为一列，便于后续按 SKU 分摊
main_file_df_1['计算结果-其他分摊费用'] = main_file_df_1['其他交易费汇总'] + main_file_df_1['移除费用'] + \
                                          main_file_df_1['合作承运费'] + main_file_df_1['合仓费'] + main_file_df_1[
                                              '超量费'] + main_file_df_1['其他FBA库存和入境服务费'] + main_file_df_1[
                                              'FBA退货处理费'] + main_file_df_1['coupon优惠券'] + main_file_df_1[
                                              'FBA月订阅费(平台店租)'] + main_file_df_1['其他服务费'] + main_file_df_1[
                                              '平台其他支出汇总']

# ---------------------------------------------------------------------------
# 三、清洗 SKU（仓库 sku → 标准 SKU）
# ---------------------------------------------------------------------------
# 仓库 sku 为空时，用 sellerSku 回填；去掉 # 尾缀（如 #FBDE、#FBFBA）
main_file_df_1['仓库sku'] = main_file_df_1['仓库sku'].fillna(main_file_df_1['sellerSku'])
main_file_df_1['仓库sku'] = main_file_df_1['仓库sku'].apply(strip_hash_suffix)


def extract_values(s):
    """内部编码（AMZN.GR.）+ 去掉 BCFBAFL 后缀（# 尾缀已在上方 strip_hash_suffix 处理）。"""
    if pd.isna(s):
        return None
    return str(extract_internal_sku(s)).split("BCFBAFL")[0]


main_file_df_1['仓库sku'] = main_file_df_1['仓库sku'].apply(extract_values)
main_file_df_1 = main_file_df_1.rename(columns={'仓库sku': 'SKU'})
main_file_df_1['SKU'] = main_file_df_1['SKU'].fillna('无')

# 历史数据中个别 SKU 单元格含换行拼接的多编码，统一替换为第一个有效编码
replacements = {
    'E02022001\nE16042004': 'E02022001',
    'E45046100\nE45047002': 'E45046100',
    'E54042001\nE54047001': 'E54042001'
}
main_file_df_1['SKU'] = main_file_df_1['SKU'].str.replace('\r\n', '\n', regex=False)
for old, new in replacements.items():
    mask = main_file_df_1['SKU'].str.contains(old, na=False)
    main_file_df_1.loc[mask, 'SKU'] = new

# ---------------------------------------------------------------------------
# 四、站点 / 平台映射（数据源：数据库 platform_shop）
# ---------------------------------------------------------------------------
main_file_df_1 = map_shop_platform_region(main_file_df_1, shop_col='店铺', site_col='站点')

# SKU-站点识别码 = 映射站点 + SKU；SKU-平台识别码 = 映射平台 + SKU
_site_code = main_file_df_1["映射站点"].fillna("").astype(str) + main_file_df_1["SKU"].fillna("").astype(str)
_platform_code = main_file_df_1["映射平台"].fillna("").astype(str) + main_file_df_1["SKU"].fillna("").astype(str)
_pos = main_file_df_1.columns.get_loc("映射站点") + 1
main_file_df_1.insert(_pos, "SKU-站点识别码", _site_code)
main_file_df_1.insert(_pos + 1, "SKU-平台识别码", _platform_code)

# 只保留后续流程需要的列（新版报表已无「历史ASIN」）
main_file_df_3 = main_file_df_1[
    ['sellerSku', 'ASIN', '产品信息', 'SKU', '店铺', '映射站点', '映射平台', 'SKU-站点识别码',
     'SKU-平台识别码', 'SD广告费', 'SP广告费', 'SB广告费', 'SBV广告费', '其他交易费汇总', '移除费用', '合作承运费',
     '合仓费', '超量费', '其他FBA库存和入境服务费', 'FBA退货处理费', 'coupon优惠券', 'FBA月订阅费(平台店租)',
     '其他服务费', '平台其他支出汇总', '计算结果-广告费', '计算结果-赔偿', '计算结果-其他分摊费用']]

# ---------------------------------------------------------------------------
# 五、映射完整性校验（有空值则中断，提示在 platform_shop 补齐）
# ---------------------------------------------------------------------------
_check_cols = ["映射站点", "映射平台"]
_missing_info = {}
for _col in _check_cols:
    _mask = main_file_df_3[_col].isna() | (main_file_df_3[_col].astype(str).str.strip() == "")
    _cnt = int(_mask.sum())
    if _cnt > 0:
        _missing_info[_col] = {
            "count": _cnt,
            "preview": main_file_df_3.loc[_mask, ["店铺", "SKU", "sellerSku", _col]].head(20),
        }

if _missing_info:
    print(f"{Color.RED} --- ====== [错误]映射结果有空值，请先在数据库 platform_shop 表补齐店铺信息后再继续 ====== --- {Color.RESET}")
    for _col, _info in _missing_info.items():
        print(f"{Color.YELLOW}[缺失]{Color.RESET} 列：{_col}，空值行数：{_info['count']}")
        print(_info["preview"].to_string(index=False))
    raise SystemExit(1)
else:
    print(f"{Color.GREEN} --- ====== [一切正常]，进入下一步（保存文件） ====== --- {Color.RESET}")

# ---------------------------------------------------------------------------
# 六、保存中间结果，供 H2 读取
# ---------------------------------------------------------------------------
output_path = fr"{SELLERSKU_PROFIT_REPORT_DIR}\(已完成-1){SELLERSKU_PROFIT_FILE_NAME}"
try:
    main_file_df_3.to_excel(output_path, index=False)
except PermissionError:
    print(f"{Color.RED}保存失败：目标文件被占用/无权限。请先关闭已打开的输出文件后重试：{output_path}{Color.RESET}")
    raise
print(f'处理完成，文件另存为：{output_path}')
