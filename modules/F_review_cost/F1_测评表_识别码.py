"""
测评表处理脚本 - F1_映射_识别码
功能：读取测评表，进行日期筛选、SKU拆分、站点映射、币种换算、VAT和平台费计算
输出：(已完成-1)测评表.xlsx、(已完成-2)测评表.xlsx

表头变更说明（2026.5 起）：
- 旧表头（≤2026.4）：币种
- 新表头（≥2026.5）：订单币种、退款币种、实际退款币种

币种业务说明：
- 订单币种：订单交易的币种（欧洲为EUR，美国为USD）
- 退款币种：申请退款的币种
- 实际退款币种：支付币种（理论上与退款币种一致）

币种换算逻辑：
- 优先使用「订单币种」换算订单金额为欧元
- 若订单币种为空，使用「退款币种」或「实际退款币种」兜底
- 最终输出：保留原始币种列 + 报表币种（EUR）
"""
import importlib.util
import re
from pathlib import Path

# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

import numpy as np
import pandas as pd
from config.A0_set_date import *
from common.sku_mapping import sku_mappings
from common.split_rows_data_SKU import split_one_rows_data
from common.platform_shop import map_site_vat_commission
from config.A0_paths import DESKTOP_ROOT
from common.style import Color

def convert_to_eur(row):
    """
    将订单金额按订单币种换算为欧元
    订单币种：订单交易时的币种（欧洲为EUR，美国为USD）
    """
    amount = row['订单金额']
    currency = row['订单币种']
    
    # 处理空值或nan
    if pd.isna(currency):
        # 订单币种为空时，尝试使用退款币种或实际退款币种
        if '退款币种' in row.index and pd.notna(row['退款币种']):
            currency = row['退款币种']
        elif '实际退款币种' in row.index and pd.notna(row['实际退款币种']):
            currency = row['实际退款币种']
        else:
            print(f'报错：订单币种为空且无退款币种，订单号：{row.get("订单号", "未知")}')
            return amount  # 返回原值

    if currency == 'RMB':
        return amount / RMB_di_EUR
    elif currency == 'USD':
        return amount * USD_to_EUR
    elif currency == 'EUR':
        return amount
    elif currency == 'CAD':
        return amount * CAD_to_EUR
    elif currency == 'CZK':
        return amount * kc_to_EUR
    elif currency == 'PLN':
        return amount * zl_to_EUR
    elif currency == 'HUF':
        return amount * Ft_to_EUR
    elif currency == 'RON':
        return amount * Lei_to_EUR
    elif currency == 'SEK':
        return amount * kr_to_EUR
    else:
        print(f'报错：未知币种，币种：{currency}，请检查！！！！！！程序终止！！！！！！！！！！！！！！！！')
        exit()

TEST_TABLE_SOURCE_DIR = Path(r"\\Betohow\数据报表\报表自动化下载\广告下载\每天\测评表")


def _pick_test_table_file_for_end_date(end_date: str) -> Path:
    """读取测评表目录下文件名日期等于 test_end_date 的 测评表{日期}.xlsx"""
    if not TEST_TABLE_SOURCE_DIR.is_dir():
        raise FileNotFoundError(f"测评表目录不存在：{TEST_TABLE_SOURCE_DIR}")
    target_date = pd.to_datetime(end_date).strftime("%Y-%m-%d")
    expected = TEST_TABLE_SOURCE_DIR / f"测评表{target_date}.xlsx"
    if not expected.is_file():
        raise FileNotFoundError(
            f"未找到日期为 {target_date} 的测评表文件：{expected}"
        )
    return expected


test_file_path = str(_pick_test_table_file_for_end_date(test_end_date))
output_dir = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\测评表"
Path(output_dir).mkdir(parents=True, exist_ok=True)
print(f"读取测评表：{Color.YELLOW} {test_file_path} {Color.RESET}")

test_file_df = pd.read_excel(test_file_path, sheet_name=test_file_sheet_name)  # 对应的月份！
test_file_df.columns = test_file_df.columns.str.strip()

# 仅保留 “退款类型” 为 
# “测评退订单金额” 或 “空包退订单金额” 或 “佣金” 或 “好评返现” 的行
# 注意：不要过滤掉 “测评退订单金额70%”
test_file_df = test_file_df[test_file_df['退款类型'].isin(['测评退订单金额','测评退订单金额70%', '空包退订单金额', '好评返现', '佣金'])]
#
# ===== 2026.5 新表头适配：币种列检查 =====
# 新表头：订单币种、退款币种、实际退款币种
# 业务说明：
#   - 订单币种：订单交易的币种（欧洲为EUR，美国为USD）- 主要使用
#   - 退款币种：申请退款的币种
#   - 实际退款币种：支付币种（理论上与退款币种一致）
# 换算逻辑：优先使用订单币种，为空时使用退款币种兜底
if '订单币种' not in test_file_df.columns:
    raise KeyError('测评表缺少必要的币种列：请确认表头是否包含「订单币种」列！')
# # 填充‘退款日期’的空值，因为部分‘退款日期’是合并的 （合并的数据，python会读取在对应的“第一行”，其余行为空）
# # .ffill()：用前面的非空值来填充当前的空值（只会填充合并拆开后的空 “退款日期”）
# test_file_df['退款日期'] = test_file_df['退款日期'].ffill()
#
# 对应的时间段！！！
start_date = pd.to_datetime(test_start_date)
end_date = pd.to_datetime(test_end_date)
# Excel 读入的「退款日期」常为字符串，须先转为日期再与 start/end 比较
test_file_df['退款日期'] = pd.to_datetime(test_file_df['退款日期'], errors='coerce')

# 筛选条件
print(f"---=== {Color.GREEN}开始筛选退款日期 {start_date} 到 {end_date} 之间的数据 {Color.RESET} ===---\n")
test_file_df = test_file_df[(test_file_df['退款日期'] >= start_date) & (test_file_df['退款日期'] <= end_date)]
# 去除 整张表 的前后空格
for col in test_file_df.columns:
    test_file_df[col] = test_file_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
# 保存修改后的（保留所有原始列：订单币种、退款币种、实际退款币种）
output_path = fr"{output_dir}\(已完成-1)测评表.xlsx"
test_file_df.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
print(f"----------------------\n")

# 拆分有“+”的sku
test_file_df_1 = split_one_rows_data(
    input_df=test_file_df,
    data_column='SKU',
    value_column='订单金额'
)
# SKU 去掉尾缀 -1、-2、-3、-4、-5、-6、-7、-8
test_file_df_1['SKU'] = test_file_df_1['SKU'].apply(
    lambda x: re.sub(r'(-1|-2|-3|-4|-5|-6|-7|-8)$', '', x) if re.search(r'(-1|-2|-3|-4|-5|-6|-7|-8)$', x) else x)
# 去除 '平台' 列的前后空格
test_file_df_1['平台'] = test_file_df_1['平台'].str.strip()
# 将 '站点' 列中的 '_' 替换为 '-'
test_file_df_1['平台'] = test_file_df_1['平台'].str.replace('_', '-')
# 替换操作
test_file_df_1['平台'] = test_file_df_1['平台'].replace('REAL', 'REAL-DE-FB')
test_file_df_1['平台'] = test_file_df_1['平台'].replace('MANO-EU', 'MANO-FR-OHPA')
test_file_df_1['平台'] = test_file_df_1['平台'].replace('LM-Toto-FR', 'LM-TOTO')
test_file_df_1['平台'] = test_file_df_1['平台'].replace('LM-FR', 'LM-FR-BTH')
test_file_df_1['平台'] = test_file_df_1['平台'].replace('MANO-BTH-FR', 'MANO-FR-BTH')
test_file_df_1['平台'] = test_file_df_1['平台'].replace('MANO-OHPA-FR', 'MANO-FR-OHPA')
test_file_df_1['平台'] = test_file_df_1['平台'].replace('Temu-Bathvogue', 'TEMU-BV')
test_file_df_1['平台'] = test_file_df_1['平台'].replace('Temu-BV', 'TEMU-BV')
test_file_df_1['平台'] = test_file_df_1['平台'].replace('Castorama', 'castorama')
test_file_df_1['平台'] = test_file_df_1['平台'].replace('OTTO', 'OTTO-BTH')
test_file_df_1['平台'] = test_file_df_1['平台'].replace('MANO-COMFR', 'MANO-FR-COM')
test_file_df_1['平台'] = test_file_df_1['平台'].replace('MANO-COMDE', 'MANO-DE-COM')

# 测评表的“平台” 就是 “站点”，更改列名
test_file_df_1 = test_file_df_1.rename(columns={'平台': '站点'})
# 在 映射平台 后插入新列 SKU-站点识别码
new_column_name = "SKU-站点识别码"  # 新列名
new_column_data = test_file_df_1["站点"] + test_file_df_1["SKU"]  # 新列数据
target_column = "站点"  # 目标列名（在其后插入）
insert_position = test_file_df_1.columns.get_loc(target_column) + 1  # 计算插入位置
test_file_df_1.insert(insert_position, new_column_name, new_column_data)  # 插入新列

# 币种 转成 欧元
test_file_df_1['报表币种'] = 'EUR'
# test_file_df_1['订单金额'] = np.round(test_file_df_1.apply(convert_to_eur, axis=1), 2)
test_file_df_1['报表金额'] = np.round(test_file_df_1.apply(convert_to_eur, axis=1), 2)
# 提现费
print(f"提现费 = 订单金额 * 0.05 + 0.34\n")
test_file_df_1['提现费'] = np.round(test_file_df_1['报表金额'] * 0.05 + 0.34, 2)

# 映射 castorama 的 佣金比例，映射不到的问：晓佳
product_map_sku_path = fr"{DESKTOP_ROOT}\castorama - SKU类目佣金比例.xlsx"
test_file_df_2 = sku_mappings(
    main_df=test_file_df_1,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="SKU",
    map_new_sku="佣金比",
    map_sku_sheet='Sheet1'
)
# 映射 平台费（佣金）、VAT税（来源：platform_shop.market_region）
test_file_df_3 = map_site_vat_commission(main_df=test_file_df_2, site_col='站点')
# 用“映射佣金比”填补“映射平台费（佣金）”的空值（castorama 等按 SKU 类目佣金）
test_file_df_3['映射平台费（佣金）'] = test_file_df_3['映射平台费（佣金）'].fillna(test_file_df_3['映射佣金比'])
test_file_df_4 = test_file_df_3
# 计算平台费和VAT（映射列须为数值；未映射到 VAT 的站点会为空，需在 DB 或 Excel 映射表补全）
for _rate_col in ('映射平台费（佣金）', '映射佣金比', '映射VAT税'):
    test_file_df_4[_rate_col] = pd.to_numeric(test_file_df_4[_rate_col], errors='coerce')
test_file_df_4['订单金额'] = pd.to_numeric(test_file_df_4['订单金额'], errors='coerce')
_unmapped_vat = test_file_df_4['映射VAT税'].isna()
if _unmapped_vat.any():
    _sites = test_file_df_4.loc[_unmapped_vat, '站点'].drop_duplicates().tolist()
    raise ValueError(
        f"以下站点未映射到 VAT税，请在 platform_shop 或「VAT、平台费-映射.xlsx」中补充后重跑：{_sites}"
    )

print(f"平台费 = 订单金额 * 映射平台费（佣金）\n")
test_file_df_4['平台费'] = np.round(test_file_df_4['订单金额'] * test_file_df_4['映射平台费（佣金）'], 2)
print(f"销售税 = 订单金额 * 映射VAT税\n")
test_file_df_4['销售税'] = np.round(test_file_df_4['订单金额'] * test_file_df_4['映射VAT税'], 2)

# 检查'订单金额'是否存在空值
if test_file_df_4['订单金额'].isnull().any():
    raise ValueError("错误：'订单金额'存在空值，请联系相应的”运营“，进行填写！")
# 保存修改后的（保留所有列，包括原始的订单币种、退款币种、实际退款币种）
output_path = fr"{output_dir}\(已完成-2)测评表.xlsx"
test_file_df_4.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')