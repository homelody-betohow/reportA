import pandas as pd
import importlib.util
from pathlib import Path
# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

# from A_报表.Z_method.sku_映射 import sku_mappings  # 测评费不再映射头程/关税/采购价/运费，暂不需要
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT
from A_报表.Z_method.style import Color

# TODO 文件路径！！！
test_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\测评表\(已完成-2)测评表.xlsx"
# alloc_fee_path = fr"{DESKTOP_ROOT}\调拨费用明细.xlsx"  # 运费映射已停用

print("检查/读取的文件：")
print(f"  测评表：{test_file_path}")
print("  测评费不再计入：头程、关税、采购价、运费（以上四项固定为 0）")

test_file_df = pd.read_excel(test_file_path)

# 测评费不再计算头程、关税、采购价、运费，直接置 0（不再从 SKU 明细/调拨费映射）
_ZERO_COST_COLS = ['头程', '关税', '采购价', '运费']
for _col in _ZERO_COST_COLS:
    test_file_df[_col] = 0

test_file_df_4 = test_file_df.copy()

# ----------------------------------------------------------------------------------------------------------------
# 以下：头程 / 关税 / 采购价 / 运费的 SKU 映射与 RMB→EUR 换算已停用（测评费不再计入）
# ----------------------------------------------------------------------------------------------------------------
# product_map_sku_path = BTH_ALL_SKU_DETAIL_PATH
# test_file_df_1 = sku_mappings(
#     main_df=test_file_df,
#     main_sku='SKU',
#     map_sku_path=product_map_sku_path,
#     map_old_sku="SKU",
#     map_new_sku="原始采购价",
#     map_sku_sheet='基础数据维护'
# )
# test_file_df_2 = sku_mappings(
#     main_df=test_file_df_1,
#     main_sku='SKU',
#     map_sku_path=product_map_sku_path,
#     map_old_sku="SKU",
#     map_new_sku="头程（RMB）",
#     map_sku_sheet='基础数据维护'
# )
# test_file_df_3 = sku_mappings(
#     main_df=test_file_df_2,
#     main_sku='SKU',
#     map_sku_path=product_map_sku_path,
#     map_old_sku="SKU",
#     map_new_sku="关税（含税）",
#     map_sku_sheet='基础数据维护'
# )
# test_file_df_3['映射原始采购价'] = pd.to_numeric(test_file_df_3['映射原始采购价'], errors='coerce')
# test_file_df_3['映射头程（RMB）'] = pd.to_numeric(test_file_df_3['映射头程（RMB）'], errors='coerce')
# test_file_df_3['映射关税（含税）'] = pd.to_numeric(test_file_df_3['映射关税（含税）'], errors='coerce')
# test_file_df_3['采购价'] = test_file_df_3['映射原始采购价'] / 7.3
# test_file_df_3['头程'] = test_file_df_3['映射头程（RMB）'] / 7.3
# test_file_df_3['关税'] = test_file_df_3['映射关税（含税）'] / 7.3
# amazon_sites = test_file_df_3['站点'].str.contains('AMAZON', case=False, na=False)
# amazon_df = test_file_df_3[amazon_sites].copy()
# product_map_sku_path = alloc_fee_path
# amazon_df_1 = sku_mappings(
#     main_df=amazon_df,
#     main_sku='SKU',
#     map_sku_path=product_map_sku_path,
#     map_old_sku="SKU",
#     map_new_sku="德国发FBA运费（EUR）",
#     map_sku_sheet='调拨费'
# )
# amazon_df_1 = amazon_df_1.rename(columns={'映射德国发FBA运费（EUR）': '运费'})
# no_amazon_df = test_file_df_3[~amazon_sites].copy()
# no_amazon_df_1 = sku_mappings(
#     main_df=no_amazon_df,
#     main_sku='SKU',
#     map_sku_path=product_map_sku_path,
#     map_old_sku="SKU",
#     map_new_sku="德国发MF/FBC运费（EUR）",
#     map_sku_sheet='调拨费'
# )
# no_amazon_df_1 = no_amazon_df_1.rename(columns={'映射德国发MF/FBC运费（EUR）': '运费'})
# test_file_df_4 = pd.concat([amazon_df_1, no_amazon_df_1], ignore_index=True)

# ----------------------------------------------------------------------------------------------------------------
# 删掉不要的成本
#  测评涉及费用：
# #print("费用项： 提现费, 销售税, 平台费（头程/关税/采购价/运费 已固定为 0，不参与测评费）")
# # 筛选条件：退款类型是 "佣金" 或 "好评返现"，将提现费、销售税、平台费 置 0
# mask_1 = test_file_df_4['退款类型'].isin(['佣金', '好评返现'])
# test_file_df_4.loc[mask_1, ['头程', '关税', '采购价', '运费', '提现费', '销售税', '平台费']] = 0
# #
# # 筛选条件：退款类型是 "空包退订单金额"
# mask_2 = test_file_df_4['退款类型'] == '空包退订单金额'
# test_file_df_4.loc[mask_2, ['头程', '关税', '采购价', '运费']] = 0
# #
# # 测评退订单金额 仅计算 提现费
# # mask_3 = test_file_df_4['退款类型'] == '测评退订单金额'
# # test_file_df_4.loc[mask_3, ['头程', '关税', '采购价', '运费','销售税', '平台费']] = 0
#
# ----------------------------------------------------------------------------------------------------------------
# 自动检查费用列是否为空（头程/关税/采购价/运费 已固定为 0，仅检查其余三项）
_check_cols = ['提现费', '销售税', '平台费']
_col_source = {
    '提现费': (test_file_path, '（测评表原列）', '提现费'),
    '销售税': (test_file_path, '（测评表原列）', '销售税'),
    '平台费': (test_file_path, '（测评表原列）', '平台费'),
}
_has_empty = False
for _col in _check_cols:
    _empty = test_file_df_4[_col].isna() | (
        test_file_df_4[_col].astype(str).str.strip().isin(['', 'nan', 'None'])
    )
    if _empty.any():
        _has_empty = True
        _src_path, _src_sheet, _src_field = _col_source[_col]
        print(f"{Color.RED}检查失败：「{_col}」存在空值，共 {_empty.sum()} 行{Color.RESET}")
        print(f"  数据来源文件：{_src_path}")
        print(f"  sheet：{_src_sheet}，字段：{_src_field}")
        print(test_file_df_4.loc[_empty, ['SKU', '站点', '退款类型', _col]].drop_duplicates().to_string(index=False))
if _has_empty:
    raise SystemExit(1)
print(f"{Color.GREEN}正常，进入下一步：运行 F3_计算_测评费.py{Color.RESET}")

# 修改列名（订单币种 → 原币种）
# 业务说明：
#   - 订单币种：订单交易的币种（欧洲为EUR，美国为USD）
#   - 退款币种：申请退款的币种
#   - 实际退款币种：支付币种（理论上与退款币种一致）
# 这里使用「报表金额」作为结算币种
test_file_df_4 = test_file_df_4.rename(columns={'订单币种': '原币种'})
# 保存目标列
test_file_df_4 = test_file_df_4[
    ['订单号', '数量', '站点', '退款日期', '退款类型', '订单金额', '原币种', '支付方式', 'SKU', 'SKU-站点识别码', '报表币种',
     '报表金额', '头程', '关税', '采购价', '运费', '提现费', '销售税', '平台费']]

# 保存结果到新的 Excel 文件
output_path = test_file_path.replace('已完成-2', '已完成-3')
test_file_df_4.to_excel(output_path, index=False)
print(f"结果已保存到文件：{output_path}")
