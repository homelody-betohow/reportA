import warnings
import pandas as pd
import importlib.util
import sys
from pathlib import Path
# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.Z_method.style import Color
from A_报表.Z_method.sku_映射 import sku_mappings
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT, SELLERSKU_PROFIT_FILE_NAME, SELLERSKU_PROFIT_REPORT_DIR

# 忽略特定的 UserWarning
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Windows 下避免输出中文乱码（Cursor/终端捕获常见编码问题）
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# TODO 文件路径！！！
# 新版利润报表：Sheet='SellerSku'，前2行为元信息，第3行为列头（header=2）
main_file_path = fr"{SELLERSKU_PROFIT_REPORT_DIR}\{SELLERSKU_PROFIT_FILE_NAME}"
main_file_df = pd.read_excel(main_file_path, sheet_name='SellerSku', header=2)
# 去除 整张表 的前后空格
for col in main_file_df.columns:
    main_file_df[col] = main_file_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

# 筛选  “店铺”不包含 ECO、Biancca、yiqianshangmao_DE 的行
main_file_df_1 = main_file_df[~main_file_df['店铺'].str.contains('ECO|Biancca|yiqianshangmao_DE', na=False)]

main_file_df_1 = main_file_df_1.rename(columns={'FBA库存赔偿汇总': '计算结果-赔偿'})  # 赔偿=FBA库存赔偿汇总

# 广告费 = SD广告费 + SP广告费 + SB广告费 + SBV广告费
main_file_df_1['计算结果-广告费'] = main_file_df_1['SD广告费'] + main_file_df_1['SP广告费'] + main_file_df_1[
    'SB广告费'] + main_file_df_1['SBV广告费']

# 其他摊分=其他交易费汇总+移除费用+合作承运费+合仓费+超量费+其他FBA库存和入境服务费+FBA退货处理费+coupon优惠券+FBA月订阅费(平台店租)+其他服务费+平台其他支出汇总
main_file_df_1['计算结果-其他分摊费用'] = main_file_df_1['其他交易费汇总'] + main_file_df_1['移除费用'] + \
                                          main_file_df_1['合作承运费'] + main_file_df_1['合仓费'] + main_file_df_1[
                                              '超量费'] + main_file_df_1['其他FBA库存和入境服务费'] + main_file_df_1[
                                              'FBA退货处理费'] + main_file_df_1['coupon优惠券'] + main_file_df_1[
                                              'FBA月订阅费(平台店租)'] + main_file_df_1['其他服务费'] + main_file_df_1[
                                              '平台其他支出汇总']

# 使用 sellerSku 列的数据填充 仓库sku 列的空值
main_file_df_1['仓库sku'] = main_file_df_1['仓库sku'].fillna(main_file_df_1['sellerSku'])


# 定义 sku 提取规则
def extract_values(s):
    if pd.isna(s):  # 检查是否为 NaN
        return None  # 如果是 NaN，返回 None 或其他默认值
    if 'amzn.gr.' in s:
        return s.split(r'amzn.gr.')[-1].split('-')[0].split('_')[0]
    else:
        return s.split('#')[0].split('BCFBAFL')[0]


# 应用提取规则，清洗 仓库sku
main_file_df_1['仓库sku'] = main_file_df_1['仓库sku'].apply(extract_values)
main_file_df_1 = main_file_df_1.rename(columns={'仓库sku': 'SKU'})

# 将列 "SKU" 中的 空值 替换为 "无"
main_file_df_1['SKU'] = main_file_df_1['SKU'].fillna('无')

# 替换操作
replacements = {
    'E02022001\nE16042004': 'E02022001',
    'E45046100\nE45047002': 'E45046100',
    'E54042001\nE54047001': 'E54042001'
}
# 统一换行符
main_file_df_1['SKU'] = main_file_df_1['SKU'].str.replace('\r\n', '\n', regex=False)
# 批量替换
for old, new in replacements.items():
    mask = main_file_df_1['SKU'].str.contains(old, na=False)
    main_file_df_1.loc[mask, 'SKU'] = new

product_map_sku_path = fr"{DESKTOP_ROOT}\站点-匹配表.xlsx"  # 改成对应的映射表
# 映射 站点
main_file_df_2 = sku_mappings(
    main_df=main_file_df_1,
    main_sku='店铺',
    map_sku_path=product_map_sku_path,
    map_old_sku="平台账号",
    map_new_sku="站点",
    map_sku_sheet='站点匹配'
)

# 在 映射站点 后插入新列 儿子-站点识别码
new_column_name = "儿子-站点识别码"  # 新列名
new_column_data = main_file_df_2["映射站点"].fillna("").astype(str) + main_file_df_2["SKU"].fillna("").astype(str)  # 新列数据
target_column = "映射站点"  # 目标列名（在其后插入）
insert_position = main_file_df_2.columns.get_loc(target_column) + 1  # 计算插入位置
main_file_df_2.insert(insert_position, new_column_name, new_column_data)  # 插入新列

# 映射 平台
main_file_df_3 = sku_mappings(
    main_df=main_file_df_2,
    main_sku='映射站点',
    map_sku_path=product_map_sku_path,
    map_old_sku="站点",
    map_new_sku="平台",
    map_sku_sheet='站点匹配'
)
# 在 儿子-站点识别码 后插入 儿子-平台识别码
new_column_name = "儿子-平台识别码"  # 新列名
new_column_data = main_file_df_3["映射平台"].fillna("").astype(str) + main_file_df_3["SKU"].fillna("").astype(str)  # 新列数据
target_column = "儿子-站点识别码"  # 目标列名（在其后插入）
insert_position = main_file_df_3.columns.get_loc(target_column) + 1  # 计算插入位置
main_file_df_3.insert(insert_position, new_column_name, new_column_data)  # 插入新列

main_file_df_3 = main_file_df_3[
    ['sellerSku', 'ASIN', '历史ASIN', '产品信息', 'SKU', '店铺', '映射站点', '映射平台', '儿子-站点识别码',
     '儿子-平台识别码', 'SD广告费', 'SP广告费', 'SB广告费', 'SBV广告费', '其他交易费汇总', '移除费用', '合作承运费',
     '合仓费', '超量费', '其他FBA库存和入境服务费', 'FBA退货处理费', 'coupon优惠券', 'FBA月订阅费(平台店租)',
     '其他服务费', '平台其他支出汇总', '计算结果-广告费', '计算结果-赔偿', '计算结果-其他分摊费用']]

# 手动检查（自动化）：映射站点、映射平台 是否有空值/空字符串
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
    print(f"{Color.RED} --- ====== [错误]映射结果有空值，请先修复映射表后再继续 ====== --- {Color.RESET}")
    for _col, _info in _missing_info.items():
        print(f"{Color.YELLOW}[缺失]{Color.RESET} 列：{_col}，空值行数：{_info['count']}")
        print(_info["preview"].to_string(index=False))
    raise SystemExit(1)
else:
    print(f"{Color.GREEN} --- ====== [一切正常]，进入下一步（保存文件） ====== --- {Color.RESET}")

# 保存修改
output_path = fr"{SELLERSKU_PROFIT_REPORT_DIR}\(已完成-1){SELLERSKU_PROFIT_FILE_NAME}"
try:
    main_file_df_3.to_excel(output_path, index=False)
except PermissionError:
    print(f"{Color.RED}保存失败：目标文件被占用/无权限。请先关闭已打开的输出文件后重试：{output_path}{Color.RESET}")
    raise
print(f'处理完成，文件另存为：{output_path}')
