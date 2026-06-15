import numpy as np
import pandas as pd
import importlib.util
from pathlib import Path
# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.Z_method.style import Color
from A_报表.Z_method.sku_映射 import sku_mappings
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import BTH_ALL_SKU_DETAIL_PATH, DESKTOP_ROOT

# TODO 文件路径！！！
main_file_path = fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\二次上架\(已完成-1)鸿羽仓-二次上架明细-{shared_date}.xlsx'
main_df = pd.read_excel(main_file_path)

# 映射 平台
product_map_sku_path = fr"{DESKTOP_ROOT}\站点-匹配表.xlsx"  # 改成对应的映射表
main_df_1 = sku_mappings(
    main_df=main_df,
    main_sku='映射站点',
    map_sku_path=product_map_sku_path,
    map_old_sku="站点",
    map_new_sku="平台",
    map_sku_sheet='站点匹配'
)


# 通用处理函数：处理 LM_BC_FR 、 LM_RP_FR 的二次上架订单映射及站点后缀
def process_lm_orders(df, account_name, order_file_path):
    # 筛选目标账号的订单
    mask = df['合并-映射账号'] == account_name
    print(f"\n{account_name}的二次上架订单:")
    # 读取订单管理表，获取原始订单的SKU映射
    order_df = pd.read_excel(order_file_path)
    for col in ['销售参考号', 'SKU']:
        order_df[col] = order_df[col].str.replace(r'^="(.*)"$', r'\1', regex=True)  # 清理CSV引号格式
    mapping = order_df.set_index('销售参考号')['SKU'].to_dict()  # 创建{销售参考号: SKU}映射字典
    # 向量化：优先用订单参考号查 SKU，查不到再用参考号（与 G2 列含义一致）
    orig_order = df.loc[mask, '订单参考号']
    orig_ref = df.loc[mask, '参考号'] if '参考号' in df.columns else pd.Series(np.nan, index=orig_order.index)
    mapped_from_order = orig_order.map(mapping)
    mapped_from_ref = orig_ref.map(mapping)
    mapped_sku = mapped_from_order.combine_first(mapped_from_ref)
    found_mask = mapped_sku.notna()
    matched_by_order = mapped_from_order.notna()
    matched_by_ref = mapped_from_ref.notna() & ~matched_by_order
    df.loc[mask, '平台sku'] = mapped_sku.combine_first(df.loc[mask, '平台sku'])
    print(
        f"[OK] 已映射 {int(found_mask.sum())} 行"
        f"（订单参考号 {int(matched_by_order.sum())}，参考号 {int(matched_by_ref.sum())}），"
        f"未找到 {int((~found_mask).sum())} 行"
    )
    # 根据平台SKU前缀，为映射站点添加后缀（ls-开头加-ls，否则加-xj）
    lm_df = df[mask].copy()
    suffix = np.where(lm_df['平台sku'].astype(str).str.startswith('ls-'), '-ls', '-xj')
    lm_df['映射站点'] = lm_df['映射站点'].astype(str) + suffix
    df.update(lm_df[['映射站点']])  # 更新回原DataFrame


# 执行处理：对两个法国账号进行同样的操作
path = r'\\Betohow\数据报表\RPA\二次上架-数据查询\订单管理\all-订单管理查询.xlsx'
# 初始化平台sku列
main_df_1['平台sku'] = ''
lm_bc = 'LM_BC_FR' in main_df_1['合并-映射账号'].values  # 合并-映射账号 中有 LM_BC_FR，则 lm_bc为True，否则 False
lm_rp = 'LM_RP_FR' in main_df_1['合并-映射账号'].values  # 合并-映射账号 中有 LM_RP_FR，则 lm_rp为True，否则 False
for account in ['LM_BC_FR', 'LM_RP_FR']:
    if account in main_df_1['合并-映射账号'].values:
        process_lm_orders(main_df_1, account, path)

# 重命名
main_df_1 = main_df_1.rename(columns={'映射站点': '站点'})
main_df_1 = main_df_1.rename(columns={'映射平台': '平台'})
# 构建识别码
main_df_1['儿子-站点识别码'] = main_df_1['站点'] + main_df_1['SKU']
main_df_1['儿子-平台识别码'] = main_df_1['平台'] + main_df_1['SKU']

# TODO 处理——退件费用(EUR) 按 实收数量 占比分摊
# 0. 确保分摊运费列存在且先置 0
main_df_1['分摊运费(EUR)'] = 0.0
# ---------- 1. OTTO 运费分摊 ----------
otto_mask = main_df_1['站点'].str.contains('OTTO', na=False)
otto_df = main_df_1[otto_mask].copy()
# 判断 otto_df 不为空
if not otto_df.empty:
    otto_df['退件类型'] = otto_df['退件类型'].fillna('')


    def get_otto_fee(return_type: str) -> float:
        if pd.isna(return_type):
            raise ValueError("退件类型为 NaN，无法处理。")
        # 客户退件：买家退件、认领
        if return_type.startswith(('买家退件', '认领')):
            return 6.1
        # 服务商退件：物流退件
        if return_type.startswith('物流退件'):
            return 5.5
        raise ValueError(f"不支持的退件类型: '{return_type}'")


    # 每个退件号整单运费
    fee_per_return = (otto_df.groupby('退件号')['退件类型']
                      .first()
                      .apply(get_otto_fee)
                      .rename('total_fee'))
    otto_df = otto_df.merge(fee_per_return, left_on='退件号', right_index=True)
    # 按实收数量占比分摊
    otto_df['qty_sum'] = otto_df.groupby('退件号')['实收数量'].transform('sum')
    otto_df['ratio'] = otto_df['实收数量'] / otto_df['qty_sum']
    otto_df['分摊运费(EUR)'] = otto_df['total_fee'] * otto_df['ratio']
    # 写回原表
    main_df_1.loc[otto_mask, '分摊运费(EUR)'] = otto_df['分摊运费(EUR)'].values

# ---------- 2. 实际退件费用（RMB）分摊 ----------
# 取每个退件号的第一行费用作为“整单费用”
fee_first = main_df_1.groupby('退件号')['退件费用(RMB)'].transform('first')
# 整单实收数量
qty_sum_all = main_df_1.groupby('退件号')['实收数量'].transform('sum')
ratio = main_df_1['实收数量'] / qty_sum_all
main_df_1['实际-退件费用(RMB)'] = np.where(
    qty_sum_all == 0,  # 防止除 0
    0,
    ratio * fee_first
)
# ---------- 3. 币种转换 ----------
main_df_1['实际-退件费用(EUR)'] = main_df_1['实际-退件费用(RMB)'] / 7.3
main_df_1['退件费用(EUR)'] = main_df_1['实际-退件费用(EUR)'] + main_df_1['分摊运费(EUR)']
# ---------- 4. 列名整理 ----------
main_df_1.rename(columns={'退件费用(RMB)': '原-退件费用(RMB)'}, inplace=True)
# 获取 原始采购价
product_map_sku_path = BTH_ALL_SKU_DETAIL_PATH
main_df_2 = sku_mappings(
    main_df=main_df_1,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="SKU",
    map_new_sku="原始采购价",
    map_sku_sheet='基础数据维护'
)
# 将字符串转换为数值类型（空字符串或非数字字符串会转换为 NaN）
main_df_2['映射原始采购价'] = pd.to_numeric(main_df_2['映射原始采购价'], errors='coerce')

# 二次上架采购成本 计算规则
# 1. OTTO 平台：
# 1.1 良品：二次上架采购成本（RMB） = 映射原始采购价 * 实收数量
# 1.2 次品：二次上架采购成本（RMB） = 0
# 1.3 NW后缀的SKU：二次上架采购成本（RMB） = 0
# 2. 其它平台：
# 2.1 NW后缀的SKU：二次上架采购成本（RMB） = 0
# 2.2 非NW后缀的SKU：二次上架采购成本（RMB） = 映射原始采购价 * 实收数量
# 3. 物流商退件
# 3.1 物流商退件：二次上架采购成本（RMB） = 映射原始采购价 * 良品数量

print(f"{Color.YELLOW}二次上架采购成本 计算规则：{Color.RESET}")
print(f"{Color.GREEN}1. OTTO 平台：")
print(f"1.1 良品：二次上架采购成本（RMB） = 映射原始采购价 * 实收数量")
print(f"1.2 次品：二次上架采购成本（RMB） = 0")
print(f"1.3 NW后缀的SKU：二次上架采购成本（RMB） = 0")
print(f"2. 其它平台：")
print(f"2.1 NW后缀的SKU：二次上架采购成本（RMB） = 0")
print(f"2.2 非NW后缀的SKU：二次上架采购成本（RMB） = 映射原始采购价 * 实收数量")
print(f"{Color.RESET}")


# 初始化列
main_df_2['二次上架采购成本（RMB）'] = 0.0

# 确保必要列的数据类型正确
main_df_2['实收数量'] = pd.to_numeric(main_df_2['实收数量'], errors='coerce').fillna(0)
main_df_2['良品'] = pd.to_numeric(main_df_2['良品'], errors='coerce').fillna(0)
main_df_2['映射原始采购价'] = main_df_2['映射原始采购价'].fillna(0)

# 规则1：OTTO 平台
otto_platform_mask = main_df_2['平台'].str.contains('OTTO', na=False)

# 1.1 OTTO 平台 - 良品（良品=1）：采购成本 = 映射原始采购价 * 实收数量
otto_good_mask = otto_platform_mask & (main_df_2['良品'] >= 1)
main_df_2.loc[otto_good_mask, '二次上架采购成本（RMB）'] = (
    main_df_2.loc[otto_good_mask, '映射原始采购价'] * 
    main_df_2.loc[otto_good_mask, '良品']
)

# 1.2 OTTO 平台 - 次品（良品=0）：采购成本 = 0（已在初始化时设为0，这里无需额外处理）
otto_defect_mask = otto_platform_mask & (main_df_2['良品'] <= 0)
main_df_2.loc[otto_defect_mask, '二次上架采购成本（RMB）'] = 0

# 1.3 NW后缀的SKU：二次上架采购成本（RMB） = 0
nw_suffix_mask = otto_platform_mask & main_df_2['SKU'].astype(str).str.upper().str.endswith('NW')
main_df_2.loc[nw_suffix_mask, '二次上架采购成本（RMB）'] = 0

# 规则2：其它平台
other_platform_mask = ~otto_platform_mask

# 2.1 非OTTO平台 - NW后缀的SKU：采购成本 = 0
nw_suffix_mask = other_platform_mask & main_df_2['SKU'].astype(str).str.upper().str.endswith('NW')
main_df_2.loc[nw_suffix_mask, '二次上架采购成本（RMB）'] = 0

# 2.2 非OTTO平台 - 非NW后缀的SKU：采购成本 = 映射原始采购价 * 实收数量
non_nw_suffix_mask = other_platform_mask & ~main_df_2['SKU'].astype(str).str.upper().str.endswith('NW')
main_df_2.loc[non_nw_suffix_mask, '二次上架采购成本（RMB）'] = (
    main_df_2.loc[non_nw_suffix_mask, '映射原始采购价'] * 
    main_df_2.loc[non_nw_suffix_mask, '良品']
)

# 输出统计信息
print(f'\n{Color.YELLOW}=== 二次上架采购成本计算完成 ==={Color.RESET}')
# print(f'OTTO平台-良品：{int(otto_good_mask.sum())} 行')
# print(f'OTTO平台-次品：{int(otto_defect_mask.sum())} 行')
# print(f'其它平台-NW后缀SKU：{int(nw_suffix_mask.sum())} 行')
# print(f'其它平台-非NW后缀SKU：{int(non_nw_suffix_mask.sum())} 行')
# print(f'总采购成本（RMB）：{main_df_2["二次上架采购成本（RMB）"].sum():.2f}\n')

# 合并-映射账号 为 LM_BC_FR 或 LM_RP_FR，且平台sku为空的记录，对应的'儿子-站点识别码', '站点'置空
mask = (main_df_2['合并-映射账号'].isin(['LM_BC_FR', 'LM_RP_FR'])) & (main_df_2['平台sku'] == '')
main_df_2.loc[mask, ['儿子-站点识别码', '站点']] = np.nan

# 当 儿子-站点识别码 为空 且 合并-映射账号 为 LM_BC_FR 时，
# 自动执行你原来提示的 Excel VLOOKUP：
# - 儿子-站点识别码 = VLOOKUP(A21, [手动-二次映射.xlsx]二次上架-LM-BC-自发货!A:I, 9, FALSE)
# - 站点           = VLOOKUP(A21, [手动-二次映射.xlsx]二次上架-LM-BC-自发货!A:I, 7, FALSE)
manual_map_file_path = fr"{DESKTOP_ROOT}\手动-二次映射.xlsx"
manual_map_sheet = "二次上架-LM-BC-自发货"
try:
    manual_map_df = pd.read_excel(manual_map_file_path, sheet_name=manual_map_sheet, usecols="A:I")

    # VLOOKUP 的 A 列作为 key，G 列(第7列)是站点，I 列(第9列)是儿子-站点识别码
    def _norm_key(x) -> str:
        s = "" if pd.isna(x) else str(x)
        s = s.strip()
        s = pd.Series([s]).str.replace(r'^="(.*)"$', r"\1", regex=True).iloc[0]  # 去 Excel 的 ="xxx"
        # 避免把像 900008 读成 900008.0
        if s.endswith(".0") and s.replace(".", "", 1).isdigit():
            s = s[:-2]
        return s.strip()

    key_series = manual_map_df.iloc[:, 0].map(_norm_key)
    site_series = manual_map_df.iloc[:, 6]
    child_site_code_series = manual_map_df.iloc[:, 8]

    manual_site_map = dict(zip(key_series, site_series))
    manual_child_site_code_map = dict(zip(key_series, child_site_code_series))

    lm_bc_fill_mask = (
        (main_df_2['合并-映射账号'] == 'LM_BC_FR')
        & (main_df_2['儿子-站点识别码'].isna() | (main_df_2['儿子-站点识别码'].astype(str).str.strip() == ''))
    )

    # 查找键兜底（按你的 Excel 公式优先级来）：
    # 1) 优先用“退件号”(A列) —— 你手工 VLOOKUP 就是 A21
    # 2) 若匹配不到，再用“订单参考号”(E列)；若有“——已映射”标记则只取前半段
    # 3) 若仍匹配不到，再用“参考号”(D列)
    key_from_rma = main_df_2.loc[lm_bc_fill_mask, '退件号'].map(_norm_key) if '退件号' in main_df_2.columns else pd.Series(
        index=main_df_2.loc[lm_bc_fill_mask].index, dtype=str
    )
    key_from_order_ref = main_df_2.loc[lm_bc_fill_mask, '订单参考号']
    key_from_order_ref = (
        key_from_order_ref.astype(str).str.split('——', n=1).str[0].map(_norm_key)
        if '订单参考号' in main_df_2.columns
        else pd.Series(index=main_df_2.loc[lm_bc_fill_mask].index, dtype=str)
    )
    key_from_ref = main_df_2.loc[lm_bc_fill_mask, '参考号'].map(_norm_key) if '参考号' in main_df_2.columns else pd.Series(
        index=main_df_2.loc[lm_bc_fill_mask].index, dtype=str
    )

    filled_site = (
        key_from_rma.map(manual_site_map)
        .combine_first(key_from_order_ref.map(manual_site_map))
        .combine_first(key_from_ref.map(manual_site_map))
    )
    filled_child_site_code = (
        key_from_rma.map(manual_child_site_code_map)
        .combine_first(key_from_order_ref.map(manual_child_site_code_map))
        .combine_first(key_from_ref.map(manual_child_site_code_map))
    )

    main_df_2.loc[lm_bc_fill_mask, '站点'] = main_df_2.loc[lm_bc_fill_mask, '站点'].combine_first(filled_site)
    main_df_2.loc[lm_bc_fill_mask, '儿子-站点识别码'] = main_df_2.loc[lm_bc_fill_mask, '儿子-站点识别码'].combine_first(
        filled_child_site_code
    )

    filled_cnt = int(filled_child_site_code.notna().sum())
    remain_cnt = int(main_df_2.loc[lm_bc_fill_mask, '儿子-站点识别码'].isna().sum())
    print(f'{Color.YELLOW}LM_BC_FR 手动映射回填：成功 {filled_cnt} 行，仍为空 {remain_cnt} 行（通常是手动表里没有对应 key）。{Color.RESET}')
except FileNotFoundError:
    print(f'{Color.YELLOW}未找到手动映射文件：{manual_map_file_path}，将跳过 LM_BC_FR 的手动VLOOKUP自动回填。{Color.RESET}')
except ValueError as e:
    print(f'{Color.YELLOW}读取手动映射文件失败（sheet/列范围可能不对）：{e}，将跳过 LM_BC_FR 的手动VLOOKUP自动回填。{Color.RESET}')

# 保存目标列
main_df_2 = main_df_2[
    ['退件号', '映射原始采购价', '订单号', '参考号', '订单参考号', '合并-映射账号', '站点', '平台', '儿子-站点识别码',
     '儿子-平台识别码', '平台sku', 'SKU', '实收数量', '良品', '原-退件费用(RMB)', '实际-退件费用(RMB)',
     '实际-退件费用(EUR)', '分摊运费(EUR)', '退件费用(EUR)', '二次上架采购成本（RMB）', '退件类型']]
# 将处理后的数据保存到新的Excel文件
output_file_path = main_file_path.replace('已完成-1', '已完成-2')
try:
    main_df_2.to_excel(output_file_path, index=False)  # index=False表示不保存索引列
except PermissionError:
    print(f'{Color.RED}保存失败：无法写入文件（权限被拒绝）{Color.RESET}')
    print(f'目标路径：{output_file_path}')
    print('请检查：')
    print('  1. 是否已在 Excel 中打开了该文件？请先关闭后重新运行。')
    print('  2. 文件是否为只读，或文件夹无写入权限。')
    raise SystemExit(1)
print(f"处理完成，结果已保存到{output_file_path}")
print(f'{Color.YELLOW} --- ====== 请检查，映射原始采购价 是否都有了！！！ ====== --- {Color.RESET}')
print('-' * 100)
if lm_bc:
    print(f'{Color.YELLOW} --- ====== 请检查，"合并-映射账号" == LM_BC_FR，"二次上架订单"是否都已映射 "平台SKU"！！！====== --- {Color.RESET}')
    print(f"{Color.YELLOW} --- ====== 没有映射'平台SKU'的，手动去替换 ['站点', '儿子-站点识别码']！！！====== --- {Color.RESET}")
    print(f" 儿子-站点识别码 {Color.RED} =VLOOKUP(A21,'C:\\Users\\BTH-windows\\Desktop\\[手动-二次映射.xlsx]二次上架-LM-BC-自发货'!$A:$I,9,FALSE){Color.RESET}")
    print(f" 站点 {Color.RED} =VLOOKUP(A21,'C:\\Users\\BTH-windows\\Desktop\\[手动-二次映射.xlsx]二次上架-LM-BC-自发货'!$A:$I,7,FALSE){Color.RESET}")
else:
    print(f'LM_BC_FR，没有 "二次上架订单"！！！')
print('-' * 100)
if lm_rp:
    print(f'{Color.YELLOW} --- ====== 请检查，"合并-映射账号" == LM_RP_FR，"二次上架订单"是否都已映射 "平台SKU"！！！====== --- {Color.RESET}')
    print(f"{Color.YELLOW} --- ====== 没有映射'平台SKU'的，手动去替换 ['站点', '儿子-站点识别码']！！！====== --- {Color.RESET}")
else:
    print('LM_RP_FR，没有 "二次上架订单"！！！')
print(f"{Color.RED}Amazon-退仓数据：二次上架采购成本（RMB） 设置为0{Color.RESET}")
