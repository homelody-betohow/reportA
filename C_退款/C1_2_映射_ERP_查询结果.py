import importlib.util
import warnings
from pathlib import Path

# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

import pandas as pd
from A_报表.Z_method.style import Color
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")
# TODO 文件路径！！！
RMA_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\RMA\RMA-{shared_date}.xlsx"
RMA_file_df = pd.read_excel(RMA_file_path, skiprows=2)  # 跳过前2行 如果文件读取失败的话，则手动删掉多余的列，只保留想要的列
# 去除 整张表 的前后空格
for col in RMA_file_df.columns:
    RMA_file_df[col] = RMA_file_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
# 检查 是否有 "退款状态" != "作废"，且 退款金额 == 0
# 筛选出“退款状态”列中不 等于 “作废”的行
RMA_file_df_1 = RMA_file_df[RMA_file_df["退款状态"] != "作废"].copy()
# 1. 找出所有退款金额为 0 的行
zero_mask = RMA_file_df_1['退款金额'] == 0
# 2. 如果有 0 值，主动报错并返回对应的“退款原退款原订单号”
if zero_mask.any():
    bad_orders = RMA_file_df_1.loc[zero_mask, '退款原订单号'].tolist()
    raise ValueError(
        f'退款金额列存在 0 值，对应"退款原订单号"：{bad_orders}，询问相应运营，是否忘记标记、审核，退款订单(TEMU退款，一帆录入，一帆审核) ！！！')

# LM_BC、LM_RP 的  退款订单  的 平台SKU 映射
shops = {
    'LM_BC_FR': r'\\Betohow\数据报表\RPA\报表-无站点-订单查询\LM-BC-退款\LM-BC-退款-订单管理-近6个月.csv',
    'LM_RP_FR': r'\\Betohow\数据报表\RPA\报表-无站点-订单查询\LM-RP-退款\LM-RP-退款-订单管理-近6个月.csv',
}
have_order_shop_list = []
RMA_file_df_1['平台sku'] = ''
for shop, path in shops.items():
    mask = RMA_file_df_1['店铺英文名'] == shop
    if not mask.any():
        print(f"\n{shop}，没有退款订单！\n")
        continue
    have_order_shop_list.append(shop)
    df_order = pd.read_csv(path)
    for c in ['销售参考号', 'SKU', '仓库SKU']:
        df_order[c] = df_order[c].str.replace(r'^="(.*)"$', r'\1', regex=True)
    # 销售参考号+仓库SKU 复合键 → 平台 SKU（避免同一订单多行 SKU 时索引重复）
    df_order['_map_key'] = df_order['销售参考号'].astype(str) + '||' + df_order['仓库SKU'].astype(str)
    sku_map = df_order.set_index('_map_key')['SKU']
    if sku_map.index.duplicated().any():
        dup_keys = sku_map.index[sku_map.index.duplicated(keep=False)].unique().tolist()
        raise ValueError(f'{shop} ERP 表中「销售参考号+仓库SKU」仍不唯一：{dup_keys[:20]}')
    rma_map_key = (
        RMA_file_df_1['退款原订单号'].astype(str) + '||' + RMA_file_df_1['RMA产品'].astype(str)
    )
    matched = mask & rma_map_key.isin(sku_map.index)
    RMA_file_df_1.loc[matched, '平台sku'] = rma_map_key[matched].map(sku_map)
    RMA_file_df_1.loc[matched, '退款原订单号'] += '——已映射"' + RMA_file_df_1.loc[matched, '平台sku'] + '"'

# 保留指定列
RMA_file_df_1 = RMA_file_df_1[
    ['平台', '店铺英文名', '订单目的国家', '退款原订单号', '平台sku', 'RMA产品', 'RMA产品数量', '退款金额', '退款状态']]

# 保存结果
output_path = RMA_file_path.rsplit('\\', 1)[0] + '\\(已完成-1)' + RMA_file_path.rsplit('\\', 1)[-1]
RMA_file_df_1.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
for have_order_shop in have_order_shop_list:
    print(f'{Color.YELLOW}~~~~~~~~~~~~~~~~~请检查，"店铺英文名" == {have_order_shop}，"退款订单"是否都已映射 "平台SKU"！！！{Color.RESET}')
