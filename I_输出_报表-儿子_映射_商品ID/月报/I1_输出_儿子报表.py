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
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-13-1)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)
# 删掉 其他分摊费用 的计算列
main_file_df.drop(columns=['原-其他分摊费用', 'EU-其他分摊费用-需要分摊的', 'US-其他分摊费用-需要分摊的'], inplace=True)

# 替换 相同产品的SKU  避免产品信息库 映射不到
main_file_df.loc[main_file_df['SKU'] == 'CY9901', 'SKU'] = 'ECY9901'  # CY9901 替换成 ECY9901

# 映射 供应商
product_map_sku_path = r"\\Betohow\数据报表\数据库\BTH全部SKU明细-v2026.06.02.xlsx"

main_file_df = sku_mappings(
    main_df=main_file_df,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="SKU",
    map_new_sku="供应商",
    map_sku_sheet='基础数据维护'
)
main_file_df = main_file_df.rename(columns={'映射供应商': '供应商'})

product_map_sku_path = r"\\Betohow\数据报表\数据库\产品信息库2025.xlsx"  # 改成对应的映射表
# 映射 商品ID
main_file_df_1 = sku_mappings(
    main_df=main_file_df,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="产品编码",
    map_new_sku="商品ID",
    map_sku_sheet='产品信息表'
)
# 重命名列
main_file_df_1 = main_file_df_1.rename(columns={'映射商品ID': '商品ID'})

main_file_df_1["站点商品ID识别码"] = main_file_df_1["站点"] + main_file_df_1["商品ID"]  # 新列数据
main_file_df_1["平台商品ID识别码"] = main_file_df_1["平台"] + main_file_df_1["商品ID"]  # 新列数据

# 映射产品状态
# 拆分数据  平台 是否包含 AMAZON
df_amazon = main_file_df_1[main_file_df_1['平台'].str.contains('AMAZON', case=False, na=False)]
df_not_amazon = main_file_df_1[~main_file_df_1['平台'].str.contains('AMAZON', case=False, na=False)]
# 映射 AMAZON 的 产品状态
df_amazon_1 = sku_mappings(
    main_df=df_amazon,
    main_sku='商品ID',
    map_sku_path=product_map_sku_path,
    map_old_sku="商品ID",
    map_new_sku="AMZ新老品",
    map_sku_sheet='产品信息表'
)
df_amazon_1 = df_amazon_1.rename(columns={'映射AMZ新老品': '产品状态'})
# 映射 非MAZON 的 产品状态
df_not_amazon_1 = sku_mappings(
    main_df=df_not_amazon,
    main_sku='商品ID',
    map_sku_path=product_map_sku_path,
    map_old_sku="商品ID",
    map_new_sku="本土平台新老品",
    map_sku_sheet='产品信息表'
)
df_not_amazon_1 = df_not_amazon_1.rename(columns={'映射本土平台新老品': '产品状态'})
# 合并数据
main_file_df_2 = pd.concat([df_amazon_1, df_not_amazon_1]).reset_index(drop=True)

# 再次映射产品状态，映射自己记录的分销。以及其它一些产品的产品状态
product_map_sku_path = fr'{DESKTOP_ROOT}\信息-映射.xlsx'
# 筛选产品状态为空的行
no_state_df = main_file_df_2[main_file_df_2['产品状态'].isna()]
no_state_df_1 = sku_mappings(
    main_df=no_state_df,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="SKU",
    map_new_sku="产品状态",
    map_sku_sheet='产品状态'
)

# 用 "映射产品状态" 替换 "产品状态"
no_state_df_1.loc[:, '产品状态'] = no_state_df_1['映射产品状态']
# 删除 "映射产品状态" 列
no_state_df_1 = no_state_df_1.drop(columns=['映射产品状态'])
# 将筛选的 no_state_df，合并回原始的 main_file_df_2
main_file_df_2.update(no_state_df_1)

# 将 SKU 以 "U88" 开头的行的“产品状态”改为“新品”
main_file_df_2.loc[main_file_df_2['SKU'].astype(str).str.startswith('U88'), '产品状态'] = '新品'
# 将 SKU 以 "-NW" 结尾的行的“产品状态”改为“不保留老品”
main_file_df_2.loc[main_file_df_2['SKU'].astype(str).str.endswith('-NW'), '产品状态'] = '不保留老品'
# 使用正则表达式匹配以'25'或'SN25'或'207'开头的SKU，将匹配到的记录对应的'产品状态'统一替换成'分销'
main_file_df_2.loc[main_file_df_2['SKU'].astype(str).str.match(r'^(25|SN25|207)'),'产品状态'] = '分销'
# 如果供应商是“智慧谷”，则将“采购成本”、“订单采购成本”、“重发采购成本”、“二次上架采购成本”替换为 0
main_file_df_2.loc[
    main_file_df_2['供应商'] == '智慧谷', ['采购成本', '订单采购成本', '重发采购成本', '二次上架采购成本']] = 0

# 表头 重新排序
main_file_df_2 = main_file_df_2[[
    "商品ID",
    "SKU",
    "站点",
    "平台",
    "SKU-站点识别码",
    "SKU-平台识别码",
    "平台商品ID识别码",
    "站点商品ID识别码",
    "产品状态",
    "销量",
    "平台销售额",
    "退款数量",
    "重发数量",
    "退款额",
    "销售额",
    "测评费",
    "秒杀费",
    "广告费(AMZ)",
    "广告费(非AMZ)",
    "广告费合计",
    "平台费(AMZ)",
    "平台费(非AMZ)",
    "平台费合计",
    "销售税(AMZ)",
    "销售税(非AMZ)",
    "销售税合计",
    "派送费",
    "提现费",
    "赔偿金额",
    "其他分摊费用",
    "二次上架数量",
    "二次上架金额",
    "订单采购成本",
    "重发采购成本",
    "二次上架采购成本",
    "采购成本",
    "头程",
    "关税",
    "供应商"
]]

# 保存修改后的文件
output_path = main_file_path.rsplit('\\', 2)[0] + f'\\{shared_date}--{folder_name}-无仓租(SKU版).xlsx'
main_file_df_2.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
print(f'{Color.RED}“产品状态”是否都有数据，没有的部分手动去判断、填写进去！！！{Color.RESET}')
