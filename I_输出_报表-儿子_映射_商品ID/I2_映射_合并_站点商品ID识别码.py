import pandas as pd
import importlib.util
from pathlib import Path
# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

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
main_file_df_1 = main_file_df_1.rename(columns={'SKU': '原-SKU'})

#  商品ID 去映射 产品信息库 的 第一个 产品编码（SKU）
main_file_df_2 = sku_mappings(
    main_df=main_file_df_1,
    main_sku='商品ID',
    map_sku_path=product_map_sku_path,
    map_old_sku="商品ID",
    map_new_sku="产品编码",
    map_sku_sheet='产品信息表'
)
# 重命名列
main_file_df_2 = main_file_df_2.rename(columns={'映射产品编码': 'SKU'})

main_file_df_2["站点商品ID识别码"] = main_file_df_2["站点"].astype(str) + main_file_df_2["商品ID"].astype(str)
main_file_df_2["平台商品ID识别码"] = main_file_df_2["平台"].astype(str) + main_file_df_2["商品ID"].astype(str)

# 按照 '站点商品ID识别码' 列进行分组，并对 费用 列进行汇总
grouped_df = main_file_df_2.groupby('站点商品ID识别码').agg({
    "商品ID": 'first',
    "SKU": 'first',
    "站点": 'first',
    "平台": 'first',
    "平台商品ID识别码": 'first',
    "分销": 'first',
    "销量": 'sum',
    "平台销售额": 'sum',
    "退款数量": 'sum',
    "重发数量": 'sum',
    "退款额": 'sum',
    "销售额": 'sum',
    "测评费": 'sum',
    "秒杀费": 'sum',
    "广告费(AMZ)": 'sum',
    "广告费(非AMZ)": 'sum',
    "广告费合计": 'sum',
    "平台费(AMZ)": 'sum',
    "平台费(非AMZ)": 'sum',
    "平台费合计": 'sum',
    "销售税(AMZ)": 'sum',
    "销售税(非AMZ)": 'sum',
    "销售税合计": 'sum',
    "派送费": 'sum',
    "提现费": 'sum',
    "赔偿金额": 'sum',
    "其他分摊费用": 'sum',
    "二次上架数量": 'sum',
    "二次上架金额": 'sum',
    "订单采购成本": 'sum',
    "重发采购成本": 'sum',
    "二次上架采购成本": 'sum',
    "采购成本": 'sum',
    "头程": 'sum',
    "关税": 'sum',
}).reset_index()

# 保存修改后的文件
output_path = main_file_path.replace('已完成-13-1', '已完成-14')
grouped_df.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
