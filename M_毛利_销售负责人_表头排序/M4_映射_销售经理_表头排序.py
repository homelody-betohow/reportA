import pandas as pd
from datetime import datetime
import importlib.util
from pathlib import Path
# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.Z_method.sku_映射 import sku_mappings
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-22)订单统计-{shared_date}.xlsx"
main_df = pd.read_excel(main_file_path)

# 映射 销售经理
product_map_sku_path = fr"{DESKTOP_ROOT}\信息-映射.xlsx"  # 改成对应的映射表
main_df_1 = sku_mappings(
    main_df=main_df,
    main_sku='销售负责人',
    map_sku_path=product_map_sku_path,
    map_old_sku="销售负责人",
    map_new_sku="销售经理",
    map_sku_sheet='销售负责人'
)
main_df_1 = main_df_1.rename(columns={'映射销售经理': '销售经理'})
# CD平台  的  销售经理  替换为 空
main_df_1.loc[main_df_1['平台'] == 'CD', '销售经理'] = ''
# 新增列：仓租识别码，平台+产品状态
main_df_1['仓租识别码'] = main_df_1['平台'] + main_df_1['产品状态'].astype(str)

# 表头 重新排序
main_df_1 = main_df_1[[
    "商品ID",
    "SKU",
    "站点",
    "平台",
    "平台商品ID识别码",
    "站点商品ID识别码",
    "仓租识别码",
    "产品状态",
    "二级分类",
    "三级分类",
    "销售经理",
    "销售负责人",
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
    "海外仓仓租费",
    "FBA仓租费",
    "仓租合计",
    "提现费",
    "月租",
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
    "毛利",
    "毛利率",
    "运营模式",
    "供应商"
]]

# 空值的地方——补 0
main_df_1 = main_df_1.fillna(0)
# 下面这些列，值为 0 的地方，替换为 空
cols = ['商品ID', 'SKU', '平台', '站点', '平台商品ID识别码', '站点商品ID识别码', '仓租识别码', '产品状态',
        '二级分类', '三级分类', '销售经理', '销售负责人', '运营模式', '供应商']
main_df_1[cols] = main_df_1[cols].replace(0, '')

# 刷新一下 仓租识别码
main_df_1['仓租识别码'] = main_df_1['平台'] + main_df_1['产品状态'].astype(str)

if folder_name == '日报':
    # 获取日报的完整日期
    current_year = datetime.now().year  # 获取当前年份
    # 解析日期
    day_11 = shared_date.split('-')[0]
    day_22 = shared_date.split('-')[1]
    month_1, day_1 = map(int, day_11.split('.'))
    month_2, day_2 = map(int, day_22.split('.'))
    # 格式化为完整的日期字符串
    full_date = f"{month_1}月{day_1}日-{month_2}月{day_2}日"
    print(full_date)
    # 找到“商品ID”列的索引位置
    product_id_index = main_df_1.columns.get_loc('商品ID')
    # 在“商品ID”列前插入“日期”列
    main_df_1.insert(product_id_index, '日期', full_date)
    # 删除 指定列   日报 不要 '订单采购成本', '重发采购成本', '二次上架采购成本'
    main_df_1.drop(columns=['订单采购成本', '重发采购成本', '二次上架采购成本'], inplace=True)
    # 映映  平台（报表）
    product_map_sku_path = fr'{DESKTOP_ROOT}\信息-映射.xlsx'
    main_df_1 = sku_mappings(
        main_df=main_df_1,
        main_sku='站点',
        map_sku_path=product_map_sku_path,
        map_old_sku="站点",
        map_new_sku="平台（报表）",
        map_sku_sheet='平台（报表）-平台映射'
    )
    # 重命名
    main_df_1 = main_df_1.rename(columns={'映射平台（报表）': '平台（报表）'})
    main_df_1 = main_df_1.rename(columns={'仓租识别码': '平台（报表识别码）'})
    main_df_1['平台（报表识别码）'] = main_df_1['平台（报表）'] + main_df_1['商品ID'].astype(str)

# 去除 整张表 的前后空格
for col in main_df_1.columns:
    main_df_1[col] = main_df_1[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

# 保存结果
output_path = main_file_path.rsplit('\\', 2)[0] + f'\\{shared_date}--{folder_name}.xlsx'
main_df_1.to_excel(output_path, index=False)
print(f"结果已保存到文件：{output_path}")
