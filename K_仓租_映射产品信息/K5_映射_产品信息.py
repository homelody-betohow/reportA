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
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-16)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)

# 运营模式、供应商：使用 BTH全部SKU明细（BTH_ALL_SKU_DETAIL_PATH）
bth_sku_detail_path = BTH_ALL_SKU_DETAIL_PATH
# 映射 运营模式
main_file_df_1 = sku_mappings(
    main_df=main_file_df,
    main_sku='SKU',
    map_sku_path=bth_sku_detail_path,
    map_old_sku="SKU",
    map_new_sku="运营模式",
    map_sku_sheet='基础数据维护'
)
main_file_df_1 = main_file_df_1.rename(columns={'映射运营模式': '运营模式'})
# 映射 供应商
main_file_df_2 = sku_mappings(
    main_df=main_file_df_1,
    main_sku='SKU',
    map_sku_path=bth_sku_detail_path,
    map_old_sku="SKU",
    map_new_sku="供应商",
    map_sku_sheet='基础数据维护'
)
main_file_df_2 = main_file_df_2.rename(columns={'映射供应商': '供应商'})

# 二级分类、三级分类、产品状态：使用 产品信息库2025.xlsx
product_map_sku_path = r"\\Betohow\数据报表\数据库\产品信息库2025.xlsx"

# 映射 二级分类
main_file_df_3 = sku_mappings(
    main_df=main_file_df_2,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="产品编码",
    map_new_sku="二级分类",
    map_sku_sheet='产品信息表'
)
# 重命名
main_file_df_3 = main_file_df_3.rename(columns={'映射二级分类': '二级分类'})
# 映射 三级分类
main_file_df_4 = sku_mappings(
    main_df=main_file_df_3,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="产品编码",
    map_new_sku="三级分类",
    map_sku_sheet='产品信息表'
)
main_file_df_4 = main_file_df_4.rename(columns={'映射三级分类': '三级分类'})

# 映射产品状态
# 拆分数据  平台 是否包含 AMAZON
df_amazon = main_file_df_4[main_file_df_4['平台'].str.contains('AMAZON', case=False, na=False)]
df_not_amazon = main_file_df_4[~main_file_df_4['平台'].str.contains('AMAZON', case=False, na=False)]
# 映射 AMAZON 的 产品状态
df_amazon_1 = sku_mappings(
    main_df=df_amazon,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="产品编码",
    map_new_sku="AMZ新老品",
    map_sku_sheet='产品信息表'
)
df_amazon_1 = df_amazon_1.rename(columns={'映射AMZ新老品': '产品状态'})
# 映射 非MAZON 的 产品状态
df_not_amazon_1 = sku_mappings(
    main_df=df_not_amazon,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="产品编码",
    map_new_sku="本土平台新老品",
    map_sku_sheet='产品信息表'
)
df_not_amazon_1 = df_not_amazon_1.rename(columns={'映射本土平台新老品': '产品状态'})
# 合并数据
main_file_df_5 = pd.concat([df_amazon_1, df_not_amazon_1]).reset_index(drop=True)

# 再次映射产品状态，映射自己记录的分销。以及其它一些产品的产品状态
product_map_sku_path = fr'{DESKTOP_ROOT}\信息-映射.xlsx'
# 筛选产品状态为空的行
no_state_df = main_file_df_5[main_file_df_5['产品状态'].isna()]
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
# 将筛选的 no_state_df，合并回原始的 main_file_df_5
main_file_df_5.update(no_state_df_1)


# 产品状态为空，且 分销为 是 的行，产品状态改为 分销
main_file_df_5.loc[(main_file_df_5['产品状态'].isna()) & (main_file_df_5['分销'] == '是'), '产品状态'] = '分销'


# 将 SKU 以 "U88" 开头的行的“产品状态”改为“新品”; “二级分类”、“三级分类”替换为 “其他”
main_file_df_5.loc[main_file_df_5['SKU'].astype(str).str.startswith('U88'), ['产品状态', '二级分类', '三级分类']] = [
    '新品', '其他', '其他']
# 将 SKU 以 "-NW" 结尾的行的“产品状态”改为“不保留老品”
main_file_df_5.loc[main_file_df_5['SKU'].astype(str).str.endswith('-NW'), '产品状态'] = '不保留老品'
# 使用正则表达式匹配以'25'或'SN25'或'207'开头的SKU，将匹配到的记录对应的'供应商'统一替换成'智慧谷'
main_file_df_5.loc[main_file_df_5['SKU'].astype(str).str.match(r'^(25|SN25|207)'), '供应商'] = '智慧谷'
# 检查供应商是否为“易速”或“智慧谷”，并替换对应的“产品状态”、“二级分类”、“三级分类”为 “分销”
main_file_df_5.loc[main_file_df_5['供应商'].isin(['易速', '智慧谷']), ['产品状态', '二级分类', '三级分类']] = '分销'
# 如果产品状态是“分销”，则将“运营模式”替换为 “自运营”; “二级分类”、“三级分类”替换为 “分销”
main_file_df_5.loc[main_file_df_5['产品状态'] == '分销', ['运营模式', '二级分类', '三级分类']] = ['自运营', '分销',
                                                                                                  '分销']
# 保存修改后的文件
output_path = main_file_path.replace('已完成-16', '已完成-17')
try:    
    main_file_df_5.to_excel(output_path, index=False)
    print(f'处理完成，文件另存为：{output_path}')
except PermissionError:
    # 常见原因：目标文件正在被 Excel 打开占用
    output_path_2 = output_path.replace(".xlsx", "-另存.xlsx")
    main_file_df_5.to_excel(output_path_2, index=False)
    print(f'处理完成（原文件被占用，已另存），文件为：{output_path_2}')
print(f"{Color.YELLOW}[注意]检查————供应商、运营模式、产品状态、三级分类，是不是都有了，没有的部分手动去判断、填写进去！！！{Color.RESET}")
print(f"新品 二、三级分类，空着；分销 供应商目前都是：智慧谷！")
if folder_name == '月报':
    print(f"{Color.RED}~~~~~~~~~~~~~~~~~站点：AMAZON-UK，如果只有 '仓租'，就把'仓租'放到AMAZON-DE，然后删掉 AMAZON-UK！{Color.RESET}")
