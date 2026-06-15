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
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-1-1)订单统计-{shared_date}.xlsx"
main_df = pd.read_excel(main_file_path)

# 去除 整张表 的前后空格
for col in main_df.columns:
    main_df[col] = main_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

# 定义例外列表  不去掉尾缀的仓库SKU
exceptions = ['XPYN2125D-1', 'EXPYN2125D-1', 'EBS8029-1']  # 这些SKU本身就是这样
main_df['仓库SKU'] = main_df['仓库SKU'].where(
    main_df['仓库SKU'].isin(exceptions),
    # 非例外的执行清理
    # 仓库SKU  去掉尾缀 -1、-2、-3、-4、-5、_NEVER_USED、-AT    去掉开头的 EXM
    main_df['仓库SKU']
    .str.replace(r'(-1|-2|-3|-4|-5|_NEVER_USED|-AT)$', '', regex=True)
    .str.replace(r'^EXM', '', regex=True)
)

# 重命名
main_df = main_df.rename(columns={'仓库SKU': '原-仓库SKU'})
# 使用正则表达式匹配AMZN.GR.和_FB之间的任意字符
main_df['仓库SKU'] = main_df['原-仓库SKU'].str.extract(r'AMZN\.GR\.(.*?)_FB')
# 如果无法提取，保留原值
main_df['仓库SKU'] = main_df['仓库SKU'].fillna(main_df['原-仓库SKU'])

product_map_sku_path = fr"{DESKTOP_ROOT}\站点-匹配表.xlsx"  # 改成对应的映射表
# 映射 平台
main_df_2 = sku_mappings(
    main_df=main_df,
    main_sku='店铺英文名',
    map_sku_path=product_map_sku_path,
    map_old_sku="平台账号",
    map_new_sku="平台",
    map_sku_sheet='站点匹配'
)
# 构建 '店铺英文名-站点'
main_df_2['店铺英文名-站点'] = main_df_2["店铺英文名"] + '-' + main_df_2["站点"]

# 映射站点
# 1) 先按“店铺英文名-站点” → 站点 映射
main_df_3 = sku_mappings(
    main_df=main_df_2,
    main_sku='店铺英文名-站点',
    map_sku_path=product_map_sku_path,
    map_old_sku='特殊-平台账号',
    map_new_sku='站点',
    map_sku_sheet='站点匹配'
)

# 2) 对空的记录再用“店铺英文名” → 站点 兜底
mask = main_df_3['映射站点'].isna()
# 先“映射站点”列删掉，否则会冲突
to_map = (
    main_df_3.loc[mask]
    .drop(columns=['映射站点'], errors='ignore')
)
tmp_df_2 = sku_mappings(
    main_df=to_map,
    main_sku='店铺英文名',
    map_sku_path=product_map_sku_path,
    map_old_sku='平台账号',
    map_new_sku='站点',
    map_sku_sheet='站点匹配'
)
# 回填
main_df_3.loc[mask, '映射站点'] = tmp_df_2['映射站点']

# LM_BC_FR 站点，单独处理
# 1. 拆分出需要处理的 LM_BC_FR
LM_BC = main_df_3[main_df_3['店铺英文名'] == 'LM_BC_FR'].copy()
# 2. 对 LM_BC_FR 的“映射站点”添加后缀("平台sku"以 "ls-"开头,则 在对应的 映射站点 后面加上 -ls；否则,加上-xj)
LM_BC['映射站点'] = LM_BC['平台sku'].apply(
    lambda x: '-ls' if x.startswith('ls-') else '-xj'
).radd(LM_BC['映射站点'])  # radd 把后缀拼到原字符串右侧
# 3. 把处理后的 LM_BC_FR 合并回原 df（按索引原地更新）
main_df_3.update(LM_BC[['映射站点']])

# LM_RP_FR 站点，单独处理
# 1. 拆分出需要处理的 LM_RP_FR
LM_BC = main_df_3[main_df_3['店铺英文名'] == 'LM_RP_FR'].copy()
# 2. 对 LM_RP_FR 的“映射站点”添加后缀("平台sku"以 "ls-"开头,则 在对应的 映射站点 后面加上 -ls；否则,加上-xj)
LM_BC['映射站点'] = LM_BC['平台sku'].apply(
    lambda x: '-ls' if x.startswith('ls-') else '-xj'
).radd(LM_BC['映射站点'])  # radd 把后缀拼到原字符串右侧
# 3. 把处理后的 LM_RP_FR 合并回原 df（按索引原地更新）
main_df_3.update(LM_BC[['映射站点']])

# 重命名列
main_df_3 = main_df_3.rename(columns={'仓库SKU': 'SKU'})
# 构建识别码
main_df_3['儿子-站点识别码'] = main_df_3['映射站点'] + main_df_3['SKU']
main_df_3['儿子-平台识别码'] = main_df_3['映射平台'] + main_df_3['SKU']

# 保留目标列
output_main_df_3 = main_df_3[
    ['平台', '店铺英文名', '站点', '店铺英文名-站点', '映射站点', '映射平台', '儿子-站点识别码', '儿子-平台识别码',
     '订单类型', '参考号', '订单号', '平台sku', 'SKU', '仓库','仓库属性', '运输方式', '仓库SKU销量', '跟踪单号', '币种',
     '订单总金额', '平台运费', 'fba费用', '头程运费', '头程税费', '派送运费']]

# 检查：映射站点、映射平台 是否有空值/空字符串
_check_cols = ['映射站点', '映射平台']
_preview_cols = ['店铺英文名', '站点', '店铺英文名-站点', '订单号', 'SKU']
_missing_info = {}
for _col in _check_cols:
    _mask = output_main_df_3[_col].isna() | (output_main_df_3[_col].astype(str).str.strip() == '')
    _cnt = int(_mask.sum())
    if _cnt > 0:
        _preview = [c for c in _preview_cols if c in output_main_df_3.columns] + [_col]
        _missing_info[_col] = {
            'count': _cnt,
            'preview': output_main_df_3.loc[_mask, _preview].drop_duplicates().head(20),
        }

if _missing_info:
    print(f"{Color.RED} --- ====== [错误]映射结果有空值，请先修复「站点-匹配表.xlsx」后再继续 ====== --- {Color.RESET}")
    for _col, _info in _missing_info.items():
        print(f"{Color.YELLOW}[缺失]{Color.RESET} 列：{_col}，空值行数：{_info['count']}")
        print(_info['preview'].to_string(index=False))
    raise SystemExit(1)

print(f"{Color.GREEN}['映射站点', '映射平台']检查通过，可进行下一步{Color.RESET}")

# 保存修改后的文件
output_path = main_file_path.replace('已完成-1-1', '已完成-2')
output_main_df_3.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
