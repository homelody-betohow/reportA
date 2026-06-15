import warnings
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

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")
# TODO 文件路径！！！
RMA_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\RMA\(已完成-1)RMA-{shared_date}.xlsx"
RMA_file_df = pd.read_excel(RMA_file_path)
# 去除 整张表 的前后空格
for col in RMA_file_df.columns:
    RMA_file_df[col] = RMA_file_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
# 筛选  “店铺英文名”不包含 ECO、Biancca、yiqianshangmao_DE 的行
RMA_file_df_1 = RMA_file_df[
    ~RMA_file_df['店铺英文名'].str.contains('ECO|Biancca|yiqianshangmao_DE', na=False)].copy()
# 定义例外列表  不去掉尾缀的仓库SKU
exceptions = ['XPYN2125D-1', 'EXPYN2125D-1', 'EBS8029-1']  # 这些SKU本身就是这样
RMA_file_df_1['RMA产品'] = RMA_file_df_1['RMA产品'].where(
    RMA_file_df_1['RMA产品'].isin(exceptions),
    # 非例外的执行清理
    # RMA产品  去掉尾缀 -1、-2、-3、-4、-5、_NEVER_USED、-AT    去掉开头的 EXM
    RMA_file_df_1['RMA产品']
    .str.replace(r'(-1|-2|-3|-4|-5|_NEVER_USED|-AT)$', '', regex=True)
    .str.replace(r'^EXM', '', regex=True)
)

product_map_sku_path = fr"{DESKTOP_ROOT}\站点-匹配表.xlsx"  # 改成对应的映射表
# 映射 平台
RMA_file_df_2 = sku_mappings(
    main_df=RMA_file_df_1,
    main_sku='店铺英文名',
    map_sku_path=product_map_sku_path,
    map_old_sku='平台账号',
    map_new_sku='平台',
    map_sku_sheet='站点匹配',
)
# 构建 '店铺英文名-站点'
RMA_file_df_2['店铺英文名-站点'] = RMA_file_df_2["店铺英文名"] + '-' + RMA_file_df_2["订单目的国家"]

# 映射站点
# 1) 先按“店铺英文名-站点” → 站点 映射
RMA_file_df_3 = sku_mappings(
    main_df=RMA_file_df_2,
    main_sku='店铺英文名-站点',
    map_sku_path=product_map_sku_path,
    map_old_sku='特殊-平台账号',
    map_new_sku='站点',
    map_sku_sheet='站点匹配'
)

# 2) 对空的记录再用“店铺英文名” → 站点 兜底
mask = RMA_file_df_3['映射站点'].isna()
# 先“映射站点”列删掉，否则会冲突
to_map = (
    RMA_file_df_3.loc[mask]
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
RMA_file_df_3.loc[mask, '映射站点'] = tmp_df_2['映射站点']

# LM_BC_FR 站点，单独处理
# 1. 拆分出需要处理的 LM_BC_FR
LM_BC = RMA_file_df_3[RMA_file_df_3['店铺英文名'] == 'LM_BC_FR'].copy()
# 2. 对 LM_BC_FR 的“映射站点”添加后缀("平台sku"以 "ls-"开头,则 在对应的 映射站点 后面加上 -ls；否则,加上-xj)
LM_BC['映射站点'] = LM_BC['平台sku'].apply(
    lambda x: '-ls' if x.startswith('ls-') else '-xj'
).radd(LM_BC['映射站点'])  # radd 把后缀拼到原字符串右侧
# 3. 把处理后的 LM_BC_FR 合并回原 df（按索引原地更新）
RMA_file_df_3.update(LM_BC[['映射站点']])

# LM_RP_FR 站点，单独处理
# 1. 拆分出需要处理的 LM_RP_FR
LM_BC = RMA_file_df_3[RMA_file_df_3['店铺英文名'] == 'LM_RP_FR'].copy()
# 2. 对 LM_RP_FR 的“映射站点”添加后缀("平台sku"以 "ls-"开头,则 在对应的 映射站点 后面加上 -ls；否则,加上-xj)
LM_BC['映射站点'] = LM_BC['平台sku'].apply(
    lambda x: '-ls' if x.startswith('ls-') else '-xj'
).radd(LM_BC['映射站点'])  # radd 把后缀拼到原字符串右侧
# 3. 把处理后的 LM_RP_FR 合并回原 df（按索引原地更新）
RMA_file_df_3.update(LM_BC[['映射站点']])

# 重命名
RMA_file_df_3 = RMA_file_df_3.rename(columns={'RMA产品': 'SKU'})
# 构建识别码
RMA_file_df_3['儿子-站点识别码'] = RMA_file_df_3['映射站点'] + RMA_file_df_3['SKU']
RMA_file_df_3['儿子-平台识别码'] = RMA_file_df_3['映射平台'] + RMA_file_df_3['SKU']
# 保留指定列
RMA_file_df_3 = RMA_file_df_3[
    ['平台', '店铺英文名', '订单目的国家', '店铺英文名-站点', '映射站点', '映射平台', '儿子-站点识别码',
     '儿子-平台识别码', '退款原订单号', '平台sku', 'SKU', 'RMA产品数量', '退款金额', '退款状态']]

# 保存结果
output_path = RMA_file_path.replace('已完成-1', '已完成-1-1')
RMA_file_df_3.to_excel(output_path, index=False)
print(f'处理完成，文件另存为：{output_path}')
print(f"{Color.YELLOW}~~~~~~~~~~~~~手动检查， ['映射站点', '映射平台']，是否有空的！！！{Color.RESET}")
