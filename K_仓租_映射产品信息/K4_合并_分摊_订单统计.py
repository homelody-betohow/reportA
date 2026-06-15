import pandas as pd
import importlib.util
from pathlib import Path
# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-15)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)

cang_zu_path = fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\仓租' + '\\(处理完成)所有-海外仓-仓租明细.xlsx'
cang_zu_df = pd.read_excel(cang_zu_path)

# 以站点商品ID识别码为键进行合并，选择左连接（left join），这样可以确保表1的所有数据都被保留
result_df = pd.merge(main_file_df, cang_zu_df[['站点商品ID识别码', '海外仓仓租费']], on='站点商品ID识别码', how='left')

# 找出 仓租表 中在 订单统计 中不存在的行
missing_rows = cang_zu_df[~cang_zu_df['站点商品ID识别码'].isin(main_file_df['站点商品ID识别码'])]
# 将这些缺失的行添加到结果中
result_df = pd.concat([result_df, missing_rows], ignore_index=True)
# 确保所有期望的列都存在
expected_columns = list(main_file_df.columns) + ['海外仓仓租费']
for col in expected_columns:
    if col not in result_df.columns:
        result_df[col] = None  # 如果列不存在，添加该列并填充为 None
# 对于表1中没有的行，将新增的列填充为0
result_df[['海外仓仓租费']] = result_df[['海外仓仓租费']].fillna(0)

# 重新排序，确保列的顺序符合要求
result_df = result_df[expected_columns]

# 获取列 '所有仓库-无平台-需要分摊的费用' 的第一行数据
all_fen_tan_cang_zu = cang_zu_df['所有仓库-无平台-需要分摊的费用'].iloc[0]
# 新建一列，并在第一个单元格写入数据
result_df['所有仓库-无平台-需要分摊的费用'] = None  # 先初始化一列，填入 None 或其他默认值
result_df.at[0, '所有仓库-无平台-需要分摊的费用'] = all_fen_tan_cang_zu  # 在第一行的 新建一列写入数据

# # -BC-ls、-BC-xj 分摊 LM-BC的仓租（先对半分，再按销量占比细分）
lm_bc_sum = cang_zu_df['LM-BC的仓租'].iloc[0]
# 筛选并提取后缀
df_filtered = result_df[
    result_df['站点'].str.contains('-BC-ls$|-BC-xj$', regex=True, na=False)
].copy()
df_filtered['suffix'] = df_filtered['站点'].str.extract(r'-(BC-ls|BC-xj)$', expand=False)
# 计算每个 suffix 的总销量（全局，不分 base_site）
df_filtered['total_sales_by_suffix'] = df_filtered.groupby('suffix')['销量'].transform('sum')
# 每个 suffix 分 lm_bc_sum 的一半
half_amount = lm_bc_sum / 2
# 计算海外仓仓租费（按 suffix 组内销量占比）
df_filtered['海外仓仓租费'] = half_amount * df_filtered['销量'] / df_filtered['total_sales_by_suffix']
df_filtered['海外仓仓租费'] = df_filtered['海外仓仓租费'].fillna(0)
# 更新回原表
result_df.loc[df_filtered.index, '海外仓仓租费'] = df_filtered['海外仓仓租费']


def replace_site_with_rent(df, ping_tai, site):
    """
    通用站点替换+仓租累加函数
    返回：处理后的新 DataFrame
    """
    site_end = site.split('-')[-1] if ping_tai != 'MANO-EU' else site
    # 1. 构造候选替换列表
    plat_df = df[df['平台'] == ping_tai]
    # 筛选站点包含“site”但不等于“site”的行
    filtered_df = plat_df[plat_df['站点'].str.contains(site_end, na=False) & (plat_df['站点'] != site)]
    # 按站点分组，计算销量总和，降序排序
    site_sales = (filtered_df.groupby('站点', as_index=False)['销量'].sum().sort_values(by='销量', ascending=False))
    # 得到动态替换顺序
    replacements = site_sales['站点'].tolist()
    print(f'M{site}的站点替换顺序，按站点销量大到小排序：\n{replacements}')
    # 2. 用上面得到的 {site}替换站点的list，去轮流替换站点为{site} 的站点
    target_rows = df[df['站点'] == site].copy()
    # 开始循环替换
    rows_to_drop = []  # 需要删除的新增行索引
    rows_to_append = []  # 若都匹配不到，最后保留的替换结果
    for _, row in target_rows.iterrows():
        matched = False
        for new_site in replacements:
            new_id = f"{new_site}{row['商品ID']}"
            # 先查是否已存在
            hit = df.index[df['站点商品ID识别码'] == new_id]
            if len(hit):
                # 累加海外仓仓租费
                df.loc[hit[0], '海外仓仓租费'] += row['海外仓仓租费']
                matched = True
                break
        if not matched:
            # 替换后的 站点商品ID识别码 都没匹配到，保留第一个替换结果
            last_site = replacements[0]
            last_id = f"{last_site}{row['商品ID']}"
            new_row = row.copy()
            new_row['站点'] = last_site
            new_row['站点商品ID识别码'] = last_id
            rows_to_append.append(new_row)
        # 论是否匹配，原 LM-BTH行最终都要删掉
        rows_to_drop.append(_)
    # 3. 删除原行
    df = df.drop(index=rows_to_drop)
    # 4 把没匹配到的行追加回去
    if rows_to_append:
        df = pd.concat([df, pd.DataFrame(rows_to_append)], ignore_index=True)
    return df


# TODO 替换站点，放入仓租
# MANO-FR，替换站点，放入仓租         ping_tai：平台       site：站点
result_df = replace_site_with_rent(result_df, ping_tai='MANO-EU', site='MANO-FR')
# REAL-FB，替换站点，放入仓租
result_df = replace_site_with_rent(result_df, ping_tai='REAL', site='REAL-FB')
# LM-BTH，替换站点，放入仓租
result_df = replace_site_with_rent(result_df, ping_tai='LM', site='LM-BTH')
# LM-BC-ls，替换站点，放入仓租
result_df = replace_site_with_rent(result_df, ping_tai='LM', site='LM-BC-ls')
# AMAZON-DE，替换站点，放入仓租
result_df = replace_site_with_rent(result_df, ping_tai='AMAZON-EU', site='AMAZON-DE')

# 保存 文件
output_path = main_file_path.replace('已完成-15', '已完成-16')
result_df.to_excel(output_path, index=False)
print(f"处理完成，结果已保存到{output_path}")
