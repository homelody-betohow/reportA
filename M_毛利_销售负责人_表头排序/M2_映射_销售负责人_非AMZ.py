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
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-20)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)

# 拆分数据
# 需要单独处理的站点
stations = ['TEMU-AIH', 'TEMU-BV', 'TEMU-HM', 'TEMU-AL', 'LM-TOTO', 'LM-ES-BTH', 'LM-FR-BTH', 'LM-IT-BTH', 'LM-PL-BTH',
            'LM-PT-BTH', 'TEMU-KR-A', 'TEMU-KR-B', 'TEMU-KR-C', 'TEMU-HJ-A', 'TEMU-HJ-B', 'TEMU-HJ-C', 'TEMU-NF-A',
            'TEMU-NF-B', 'TEMU-NF-C', 'LM-FR-BC-ls', 'LM-FR-BC-xj', 'LM-ES-BC-ls', 'LM-ES-BC-xj', 'LM-PT-BC-ls',
            'LM-PT-BC-xj', 'LM-IT-BC-ls', 'LM-IT-BC-xj', 'TEMU-BZ', 'TEMU-AQ', 'LM-FR-RP-ls', 'LM-FR-RP-xj',
            'LM-ES-RP-ls', 'LM-ES-RP-xj', 'LM-PT-RP-ls', 'LM-PT-RP-xj', 'LM-IT-RP-ls', 'LM-IT-RP-xj']
# 拆分：sp_stations 部分
st_df = main_file_df[main_file_df['站点'].isin(stations)]  # 站点 包含
no_st_df = main_file_df[~main_file_df['站点'].isin(stations)]  # 站点 不包含

product_map_sku_path = fr"{DESKTOP_ROOT}\信息-映射.xlsx"  # 改成对应的映射表
# 映射 销售负责人  通过 站点 映射
st_df_1 = sku_mappings(
    main_df=st_df,
    main_sku='站点',
    map_sku_path=product_map_sku_path,
    map_old_sku="站点",
    map_new_sku="销售负责人-站点",
    map_sku_sheet='销售负责人'
)
st_df_1 = st_df_1.rename(columns={'映射销售负责人-站点': '销售负责人'})

# 映射 销售负责人  通过 平台 映射
no_st_df_1 = sku_mappings(
    main_df=no_st_df,
    main_sku='平台',
    map_sku_path=product_map_sku_path,
    map_old_sku="平台",
    map_new_sku="销售负责人-平台",
    map_sku_sheet='销售负责人'
)
no_st_df_1 = no_st_df_1.rename(columns={'映射销售负责人-平台': '销售负责人'})
# 合并数据
main_file_df_1 = pd.concat([st_df_1, no_st_df_1]).reset_index(drop=True)

# 平台 == MANO-EU 且 站点 包含 BTH 对应的 销售负责人 替换成 无负责人
main_file_df_1.loc[(main_file_df_1['平台'] == 'MANO-EU') & (
    main_file_df_1['站点'].str.contains('BTH', na=False)), '销售负责人'] = '无负责人'

# 保存结果到新的 Excel 文件
output_path = main_file_path.replace('已完成-20', '已完成-21')
main_file_df_1.to_excel(output_path, index=False)
print(f"结果已保存到 {output_path}")
