import importlib.util
from pathlib import Path

# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

import pandas as pd
from common.transaction_excel import (
    extract_platform_sku,
    filter_fba_fee_rows,
    read_transaction_excel,
    resolve_transaction_source_file,
)
from config.A0_set_date import shared_date, folder_name, transaction_date
from config.A0_paths import DESKTOP_ROOT

print(f"{transaction_date}")

tx_dir = Path(DESKTOP_ROOT) / f"{folder_name}{shared_date}" / "transaction交易明细"
transaction_path_1 = resolve_transaction_source_file(
    tx_dir, kind="released", transaction_date=transaction_date
)
transaction_path_2 = resolve_transaction_source_file(
    tx_dir, kind="deferred", transaction_date=transaction_date
)

transaction_df_1 = read_transaction_excel(transaction_path_1)
transaction_df_2 = read_transaction_excel(transaction_path_2)

# 合并（纵向拼接）
merged_df = pd.concat([transaction_df_1, transaction_df_2], ignore_index=True)

# 去除整张表的前后空格
for col in merged_df.columns:
    merged_df[col] = merged_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

# 筛选有效订单行（新格式含费用类型；仅保留 payment_order / refund_order）
merged_df = filter_fba_fee_rows(merged_df)

# 应用提取规则，清洗仓库 sku
merged_df["seller sku"] = merged_df["seller sku"].apply(extract_platform_sku)
merged_df_1 = merged_df.rename(columns={"seller sku": "SKU"})

# 在 order id 后插入新列 order-id识别码
new_column_name = "order-id识别码"
new_column_data = merged_df_1["order id"].astype(str) + merged_df_1["SKU"].astype(str)
target_column = "order id"
insert_position = merged_df_1.columns.get_loc(target_column) + 1
merged_df_1.insert(insert_position, new_column_name, new_column_data)

# fba fees 的正数变负数，负数变正数
merged_df_2 = merged_df_1.copy()
merged_df_2["fba fees"] = merged_df_2["fba fees"].apply(lambda x: -x)

# 保留指定列
merged_df_2 = merged_df_2[["order id", "SKU", "order-id识别码", "fba fees"]]
# 按照 order-id识别码 列进行分组汇总
merged_df_3 = merged_df_2.groupby("order-id识别码").agg({
    "order id": "first",
    "SKU": "first",
    "fba fees": "sum",
}).reset_index()

output_path = tx_dir / f"(处理完成)transaction交易明细_已发放-推迟订单{transaction_date}.xlsx"
merged_df_3.to_excel(output_path, index=False)
print(f"处理完成，output_path：{output_path}")
