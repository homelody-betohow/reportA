import json
import numpy as np
import pandas as pd
import importlib.util
import openpyxl
from openpyxl.styles import Alignment
from pathlib import Path

# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_PROJECT_ROOT = _epr_mod.bootstrap(__file__)

from common.style import Color
from common.platform_shop import map_site_vat_commission
from config.A0_set_date import shared_date, folder_name
from config.A0_paths import DESKTOP_ROOT

# 本机映射（取代桌面「castorama - SKU类目佣金比例.xlsx」）
CASTORAMA_COMMISSION_PATH = _PROJECT_ROOT / "runtime" / "local" / "castorama_commission.json"

_COL_SKU = "SKU"
_COL_RATE = "佣金比"
_COL_MAPPED_RATE = "映射佣金比"


def _normalize_sku(sku) -> str:
    """与 sku_mapping 一致：strip；剥尾缀 -NW。"""
    if sku is None or (isinstance(sku, float) and pd.isna(sku)):
        return ""
    s = str(sku).strip()
    if s.endswith("-NW"):
        s = s[:-3]
    return s


def _parse_rate(val) -> float | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def _dump_commission_json(payload: dict, json_path: Path) -> None:
    """写出 castorama 佣金 JSON：items 中每个 {} 占一行，便于对照编辑。"""
    items = payload.get("items") if isinstance(payload, dict) else None
    if not isinstance(items, list):
        json_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        return

    meta = {k: v for k, v in payload.items() if k != "items"}
    lines = ["{"]
    for key, val in meta.items():
        lines.append(
            f"  {json.dumps(key, ensure_ascii=False)}: "
            f"{json.dumps(val, ensure_ascii=False)},"
        )
    item_field_order = (_COL_SKU, _COL_RATE)
    lines.append('  "items": [')
    for i, row in enumerate(items):
        if isinstance(row, dict):
            ordered = {k: row.get(k) for k in item_field_order}
            for k, v in row.items():
                if k not in ordered:
                    ordered[k] = v
            row = ordered
        row_json = json.dumps(row, ensure_ascii=False, separators=(", ", ": "))
        suffix = "," if i < len(items) - 1 else ""
        lines.append(f"    {row_json}{suffix}")
    lines.append("  ]")
    lines.append("}")
    lines.append("")
    json_path.write_text("\n".join(lines), encoding="utf-8")


def _load_castorama_commission(json_path: Path) -> dict[str, float]:
    """
    读取 runtime/local/castorama_commission.json → {规范化SKU: 佣金比}。
    items 形如：[{"SKU": "...", "佣金比": 0.1}, ...]
    佣金比为 null / 非数字则跳过。
    """
    if not json_path.is_file():
        return {}
    try:
        payload = json.loads(json_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        print(f"{Color.YELLOW}[C5] 无法读取 castorama 佣金 JSON {json_path}：{exc}{Color.RESET}")
        return {}

    items = payload.get("items") if isinstance(payload, dict) else None
    if not isinstance(items, list):
        print(f"{Color.YELLOW}[C5] JSON 缺少 items 列表，已跳过：{json_path}{Color.RESET}")
        return {}

    out: dict[str, float] = {}
    for row in items:
        if not isinstance(row, dict):
            continue
        rate = _parse_rate(row.get(_COL_RATE))
        if rate is None:
            continue
        sku = _normalize_sku(row.get(_COL_SKU))
        if not sku:
            continue
        out[sku] = rate
    return out


def _read_commission_payload(json_path: Path) -> dict:
    """读取 castorama_commission.json；文件不存在或损坏时返回空骨架。"""
    default = {
        "version": 1,
        "description": (
            "Castorama SKU 类目佣金比例本机映射"
            "（取代桌面 castorama - SKU类目佣金比例.xlsx）。"
            "字段：SKU、佣金比。"
        ),
        "items": [],
    }
    if not json_path.is_file():
        return default
    try:
        payload = json.loads(json_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        print(f"{Color.YELLOW}[C5] 读取 {json_path} 失败，将重建：{exc}{Color.RESET}")
        return default
    if not isinstance(payload, dict):
        return default
    if not isinstance(payload.get("items"), list):
        payload["items"] = []
    payload.setdefault("version", 1)
    payload.setdefault("description", default["description"])
    return payload


def apply_castorama_commission_from_json(df: pd.DataFrame, json_path: Path) -> pd.DataFrame:
    """
    用本机 JSON 按 SKU 填充「映射佣金比」（取代 Excel sku_mappings）。
    匹配键：规范化 SKU（剥 -NW）。
    """
    out = df.copy()
    rate_map = _load_castorama_commission(json_path)
    keys = out[_COL_SKU].map(_normalize_sku)
    mapped = keys.map(rate_map)
    out[_COL_MAPPED_RATE] = pd.to_numeric(mapped, errors="coerce")

    hit = int(out[_COL_MAPPED_RATE].notna().sum())
    total = len(out)
    print(
        f"{Color.CYAN}[C5] castorama_commission.json 映射：{hit}/{total} 行命中「{_COL_MAPPED_RATE}」"
        f"\n  文件：{json_path}{Color.RESET}"
    )
    if not rate_map:
        print(
            f"{Color.YELLOW}[C5] JSON 为空或未启用；castorama 将依赖 SKU 第4-5位规则兜底{Color.RESET}"
        )
    return out


def _merge_missing_into_castorama_commission_json(df: pd.DataFrame, json_path: Path) -> int:
    """
    将 castorama 仍缺「映射平台费（佣金）」的 SKU 追加进 castorama_commission.json。
    - 已存在的 SKU：保留原佣金比；
    - 不存在的：追加一条，佣金比=null，待手工填写后重跑 C5。
    返回新追加条数。
    """
    castorama = df["平台"].astype(str).str.lower() == "castorama"
    empty = castorama & (
        df["映射平台费（佣金）"].isna()
        | df["映射平台费（佣金）"].astype(str).str.strip().isin(["", "nan", "None"])
    )
    miss_df = df.loc[empty]
    if miss_df.empty:
        return 0

    pending: dict[str, dict] = {}
    for _, r in miss_df.iterrows():
        sku = _normalize_sku(r.get(_COL_SKU))
        if not sku or sku in pending:
            continue
        pending[sku] = {_COL_SKU: sku, _COL_RATE: None}
    if not pending:
        return 0

    json_path.parent.mkdir(parents=True, exist_ok=True)
    payload = _read_commission_payload(json_path)
    existing_items: list[dict] = []
    existing_keys: set[str] = set()

    for row in payload["items"]:
        if not isinstance(row, dict):
            continue
        key = _normalize_sku(row.get(_COL_SKU))
        if not key:
            continue
        existing_keys.add(key)
        row[_COL_SKU] = key
        existing_items.append(row)

    n_added = 0
    for key, row in pending.items():
        if key in existing_keys:
            continue
        existing_items.append(row)
        n_added += 1

    existing_items.sort(key=lambda x: _normalize_sku(x.get(_COL_SKU)))
    payload["items"] = existing_items
    _dump_commission_json(payload, json_path)
    print(
        f"{Color.YELLOW}[C5] 已写入 {json_path}："
        f"新增待填 {n_added} 条，合计 {len(existing_items)} 条"
        f"（请填写「{_COL_RATE}」后重跑 C5）{Color.RESET}"
    )
    return n_added


# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-7)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)
# 保留「分销」列（C4 fillna(0) 可能将其变为 0）
if '分销' not in main_file_df.columns:
    main_file_df['分销'] = '否'
else:
    main_file_df['分销'] = main_file_df['分销'].replace({0: '否', '0': '否'})
    main_file_df['分销'] = main_file_df['分销'].fillna('否')
# 重命名
main_file_df = main_file_df.rename(columns={'映射VAT税': 'amazon-VAT税'})

# 映射 castorama 的 佣金比例（本机 JSON，取代桌面 xlsx）
main_file_df = apply_castorama_commission_from_json(main_file_df, CASTORAMA_COMMISSION_PATH)

# 映射 平台费（佣金）、VAT税（仅 DB platform_shop）
main_file_df_1 = map_site_vat_commission(
    main_df=main_file_df, site_col='站点', excel_fallback=False
)

# castorama：平台费不用 DB，一律用 SKU 类目「映射佣金比」
_castorama = main_file_df_1['平台'].astype(str).str.lower() == 'castorama'
main_file_df_1.loc[_castorama, '映射平台费（佣金）'] = main_file_df_1.loc[_castorama, '映射佣金比']

# 其他平台：用“映射佣金比”填补“映射平台费（佣金）”的空值
main_file_df_1['映射平台费（佣金）'] = main_file_df_1['映射平台费（佣金）'].fillna(main_file_df_1['映射佣金比'])

# castorama 仍为空：按 SKU 第4-5位兜底（01/02→0.1，03→0.12；非数字如 BF 不转换）
_sku_rule = main_file_df_1['SKU'].astype(str).str.slice(3, 5).map(
    {'01': 0.1, '02': 0.1, '03': 0.12}
)
_need_rule = _castorama & main_file_df_1['映射平台费（佣金）'].isna()
main_file_df_1.loc[_need_rule, '映射平台费（佣金）'] = _sku_rule.loc[_need_rule]
_rule_hit = int((_need_rule & main_file_df_1['映射平台费（佣金）'].notna()).sum())
if _rule_hit:
    print(f"{Color.CYAN}[C5] SKU 第4-5位规则兜底：补全 {_rule_hit} 行{Color.RESET}")

main_file_df_2 = main_file_df_1

"""
如果 平台 在["AMAZON-EU", "AMAZON-US", "DLZ-EU"]，则  平台费、VAT = 平台销售额 * 对应比例
(AMAZON的VAT、平台费先这样计算，后面没有VAT的会替换0；平台费会去 - 退款的佣金)
否则，平台费、VAT = 销售额 * 对应比例
"""
cond = (
    (main_file_df_2['平台'].isin(["AMAZON-EU", "AMAZON-US", "DLZ-EU"]))
)

# 1 计算平台费
main_file_df_2['平台费'] = np.where(
    cond,
    main_file_df_2['平台销售额'] * main_file_df_2['映射平台费（佣金）'],
    main_file_df_2['销售额'] * main_file_df_2['映射平台费（佣金）']
)
# 2 VAT
main_file_df_2['销售税-本土'] = np.where(
    cond,
    main_file_df_2['平台销售额'] * main_file_df_2['映射VAT税'],
    main_file_df_2['销售额'] * main_file_df_2['映射VAT税']
)

# 将“平台”中包含 "AMAZON" 的"销售税-本土"替换 0
main_file_df_2.loc[main_file_df_2["平台"].str.contains("AMAZON", na=False), "销售税-本土"] = 0
# 亚马逊的VAT = 平台销售额VAT-amazon - 销售退款金额VAT-amazon （本土平台对应位置为0）
main_file_df_2['销售税'] = main_file_df_2['销售税-本土'] + main_file_df_2['平台销售额VAT-amazon'] - main_file_df_2[
    '销售退款金额VAT-amazon']
# 亚马逊的平台费 = 平台销售额的平台费 - 销售退款金额的佣金
main_file_df_2['平台费'] = main_file_df_2['平台费'] - main_file_df_2['销售退款金额的佣金']
# 确保相关列的数据类型为数值类型
main_file_df_2['平台费'] = np.round(pd.to_numeric(main_file_df_2['平台费'], errors='coerce'), 2)
main_file_df_2['销售税'] = np.round(pd.to_numeric(main_file_df_2['销售税'], errors='coerce'), 2)

# 创建新列“平台费(AMZ)”和“销售税(AMZ)”，初始值为NaN
main_file_df_2['平台费(AMZ)'] = np.nan
main_file_df_2['销售税(AMZ)'] = np.nan

# 如果“平台”列等于“amazon”，则将“平台费”的值移动到“平台费(AMZ)”，并将“销售税”的值移动到“销售税(AMZ)”
mask = main_file_df_2['平台'].str.contains('AMAZON', case=False, na=False)
main_file_df_2.loc[mask, '平台费(AMZ)'] = main_file_df_2.loc[mask, '平台费']
main_file_df_2.loc[mask, '销售税(AMZ)'] = main_file_df_2.loc[mask, '销售税']
# 重命名列
main_file_df_2 = main_file_df_2.rename(columns={'平台费': '平台费(非AMZ)'})
main_file_df_2 = main_file_df_2.rename(columns={'销售税': '销售税(非AMZ)'})

# 将“平台”列等于“amazon”的“平台费(非AMZ)”和“销售税(非AMZ)”中对应的值设置为0
mask = main_file_df_2['平台'].str.contains('AMAZON', case=False, na=False)
main_file_df_2.loc[mask, '平台费(非AMZ)'] = 0
main_file_df_2.loc[mask, '销售税(非AMZ)'] = 0

# 将所有相关列的空值填充为0
main_file_df_2[['平台费(非AMZ)', '平台费(AMZ)', '销售税(非AMZ)', '销售税(AMZ)']] = main_file_df_2[
    ['平台费(非AMZ)', '平台费(AMZ)', '销售税(非AMZ)', '销售税(AMZ)']].fillna(0)

main_file_df_2['平台费合计'] = main_file_df_2['平台费(AMZ)'] + main_file_df_2['平台费(非AMZ)']
main_file_df_2['销售税合计'] = main_file_df_2['销售税(AMZ)'] + main_file_df_2['销售税(非AMZ)']

# 3 提现费
"""
如果 站点 == OTTO-BTH，则 提现费 = 销售额 * 0.03；
否则，提现费 = 平台销售额 * 0.01
"""
main_file_df_2['提现费'] = np.select(
    [
        main_file_df_2['站点'] == 'OTTO-BTH'
    ],
    [
        np.round(main_file_df_2['销售额'] * 0.03, 2)
    ],
    default=np.round(main_file_df_2['平台销售额'] * 0.01, 2)  # 其余所有站点
)

# 当销售额 = 0 时，把平台费、VAT、提现费强制设为 0
main_file_df_2.loc[
    main_file_df_2['销售额'] == 0, ['平台费(AMZ)', '平台费(非AMZ)', '平台费合计',
                                    '销售税(AMZ)', '销售税(非AMZ)', '销售税合计', '提现费']] = 0

# 按照 'SKU-站点识别码' 列进行分组，进行汇总
main_file_df_2 = main_file_df_2.groupby('SKU-站点识别码').agg({
    'SKU': 'first',  # 保留每组的第一行数据
    '站点': 'first',  # 保留每组的第一行数据
    '平台': 'first',  # 保留每组的第一行数据
    'SKU-平台识别码': 'first',  # 保留每组的第一行数据
    '分销': 'first',  # 保留每组的第一行数据
    '映射平台费（佣金）': 'first',  # 保留每组的第一行数据
    '映射佣金比': 'first',  # 保留每组的第一行数据
    '映射VAT税': 'first',  # 保留每组的第一行数据
    '平台销售额': 'sum',  # 汇总
    '头程': 'sum',  # 汇总
    '关税': 'sum',  # 汇总
    '派送费': 'sum',  # 汇总
    '销量': 'sum',  # 汇总
    '重发数量': 'sum',  # 汇总
    '订单采购成本': 'sum',  # 汇总
    '重发采购成本': 'sum',  # 汇总
    '退款额': 'sum',  # 汇总
    '退款数量': 'sum',  # 汇总
    '销售额': 'sum',  # 汇总
    '平台费(非AMZ)': 'sum',  # 汇总
    '销售税(非AMZ)': 'sum',  # 汇总
    '平台费(AMZ)': 'sum',  # 汇总
    '销售税(AMZ)': 'sum',  # 汇总
    '平台费合计': 'sum',  # 汇总
    '销售税合计': 'sum',  # 汇总
    '提现费': 'sum'  # 汇总
}).reset_index()

# 映射平台费（佣金） 为空 => 平台费(非AMZ)、平台费合计 置空
mask_null = main_file_df_2['映射平台费（佣金）'].isna()
main_file_df_2.loc[mask_null, ['平台费(非AMZ)', '平台费合计']] = np.nan


# 将「分销」列移到最后
_cols = [c for c in main_file_df_2.columns if c != '分销'] + ['分销']
main_file_df_2 = main_file_df_2[_cols]

# 保存结果
output_path = main_file_path.replace('已完成-7', '已完成-8')
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    main_file_df_2.to_excel(writer, index=False)
    worksheet = writer.sheets['Sheet1']
    fenxiao_col_idx = main_file_df_2.columns.get_loc('分销') + 1
    fenxiao_col_letter = openpyxl.utils.get_column_letter(fenxiao_col_idx)
    center_align = Alignment(horizontal='center')
    for cell in worksheet[fenxiao_col_letter]:
        cell.alignment = center_align
print(f'处理完成，output_path：{output_path}')

# 自动检查：平台 == castorama 的「映射平台费（佣金）」是否为空
_castorama = main_file_df_2['平台'].astype(str).str.lower() == 'castorama'
_empty_commission = _castorama & (
    main_file_df_2['映射平台费（佣金）'].isna()
    | main_file_df_2['映射平台费（佣金）'].astype(str).str.strip().isin(['', 'nan', 'None'])
)
if _empty_commission.any():
    _merge_missing_into_castorama_commission_json(main_file_df_2, CASTORAMA_COMMISSION_PATH)
    print(
        f'{Color.RED}检查失败：castorama 的「映射平台费（佣金）」存在空值，'
        f'共 {_empty_commission.sum()} 行；'
        f'请编辑 {CASTORAMA_COMMISSION_PATH} 填写「佣金比」后重跑本脚本{Color.RESET}'
    )
    print(main_file_df_2.loc[_empty_commission, ['SKU-站点识别码', 'SKU', '站点', '平台', '映射平台费（佣金）', '映射佣金比']].drop_duplicates().to_string(index=False))
    raise SystemExit(1)

    
print(f'{Color.GREEN}一切正常，请进行下一步操作{Color.RESET}')
