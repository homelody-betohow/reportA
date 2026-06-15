"""
A2：TEMU 订单 —— 映射产品单价与运费回款（EUR）

流程概要
--------
1. 读取「订单统计」Excel，校验 A3 单元格币种为 EUR。
2. 筛选平台为 semitemu 的 TEMU 行，生成「产品单价-识别码」（参考号 + 原平台sku）。
3. 用 A1 产出的「TEMU-产品单价」表做 VLOOKUP 式映射，得到：
   - 映射产品单价（EUR）
   - 映射运费回款（EUR）
4. 第一次映射仍为空的行，从「手动-二次映射.xlsx」→「TEMU定价」按「产品单价-识别码」二次补全
   （等价于 Excel：单价 =VLOOKUP(识别码,TEMU定价!$C:$E,3,FALSE)；运费 =VLOOKUP(...,$C:$D,2,FALSE)）。
5. 重发订单（订单类型为「重发订单」或 resend）：上述两列强制为 0。
6. 另存为「只有TEMU(已完成-1)订单统计-*.xlsx」，供 A3 计算订单总金额。

说明
----
- 重发订单的原平台 sku 多为「--」，无法从产品单价表匹配；销售额按 0 处理。
- 「手动-二次映射」中 C=识别码、D=映射运费回款（EUR）、E=映射产品单价（EUR）。
"""

import importlib.util
from pathlib import Path

# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

import re
import sys
import warnings
import pandas as pd
from A_报表.Z_method.style import Color
from A_报表.Z_method.sku_映射 import sku_mappings
from A_报表.A0_设置_时间段.A0_set_date import shared_date, folder_name
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")

# sku_mappings 写入的列名（=「映射」+ 定价表列名）
COL_KEY = "产品单价-识别码"
COL_MAP_UNIT_PRICE = "映射产品单价（EUR）"
COL_MAP_SHIPPING = "映射运费回款（EUR）"
MANUAL_SHEET = "TEMU定价"
# 重发订单在订单统计中的两种写法（B7 合并脚本会把「重发订单」替换成 resend）
RESEND_ORDER_TYPES = ("重发订单", "resend")

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\订单统计-{shared_date}.xlsx"
product_map_sku_path = (
    fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\TEMU-产品单价\(处理完成)TEMU-产品单价.xlsx"
)
manual_map_path = fr"{DESKTOP_ROOT}\手动-二次映射.xlsx"


def _parse_currency_code(cell) -> str:
    """
    从 A3 类单元格解析 ISO 币种代码（大写）。
    支持：EUR、币种:EUR、Currency: eur、币种： EUR 等常见写法。
    """
    if pd.isna(cell):
        return ""
    s = str(cell).strip()
    if not s:
        return ""
    compact = re.sub(r"\s+", "", s.upper())
    if re.fullmatch(r"[A-Z]{3}", compact):
        return compact
    for sep in (":", "："):
        if sep not in s:
            continue
        tail = re.sub(r"\s+", "", s.split(sep)[-1].strip().upper())
        m = re.match(r"^([A-Z]{3})\b", tail)
        if m:
            return m.group(1)
    return ""


def _apply_resend_zero(df: pd.DataFrame) -> int:
    """重发订单：映射产品单价、运费回款均为 0。返回处理的行数。"""
    if "订单类型" not in df.columns:
        return 0
    mask = df["订单类型"].isin(RESEND_ORDER_TYPES)
    n = int(mask.sum())
    if n:
        df.loc[mask, [COL_MAP_UNIT_PRICE, COL_MAP_SHIPPING]] = 0
    return n


def _find_col(columns, *substrings: str) -> str:
    """按列名子串匹配，返回第一个命中的列名。"""
    for sub in substrings:
        for col in columns:
            if sub in str(col):
                return col
    raise KeyError(f"未找到包含 {substrings} 的列，现有列：{list(columns)}")


def _norm_key(val) -> str | None:
    if pd.isna(val):
        return None
    s = str(val).strip()
    return s if s and s.lower() not in ("nan", "none") else None


def _is_unmapped(series: pd.Series) -> pd.Series:
    """第一次映射失败：空 / NaN（0 视为已映射，如重发前的有效值）。"""
    if series.dtype == object:
        stripped = series.astype(str).str.strip()
        return series.isna() | stripped.isin(("", "nan", "None", "none"))
    return series.isna()


def _load_manual_temu_maps(path: str) -> tuple[dict[str, float], dict[str, float]]:
    """
    读取「手动-二次映射.xlsx」→「TEMU定价」。
    C 列识别码 → D 映射运费回款（EUR）、E 映射产品单价（EUR）（与 Excel VLOOKUP 一致）。
    重复键保留最后一行（与 sku_mappings 一致）。
    """
    manual_df = pd.read_excel(path, sheet_name=MANUAL_SHEET)
    manual_df = manual_df.map(lambda x: x.strip() if isinstance(x, str) else x)

    key_col = _find_col(manual_df.columns, COL_KEY)
    ship_col = _find_col(manual_df.columns, COL_MAP_SHIPPING, "映射运费回款")
    price_col = _find_col(manual_df.columns, COL_MAP_UNIT_PRICE, "映射产品单价")

    price_map: dict[str, float] = {}
    ship_map: dict[str, float] = {}
    for _, row in manual_df.iterrows():
        key = _norm_key(row[key_col])
        if not key:
            continue
        if pd.notna(row[price_col]):
            price_map[key] = float(row[price_col])
        if pd.notna(row[ship_col]):
            ship_map[key] = float(row[ship_col])
    return price_map, ship_map


def _is_resend(df: pd.DataFrame) -> pd.Series:
    if "订单类型" not in df.columns:
        return pd.Series(False, index=df.index)
    return df["订单类型"].isin(RESEND_ORDER_TYPES)


def _not_zero(series: pd.Series) -> pd.Series:
    """已填且数值不为 0（含无法转为数字的非空字符串）。"""
    num = pd.to_numeric(series, errors="coerce")
    return series.notna() & ~_is_unmapped(series) & num.ne(0)


def _spot_check_eur_mappings(df: pd.DataFrame) -> bool:
    """
    抽查「映射产品单价（EUR）」「映射运费回款（EUR）」：
    - 非重发：两列均不得为空；
    - 重发：两列均应为 0。
    返回 True 表示通过。
    """
    resend = _is_resend(df)
    non_resend = ~resend

    checks: list[tuple[str, pd.Series, list[str]]] = [
        (
            f"非重发 —「{COL_MAP_UNIT_PRICE}」为空",
            non_resend & _is_unmapped(df[COL_MAP_UNIT_PRICE]),
            [COL_KEY, "参考号", "原平台sku", "订单号", "订单类型", COL_MAP_UNIT_PRICE],
        ),
        (
            f"非重发 —「{COL_MAP_SHIPPING}」为空",
            non_resend & _is_unmapped(df[COL_MAP_SHIPPING]),
            [COL_KEY, "参考号", "原平台sku", "订单号", "订单类型", COL_MAP_SHIPPING],
        ),
        (
            f"重发订单 —「{COL_MAP_UNIT_PRICE}」应为 0（当前为空或非 0）",
            resend & (_is_unmapped(df[COL_MAP_UNIT_PRICE]) | _not_zero(df[COL_MAP_UNIT_PRICE])),
            [COL_KEY, "参考号", "订单号", "订单类型", COL_MAP_UNIT_PRICE, COL_MAP_SHIPPING],
        ),
        (
            f"重发订单 —「{COL_MAP_SHIPPING}」应为 0（当前为空或非 0）",
            resend & (_is_unmapped(df[COL_MAP_SHIPPING]) | _not_zero(df[COL_MAP_SHIPPING])),
            [COL_KEY, "参考号", "订单号", "订单类型", COL_MAP_UNIT_PRICE, COL_MAP_SHIPPING],
        ),
    ]

    any_issue = pd.Series(False, index=df.index)

    def _print_block(title: str, mask: pd.Series, cols: list[str]) -> None:
        nonlocal any_issue
        n = int(mask.sum())
        any_issue = any_issue | mask
        if n == 0:
            # print(f"  {Color.GREEN}[OK] {title}{Color.RESET}")
            return
        print(f"  {Color.RED}[!!] {title}：{n} 行{Color.RESET}")
        show_cols = [c for c in cols if c in df.columns]
        sample = df.loc[mask, show_cols].drop_duplicates().head(10)
        for _, row in sample.iterrows():
            print("    ", " | ".join(f"{c}={row[c]}" for c in show_cols))
        if n > 10:
            print(f"    ... 另有 {n - 10} 行未列出")

    print(f"\n{Color.CYAN}{'=' * 60}")
    print("抽查：映射产品单价 / 映射运费回款（EUR）")
    print(f"{'=' * 60}{Color.RESET}")
    print(f"  TEMU 行数: {len(df)} | 重发: {int(resend.sum())} | 非重发: {int(non_resend.sum())}")

    for title, mask, cols in checks:
        _print_block(title, mask, cols)

    issue_rows = int(any_issue.sum())
    if issue_rows == 0:
        print(
            f"\n{Color.GREEN}抽查通过：「{COL_MAP_UNIT_PRICE}」「{COL_MAP_SHIPPING}」"
            f"均已填妥（重发订单为 0）{Color.RESET}"
        )
        print(
            f"{Color.GREEN}可进行下一步：运行 A3_计算_TEMU_订单总金额.py{Color.RESET}"
        )
        print(f"{Color.CYAN}{'=' * 60}{Color.RESET}\n")
        return True

    print(f"\n{Color.RED}合计异常行（去重后）: {issue_rows} 行，请先处理后再执行 A3{Color.RESET}")
    print(
        f"{Color.YELLOW}非重发缺单价：补全「{MANUAL_SHEET}」或 A1 产品单价表；"
        f"重发订单应全部为 0{Color.RESET}"
    )
    print(f"{Color.CYAN}{'=' * 60}{Color.RESET}\n")
    return False


def _apply_manual_secondary_map(
    df: pd.DataFrame, path: str
) -> tuple[int, int, int]:
    """
    仅对第一次映射仍为空的行，用手动表补全单价与运费。
    返回 (补全单价行数, 补全运费行数, 仍缺单价行数)。
    """
    if not Path(path).is_file():
        print(
            f"{Color.YELLOW}未找到手动二次映射文件，跳过：{path}{Color.RESET}"
        )
        return 0, 0, int(_is_unmapped(df[COL_MAP_UNIT_PRICE]).sum())

    price_map, ship_map = _load_manual_temu_maps(path)
    resend_mask = (
        df["订单类型"].isin(RESEND_ORDER_TYPES)
        if "订单类型" in df.columns
        else pd.Series(False, index=df.index)
    )

    keys = df[COL_KEY].map(_norm_key)
    price_empty = _is_unmapped(df[COL_MAP_UNIT_PRICE]) & ~resend_mask
    ship_empty = _is_unmapped(df[COL_MAP_SHIPPING]) & ~resend_mask

    n_price = n_ship = 0
    if price_empty.any():
        filled = keys.map(price_map)
        hit = price_empty & filled.notna()
        n_price = int(hit.sum())
        if n_price:
            df.loc[hit, COL_MAP_UNIT_PRICE] = filled[hit].astype(float)

    if ship_empty.any():
        filled = keys.map(ship_map)
        hit = ship_empty & filled.notna()
        n_ship = int(hit.sum())
        if n_ship:
            df.loc[hit, COL_MAP_SHIPPING] = filled[hit].astype(float)

    still_empty = int(
        (_is_unmapped(df[COL_MAP_UNIT_PRICE]) & ~resend_mask).sum()
    )
    return n_price, n_ship, still_empty


# ---------- 1. 校验币种（A3 须为 EUR）----------
_a3_df = pd.read_excel(main_file_path, header=None, usecols=[0], skiprows=2, nrows=1)
_a3_raw = _a3_df.iloc[0, 0]
_currency_code = _parse_currency_code(_a3_raw)
print(
    f"{Color.YELLOW}表格中币种信息：原始值={_a3_raw!r}，解析币种代码={_currency_code!r}{Color.RESET}"
)
if _currency_code != "EUR":
    print(f"{Color.RED} --- ================== 币种非 EUR，脚本停止执行。 ================== {Color.RESET}")
    sys.exit(1)

# ---------- 2. 读取订单统计，只保留 TEMU ----------
main_file_df = pd.read_excel(main_file_path, skiprows=4)  # 跳过表头前 4 行
main_file_df = main_file_df.map(lambda x: x.strip() if isinstance(x, str) else x)

temu_df = main_file_df[main_file_df["平台"] == "semitemu"].copy()

# 在原平台sku 后插入「产品单价-识别码」= 参考号 + 原平台sku（与 A1 / 手动二次映射键一致）
new_column_data = temu_df["参考号"] + temu_df["原平台sku"].astype(str)
insert_position = temu_df.columns.get_loc("原平台sku") + 1
temu_df.insert(insert_position, COL_KEY, new_column_data)

# ---------- 3. 映射产品单价、运费回款 ----------
temu_df_1 = sku_mappings(
    main_df=temu_df,
    main_sku=COL_KEY,
    map_sku_path=product_map_sku_path,
    map_old_sku=COL_KEY,
    map_new_sku="产品单价（EUR）",
    map_sku_sheet="Sheet1",
)
temu_df_2 = sku_mappings(
    main_df=temu_df_1,
    main_sku=COL_KEY,
    map_sku_path=product_map_sku_path,
    map_old_sku=COL_KEY,
    map_new_sku="运费回款（EUR）",
    map_sku_sheet="Sheet1",
)

# ---------- 4. 手动二次映射（TEMU定价：C→D 运费、C→E 单价）----------
n_manual_price, n_manual_ship, still_empty = _apply_manual_secondary_map(
    temu_df_2, manual_map_path
)
if n_manual_price or n_manual_ship:
    print(
        f"手动二次映射：补全单价 {n_manual_price} 行，补全运费 {n_manual_ship} 行（来源：{manual_map_path}）"
    )
if still_empty:
    print(
        f"{Color.RED}仍有 {still_empty} 行「{COL_MAP_UNIT_PRICE}」为空"
        f"（非重发），请检查「{MANUAL_SHEET}」或识别码是否一致{Color.RESET}"
    )

# ---------- 5. 重发订单两列 EUR 置 0 ----------
resend_count = _apply_resend_zero(temu_df_2)
if resend_count:
    print(f"{Color.YELLOW}已将 {resend_count} 条重发订单的「{COL_MAP_UNIT_PRICE}」「{COL_MAP_SHIPPING}」设为 0{Color.RESET}")

# ---------- 6. 另存 TEMU 子表 ----------
output_path = (
    main_file_path.rsplit("\\", 1)[0]
    + "\\只有TEMU(已完成-1)"
    + main_file_path.rsplit("\\", 1)[1]
)
temu_df_2.to_excel(output_path, index=False)
print(f"处理完成，文件另存为：{output_path}")
_spot_check_eur_mappings(temu_df_2)
