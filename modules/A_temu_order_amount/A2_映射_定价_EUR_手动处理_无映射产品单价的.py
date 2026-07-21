"""
A2：TEMU 订单 —— 映射产品单价与运费回款（EUR）

流程概要
--------
1. 读取「订单统计」Excel，校验 A3 单元格币种为 EUR。
2. 筛选平台为 semitemu 的 TEMU 行，生成「产品单价-识别码」（参考号 + 原平台sku）。
3. 从数据库 temu_order_item 按参考号=order_no 查询，SKU 依次尝试：
   原平台sku、平台sku（含 -NW 变体、逗号分隔首段等）并折算 EUR：
   - 映射产品单价（EUR）← declared_price（缺省用 order_payment/quantity）
   - 映射运费回款（EUR）← shipping_income
4. 第一次映射仍为空的行，从「手动-二次映射.xlsx」→「TEMU定价」按「产品单价-识别码」二次补全
   （等价于 Excel：单价 =VLOOKUP(识别码,TEMU定价!$C:$E,3,FALSE)；运费 =VLOOKUP(...,$C:$D,2,FALSE)）。
5. 重发订单（订单类型为「重发订单」或 resend）：上述两列强制为 0。
6. 另存为「只有TEMU(已完成-1)订单统计-*.xlsx」，供 A3 计算订单总金额。

说明
----
- 重发订单的原平台 sku 多为「--」，无法从 temu_order_item 匹配；销售额按 0 处理。
- 币种折算口径与 reportPRA/scripts/dataImport/order_temu.py 一致（A0_set_date 汇率）。
- 「手动-二次映射」中 C=识别码、D=映射运费回款（EUR）、E=映射产品单价（EUR）。
"""

import importlib.util
import sys
from decimal import Decimal, InvalidOperation
from pathlib import Path

# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

import re
import warnings
from datetime import datetime

import pandas as pd
import pymysql.cursors

from common.style import Color
from config.A0_set_date import (
    shared_date,
    folder_name,
    RMB_di_EUR,
    USD_to_EUR,
    kc_to_EUR,
    zl_to_EUR,
    Ft_to_EUR,
    CAD_to_EUR,
    kr_to_EUR,
    Lei_to_EUR,
)
from config.A0_paths import DESKTOP_ROOT

_REPORT_PRA_ROOT = next(
    (p / "reportPRA" for p in Path(__file__).resolve().parents if (p / "reportPRA").is_dir()),
    None,
)
if _REPORT_PRA_ROOT and str(_REPORT_PRA_ROOT) not in sys.path:
    sys.path.append(str(_REPORT_PRA_ROOT))

from database.db_connection import DatabaseConfig, get_db_manager  # noqa: E402  # pyright: ignore[reportMissingImports]

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl.styles.stylesheet")

# sku_mappings 写入的列名（=「映射」+ 定价表列名）
COL_KEY = "产品单价-识别码"
COL_MAP_UNIT_PRICE = "映射产品单价（EUR）"
COL_MAP_SHIPPING = "映射运费回款（EUR）"
MANUAL_SHEET = "TEMU定价"
# 重发订单在订单统计中的两种写法（B7 合并脚本会把「重发订单」替换成 resend）
RESEND_ORDER_TYPES = ("重发订单", "resend")

TEMU_ORDER_ITEM_TABLE = "temu_order_item"
_KEY_CHUNK = 200
_QUANTIZE_EUR = Decimal("0.0001")

_FX_VAR_BY_CURRENCY: dict[str, float] = {
    "USD": USD_to_EUR,
    "CAD": CAD_to_EUR,
    "CZK": kc_to_EUR,
    "PLN": zl_to_EUR,
    "HUF": Ft_to_EUR,
    "RON": Lei_to_EUR,
    "SEK": kr_to_EUR,
}

_TEMU_SELECT_COLS: tuple[str, ...] = (
    "order_no",
    "sku_id",
    "currency",
    "declared_price",
    "order_payment",
    "shipping_income",
    "quantity",
)

# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\订单统计-{shared_date}.xlsx"
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


def _strip_nw_sku(sku: str) -> str:
    """去掉 -NW 尾缀（与 sku_mapping 一致）。"""
    s = str(sku or "").strip()
    return re.sub(r"-NW$", "", s) if s.endswith("-NW") else s


def _sku_variants(sku) -> list[str]:
    """生成可用于关联的 SKU 变体（原值、去 -NW、逗号分隔各段）。"""
    s = _norm_key(sku)
    if not s:
        return []
    variants: list[str] = []
    seen: set[str] = set()

    def _add(v: str) -> None:
        v = str(v or "").strip()
        if v and v not in seen:
            seen.add(v)
            variants.append(v)

    _add(s)
    _add(_strip_nw_sku(s))
    if "," in s:
        for part in s.split(","):
            p = part.strip()
            _add(p)
            _add(_strip_nw_sku(p))
    return variants


def _row_lookup_keys(row) -> list[tuple[str, str]]:
    """订单行 → 待查询的 (参考号, sku) 键列表（原平台sku 优先，再平台sku）。"""
    ref = _norm_key(row.get("参考号"))
    if not ref:
        return []
    keys: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for col in ("原平台sku", "平台sku"):
        if col not in row.index:
            continue
        for sku in _sku_variants(row.get(col)):
            key = (ref, sku)
            if key not in seen:
                seen.add(key)
                keys.append(key)
    return keys


def _register_temu_row(store: dict[tuple[str, str], dict], row: dict) -> None:
    """将库表行按 order_no + sku 各变体写入索引（精确键优先）。"""
    order_no = str(row.get("order_no") or "").strip()
    sku_id = str(row.get("sku_id") or "").strip()
    if not order_no or not sku_id:
        return
    store[(order_no, sku_id)] = row
    for sku in _sku_variants(sku_id):
        store.setdefault((order_no, sku), row)


def _resolve_temu_row(row, temu_map: dict[tuple[str, str], dict]) -> dict | None:
    """按参考号 + 原平台sku/平台sku 变体在索引中解析 temu 行。"""
    for key in _row_lookup_keys(row):
        hit = temu_map.get(key)
        if hit is not None:
            return hit
    return None


def _to_decimal(v) -> Decimal:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return Decimal("0")
    if isinstance(v, Decimal):
        return v
    try:
        return Decimal(str(v))
    except (InvalidOperation, ValueError, TypeError):
        return Decimal("0")


def _normalize_currency(currency) -> str:
    c = str(currency or "").strip().upper()
    aliases = {
        "€": "EUR",
        "RMB": "CNY",
        "KC": "CZK",
        "ZL": "PLN",
        "FT": "HUF",
        "LEI": "RON",
        "KR": "SEK",
    }
    return aliases.get(c, c)


def _fx_rate_to_eur(currency) -> Decimal | None:
    """1 单位付款币 = ? EUR（与 order_temu.py 口径一致）。"""
    c = _normalize_currency(currency)
    if not c:
        return None
    if c == "EUR":
        return Decimal("1")
    if c == "CNY":
        return Decimal("1") / Decimal(str(RMB_di_EUR))
    rate = _FX_VAR_BY_CURRENCY.get(c)
    if rate is None:
        return None
    return Decimal(str(rate))


def _amount_to_eur(amount, currency, *, quantize: Decimal = _QUANTIZE_EUR) -> float | None:
    rate = _fx_rate_to_eur(currency)
    if rate is None:
        return None
    val = (_to_decimal(amount) * rate).quantize(quantize)
    return float(val)


def _round_shipping_eur(val) -> float:
    return float(_to_decimal(val).quantize(_QUANTIZE_EUR))


def _temu_unit_price_eur(row: dict) -> float | None:
    """产品单价（EUR）：优先 declared_price，否则 order_payment / quantity；库中为 0 时返回 0.0。"""
    qty = int(row.get("quantity") or 1)
    if qty <= 0:
        qty = 1
    unit = _to_decimal(row.get("declared_price"))
    if unit == 0:
        unit = _to_decimal(row.get("order_payment")) / Decimal(qty)
    if unit == 0:
        return 0.0
    return _amount_to_eur(unit, row.get("currency"))


def _temu_shipping_eur(row: dict) -> float | None:
    """运费回款（EUR）；库中为 0 时返回 0.0。"""
    ship = _to_decimal(row.get("shipping_income"))
    if ship == 0:
        return 0.0
    return _amount_to_eur(ship, row.get("currency"), quantize=_QUANTIZE_EUR)


def _fetch_temu_order_items(keys: list[tuple[str, str]]) -> dict[tuple[str, str], dict]:
    """按 (order_no, sku_id) 批量查询 temu_order_item。"""
    unique = sorted({(a, b) for a, b in keys if a and b})
    if not unique:
        return {}

    cols = ", ".join(f"`{c}`" for c in _TEMU_SELECT_COLS)
    result: dict[tuple[str, str], dict] = {}
    db = get_db_manager(DatabaseConfig())
    conn = db.get_connection()
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        for i in range(0, len(unique), _KEY_CHUNK):
            part = unique[i : i + _KEY_CHUNK]
            placeholders = ",".join(["(%s,%s)"] * len(part))
            sql = (
                f"SELECT {cols} FROM `{TEMU_ORDER_ITEM_TABLE}` "
                f"WHERE (`order_no`, `sku_id`) IN ({placeholders})"
            )
            params: list[str] = []
            for order_no, sku_id in part:
                params.extend([order_no, sku_id])
            cur.execute(sql, params)
            for row in cur.fetchall():
                _register_temu_row(result, row)
    finally:
        cur.close()
        conn.close()
    return result


def _is_unmapped(series: pd.Series) -> pd.Series:
    """第一次映射失败：空 / NaN（0 视为已映射，如重发前的有效值）。"""
    if series.dtype == object:
        stripped = series.astype(str).str.strip()
        return series.isna() | stripped.isin(("", "nan", "None", "none"))
    return series.isna()


def _is_unmapped_val(val) -> bool:
    if pd.isna(val):
        return True
    s = str(val).strip()
    return s in ("", "nan", "None", "none")


def _is_resend(df: pd.DataFrame) -> pd.Series:
    if "订单类型" not in df.columns:
        return pd.Series(False, index=df.index)
    return df["订单类型"].isin(RESEND_ORDER_TYPES)


def _apply_temu_order_item_map(df: pd.DataFrame) -> tuple[int, int, int]:
    """
    从 temu_order_item 映射单价与运费（EUR）。
    关联键：参考号=order_no；SKU 尝试原平台sku、平台sku 及其变体。
    返回 (补全单价行数, 补全运费行数, 未知币种行数)。
    """
    if COL_MAP_UNIT_PRICE not in df.columns:
        insert_pos = df.columns.get_loc(COL_KEY) + 1
        df.insert(insert_pos, COL_MAP_UNIT_PRICE, None)
    if COL_MAP_SHIPPING not in df.columns:
        insert_pos = df.columns.get_loc(COL_MAP_UNIT_PRICE) + 1
        df.insert(insert_pos, COL_MAP_SHIPPING, None)

    resend_mask = _is_resend(df)
    need_mask = ~resend_mask
    if not need_mask.any():
        return 0, 0, 0

    lookup_keys: list[tuple[str, str]] = []
    for _, row in df.loc[need_mask].iterrows():
        lookup_keys.extend(_row_lookup_keys(row))

    try:
        temu_map = _fetch_temu_order_items(lookup_keys)
    except Exception as exc:
        print(
            f"{Color.RED}[错误] 查询 temu_order_item 失败：{exc}{Color.RESET}"
        )
        print(
            f"{Color.YELLOW}[提示] 将跳过数据库映射，仅尝试手动二次映射{Color.RESET}"
        )
        return 0, 0, 0

    if not temu_map:
        print(
            f"{Color.YELLOW}temu_order_item 未命中任何 (参考号, sku)，"
            f"共查询 {len(set(lookup_keys))} 个键{Color.RESET}"
        )
        return 0, 0, 0

    n_price = n_ship = n_bad_fx = 0
    unmatched_refs: list[str] = []
    for idx, row in df.loc[need_mask].iterrows():
        temu = _resolve_temu_row(row, temu_map)
        if not temu:
            ref = _norm_key(row.get("参考号"))
            if ref:
                unmatched_refs.append(ref)
            continue

        if _is_unmapped_val(df.at[idx, COL_MAP_UNIT_PRICE]):
            price_eur = _temu_unit_price_eur(temu)
            if price_eur is not None:
                df.at[idx, COL_MAP_UNIT_PRICE] = price_eur
                n_price += 1
            elif _fx_rate_to_eur(temu.get("currency")) is None:
                n_bad_fx += 1

        if _is_unmapped_val(df.at[idx, COL_MAP_SHIPPING]):
            ship_eur = _temu_shipping_eur(temu)
            # 0.0 也需写入表格（库中运费/回款为 0 时不能留空）
            if ship_eur is not None:
                df.at[idx, COL_MAP_SHIPPING] = ship_eur
                n_ship += 1
            elif _fx_rate_to_eur(temu.get("currency")) is None:
                n_bad_fx += 1

    unique_rows = len({(r.get("order_no"), r.get("sku_id")) for r in temu_map.values()})
    print(
        f"{Color.GREEN}temu_order_item 映射：单价 {n_price} 行，运费 {n_ship} 行"
        f"（命中 {unique_rows} 条库表记录）{Color.RESET}"
    )
    if unmatched_refs:
        sample = list(dict.fromkeys(unmatched_refs))[:10]
        print(
            f"{Color.YELLOW}仍有 {len(set(unmatched_refs))} 个参考号未匹配到 temu_order_item，"
            f"示例：{', '.join(sample)}{Color.RESET}"
        )
    if n_bad_fx:
        print(
            f"{Color.RED}有 {n_bad_fx} 行 temu_order_item 币种无法折算 EUR，"
            f"请检查 currency 字段或 A0_set_date 汇率配置{Color.RESET}"
        )
    return n_price, n_ship, n_bad_fx


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
            ship_map[key] = _round_shipping_eur(row[ship_col])
    return price_map, ship_map


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
        f"{Color.YELLOW}非重发缺单价：补全「{MANUAL_SHEET}」或检查 temu_order_item；"
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


def _save_temu_excel(df: pd.DataFrame, path: str) -> str:
    """
    保存 TEMU 子表。目标 xlsx 被 Excel/WPS 占用时给出明确提示，并尝试写入带时间戳的备用文件。
    """
    try:
        df.to_excel(path, index=False)
        return path
    except PermissionError:
        stamp = datetime.now().strftime("%H%M%S")
        alt_path = path.replace(".xlsx", f"_备用_{stamp}.xlsx")
        print(
            f"\n{Color.RED}[错误] 无法写入文件（Permission denied）：{path}{Color.RESET}"
        )
        print(
            f"{Color.YELLOW}常见原因：该 Excel 正在 Excel/WPS 中打开。"
            f"请先关闭后重新运行本脚本。{Color.RESET}"
        )
        try:
            df.to_excel(alt_path, index=False)
            print(
                f"{Color.GREEN}已改存备用文件（请关闭原文件后手动改名覆盖）：{alt_path}{Color.RESET}"
            )
            return alt_path
        except PermissionError:
            print(
                f"{Color.RED}备用文件也无法写入：{alt_path}"
                f"\n请关闭桌面「订单统计」目录下所有相关 xlsx 后重试。{Color.RESET}"
            )
            sys.exit(1)


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


# ====== 因订单统计从 付款时间维度调整到发货时间维度，需要过滤掉 付款时间小于 2026-06-01 的行 ======
# 仅保留 付款时间大于 2026-06-01 的行
# temu_df = temu_df[temu_df["付款时间"] >= "2026-06-01"]
# ====== 因订单统计从 付款时间维度调整到发货时间维度，需要过滤掉 付款时间小于 2026-06-01 的行 ======
# 过滤掉 “订单销售状态” 为 问题件 的行
temu_df = temu_df[temu_df["订单销售状态"] != "问题件"]
# ============================================================================================


# 在原平台sku 后插入「产品单价-识别码」= 参考号 + 原平台sku（与 A1 / 手动二次映射键一致）
new_column_data = temu_df["参考号"] + temu_df["原平台sku"].astype(str)
insert_position = temu_df.columns.get_loc("原平台sku") + 1
temu_df.insert(insert_position, COL_KEY, new_column_data)

# ---------- 3. 从 temu_order_item 映射产品单价、运费回款（EUR）----------
_apply_temu_order_item_map(temu_df)

# ---------- 4. 手动二次映射（TEMU定价：C→D 运费、C→E 单价）----------
n_manual_price, n_manual_ship, still_empty = _apply_manual_secondary_map(
    temu_df, manual_map_path
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
resend_count = _apply_resend_zero(temu_df)
if resend_count:
    print(f"{Color.YELLOW}已将 {resend_count} 条重发订单的「{COL_MAP_UNIT_PRICE}」「{COL_MAP_SHIPPING}」设为 0{Color.RESET}")

# ---------- 6. 另存 TEMU 子表 ----------
if COL_MAP_SHIPPING in temu_df.columns:
    mapped = temu_df[COL_MAP_SHIPPING].notna() & ~_is_unmapped(temu_df[COL_MAP_SHIPPING])
    if mapped.any():
        temu_df.loc[mapped, COL_MAP_SHIPPING] = temu_df.loc[mapped, COL_MAP_SHIPPING].map(
            _round_shipping_eur
        )
output_path = (
    main_file_path.rsplit("\\", 1)[0]
    + "\\只有TEMU(已完成-1)"
    + main_file_path.rsplit("\\", 1)[1]
)
output_path = _save_temu_excel(temu_df, output_path)
print(f"处理完成，文件另存为：{output_path}")
_spot_check_eur_mappings(temu_df)
