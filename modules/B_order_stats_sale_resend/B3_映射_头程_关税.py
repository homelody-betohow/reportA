import importlib.util
from pathlib import Path

# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

import sys

import pandas as pd
from common.style import Color
from common.sku_mapping import sku_mappings
from config.A0_set_date import shared_date, folder_name
from config.A0_paths import BTH_ALL_SKU_DETAIL_PATH, DESKTOP_ROOT

# 分销 SKU：头程、关税允许为 0
_FENXIAO_SKU_PATTERN = r'^(25|SN25|207)'

# 订单「站点」为该列表时，按非美国市场处理（即使映射站点名含 US，如 DLZ-US）
_NON_US_SITE_CODES = (
    'PT', 'DE', 'FR', 'IT', 'ES', 'UK', 'GB', 'NL', 'BE', 'PL', 'SE', 'AT', 'CH',
    'IE', 'DK', 'FI', 'NO', 'CZ', 'HU', 'RO', 'GR', 'AU', 'CA', 'MX', 'JP',
)
_NON_US_SITE_PATTERN = (
    r'(?:^|[-_/])(?:' + '|'.join(_NON_US_SITE_CODES) + r')(?:$|[-_/])'
)


def _as_bool_series(mask, index: pd.Index) -> pd.Series:
    """保证为与 index 对齐的一维 bool Series（避免重复列名导致 DataFrame）。"""
    if isinstance(mask, pd.DataFrame):
        mask = mask.all(axis=1)
    s = pd.Series(mask, index=index) if not isinstance(mask, pd.Series) else mask
    return s.reindex(index, fill_value=False).fillna(False).astype(bool)


def _is_us_market(df: pd.DataFrame) -> pd.Series:
    """
    是否为美国市场（头程走 BTH 的 US 列、关税为 0）。
    映射站点含 US 不等于美国站：DLZ-US 为店铺账号，实盘可能为 PT/EU。
    """
    mapped = df['映射站点'].astype(str).str.strip()
    is_us = mapped.str.contains('US', case=False, na=False)
    # 店铺账号，非美国市场
    is_us = is_us & ~mapped.str.fullmatch(r'DLZ-US', case=False, na=False)

    if '站点' in df.columns:
        site = df['站点'].astype(str).str.strip().str.upper()
        non_us_site = site.str.contains(_NON_US_SITE_PATTERN, regex=True, na=False)
        is_us = is_us & ~non_us_site

    if '仓库' in df.columns:
        wh = df['仓库'].astype(str)
        eu_wh = wh.str.contains(r'HY-DLZ-DE|德国', regex=True, na=False)
        is_us = is_us & ~eu_wh

    return _as_bool_series(is_us, df.index)


# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-2)订单统计-{shared_date}.xlsx"
main_file_df = pd.read_excel(main_file_path)

mask = _is_us_market(main_file_df)
# 拆分
us_df = main_file_df[mask].copy()  # 包含 US
not_us_df = main_file_df[~mask].copy()  # 不包含 US

product_map_sku_path = BTH_ALL_SKU_DETAIL_PATH
# 映射 头程  US
us_df_1 = sku_mappings(
    main_df=us_df,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="SKU",  # 用于定位 产品编码
    map_new_sku="头程（RMB）",  # 用于定位 头程（RMB）  的 US
    map_sku_sheet='基础数据维护',
    xuan_lie_2_ci='US'  # 2次选列，选到 US列 （列名：Unnamed: 14） 用于定位 头程（RMB）  的 US
)
us_df_1 = us_df_1.rename(columns={'映射Unnamed: 14': '单个-头程运费'})

# US 无关税
us_df_1['单个-头程税费'] = 0  # 映射增一列：头程税费，数据为 0

# 映射 头程  not_us
not_us_df_1 = sku_mappings(
    main_df=not_us_df,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="SKU",
    map_new_sku="头程（RMB）",  # 用于定位 头程（RMB）  的 EU/AU
    map_sku_sheet='基础数据维护'
)
not_us_df_1 = not_us_df_1.rename(columns={'映射头程（RMB）': '单个-头程运费'})

# 映射 关税  not_us
not_us_df_2 = sku_mappings(
    main_df=not_us_df_1,
    main_sku='SKU',
    map_sku_path=product_map_sku_path,
    map_old_sku="SKU",
    map_new_sku="关税（含税）",  # 用于定位 头程（RMB）  的 EU/AU
    map_sku_sheet='基础数据维护'
)
not_us_df_2 = not_us_df_2.rename(columns={'映射关税（含税）': '单个-头程税费'})

# 合并数据
main_df_1 = pd.concat([us_df_1, not_us_df_2]).reset_index(drop=True)

main_df_1 = main_df_1.rename(columns={'头程运费': '原-头程运费'})
main_df_1 = main_df_1.rename(columns={'头程税费': '原-头程税费'})


def _need_bth_cost_rows(df: pd.DataFrame) -> pd.Series:
    is_zhg = df['仓库'].astype(str).str.contains('ZHG', na=False)
    is_fenxiao = df['SKU'].astype(str).str.match(_FENXIAO_SKU_PATTERN, na=False)
    qty = pd.to_numeric(df['仓库SKU销量'], errors='coerce').fillna(0)
    return ~is_zhg & ~is_fenxiao & (qty > 0)


def _load_bth_sku_set(bth_path: str) -> set[str]:
    """BTH「基础数据维护」中已有 SKU（与 sku_mappings 一致：去空格、用于查表）。"""
    bth = pd.read_excel(bth_path, sheet_name='基础数据维护', usecols=['SKU'])
    return set(bth['SKU'].dropna().astype(str).str.strip())


def _sku_lookup_keys(series: pd.Series) -> pd.Series:
    """与 sku_mappings 一致：去空格，去掉尾缀 -NW 后作为 BTH 查表键。"""
    s = series.astype(str).str.strip()
    return s.str.replace(r'-NW$', '', regex=True)


def _print_bth_issue_block(
    title: str,
    reason: str,
    action: str,
    issue_mask,
    df: pd.DataFrame,
    extra_cols: list[str] | None = None,
) -> None:
    issue_mask = _as_bool_series(issue_mask, df.index)
    n = int(issue_mask.sum())
    if n == 0:
        return
    print(f'\n  {Color.RED}[!!] {title}：{n} 行{Color.RESET}')
    print(f'  {Color.YELLOW}原因：{reason}{Color.RESET}')
    print(f'  {Color.YELLOW}处理：{action}{Color.RESET}')
    show_cols = ['订单号', 'SKU', '映射站点', '仓库', '仓库SKU销量']
    if extra_cols:
        show_cols.extend(extra_cols)
    show_cols = [c for c in show_cols if c in df.columns]
    sample = df.loc[issue_mask, show_cols].drop_duplicates().head(20)
    for _, row in sample.iterrows():
        print('   ', ' | '.join(f'{c}={row[c]}' for c in show_cols))
    if n > 20:
        print(f'    ... 另有 {n - 20} 行未列出')
    bad_skus = df.loc[issue_mask, 'SKU'].dropna().astype(str).str.strip().unique().tolist()
    print(f'\n  {Color.CYAN}涉及 SKU（去重，共 {len(bad_skus)} 个）：{Color.RESET}')
    print('   ', ', '.join(bad_skus[:30]))
    if len(bad_skus) > 30:
        print(f'    ... 另有 {len(bad_skus) - 30} 个未列出')


def _abort_if_bth_sku_missing(df: pd.DataFrame, bth_path: str) -> None:
    """BTH 查不到 SKU 或费用列为空时，分原因打印说明并终止（非 US 头程逻辑错误）。"""
    is_us = _is_us_market(df)
    need_data = _need_bth_cost_rows(df)
    sku_str = df['SKU'].astype(str).str.strip()
    lookup_key = _sku_lookup_keys(df['SKU'])
    bth_skus = _load_bth_sku_set(bth_path)
    in_bth = lookup_key.isin(bth_skus) | sku_str.isin(bth_skus)

    checks = [
        (
            '单个-头程运费',
            '头程（RMB）',
            need_data,
            'US 行：填写「头程（RMB）」右侧 US 列；非 US 行：填写「头程（RMB）」列',
        ),
        (
            '单个-头程税费',
            '关税（含税）',
            need_data & ~is_us,
            '在「关税（含税）」列填写数值',
        ),
    ]

    any_fail = False
    print(f'\n{Color.RED}{"=" * 60}')
    print('BTH全部SKU明细 — 头程/关税映射检查失败')
    print(f'{"=" * 60}{Color.RESET}')
    print(f'  {Color.CYAN}数据文件：{bth_path}{Color.RESET}')
    print(f'  {Color.CYAN}工作表：基础数据维护（按 SKU 列匹配）{Color.RESET}')

    for col, bth_field, row_mask, fill_hint in checks:
        if col not in df.columns:
            continue
        row_mask = _as_bool_series(row_mask, df.index)
        in_bth_mask = _as_bool_series(in_bth, df.index)
        numeric = pd.to_numeric(df[col], errors='coerce')
        val_str = df[col].astype(str).str.strip()
        sku_echo = row_mask & df[col].notna() & val_str.eq(sku_str)
        missing = _as_bool_series(row_mask & (numeric.isna() | sku_echo), df.index)
        if not missing.any():
            continue
        any_fail = True

        not_in_bth = missing & ~in_bth_mask
        empty_in_bth = missing & in_bth_mask

        _print_bth_issue_block(
            title=f'{col} — SKU 不在 BTH 表',
            reason=(
                '订单 SKU 在「BTH全部SKU明细 → 基础数据维护」中不存在'
                '（查表时会去掉 -NW 后缀再匹配）。与 US/非 US 头程列无关。'
            ),
            action=f'在 BTH 表中新增该 SKU，并填写「{bth_field}」等字段后重跑。',
            issue_mask=not_in_bth,
            df=df,
            extra_cols=[col],
        )
        _print_bth_issue_block(
            title=f'{col} — SKU 已在 BTH 表，但费用为空',
            reason=f'SKU 已在 BTH 中，但「{bth_field}」对应列无有效数字。',
            action=fill_hint,
            issue_mask=empty_in_bth,
            df=df,
            extra_cols=[col],
        )

    if not any_fail:
        return

    print(f'\n{Color.YELLOW}请更新 BTH 表后重跑本脚本{Color.RESET}')
    print(f'{Color.YELLOW}仍无法确认请联系李杨维护 SKU 主数据{Color.RESET}')
    print(f'{Color.RED}{"=" * 60}{Color.RESET}\n')
    sys.exit(1)


_abort_if_bth_sku_missing(main_df_1, product_map_sku_path)

# 费用 = 单个费用 *  仓库SKU销量                                             RMB 转 EUR
_single_freight = pd.to_numeric(main_df_1['单个-头程运费'], errors='coerce')
_single_tax = pd.to_numeric(main_df_1['单个-头程税费'], errors='coerce')
main_df_1['映射-头程运费'] = _single_freight * main_df_1['仓库SKU销量'] / 7.3
main_df_1['映射-头程税费'] = _single_tax * main_df_1['仓库SKU销量'] / 7.3

# 映射增一列：头程运费， 数据为： 原-头程运费 不为 0 的 、原-头程运费 为 0 的 对应的 映射-头程运费
main_df_1['头程运费'] = main_df_1['原-头程运费'].where(main_df_1['原-头程运费'] != 0,
                                                                 main_df_1['映射-头程运费'])
# 同理，映射增一列：头程税费
main_df_1['头程税费'] = main_df_1['原-头程税费'].where(main_df_1['原-头程税费'] != 0,
                                                                 main_df_1['映射-头程税费'])
# 美国市场无关税：覆盖 ERP 原值（避免 原-头程税费≠0 时仍保留税费）
_us_mask = _is_us_market(main_df_1)
main_df_1.loc[_us_mask, ['单个-头程税费', '映射-头程税费', '头程税费']] = 0
# '仓库'包含 ZHG, 则：头程运费、头程税费 为 0
mask = main_df_1['仓库'].str.contains('ZHG', na=False)
main_df_1.loc[mask, ['头程运费', '头程税费']] = 0


def _check_first_leg_costs(df: pd.DataFrame) -> None:
    """检查头程运费、头程税费是否齐全；US 站点关税应为 0。"""
    is_us = _is_us_market(df)
    is_zhg = df['仓库'].astype(str).str.contains('ZHG', na=False)
    is_fenxiao = df['SKU'].astype(str).str.match(_FENXIAO_SKU_PATTERN, na=False)
    need_data = _need_bth_cost_rows(df)

    single_freight = pd.to_numeric(df['单个-头程运费'], errors='coerce')
    single_tax = pd.to_numeric(df['单个-头程税费'], errors='coerce')
    freight = pd.to_numeric(df['头程运费'], errors='coerce')
    tax = pd.to_numeric(df['头程税费'], errors='coerce')

    map_freight_fail = need_data & single_freight.isna()
    map_tax_fail = need_data & ~is_us & single_tax.isna()
    freight_missing = need_data & freight.isna()
    tax_missing = need_data & ~is_us & tax.isna()
    freight_zero = need_data & freight.notna() & freight.eq(0)
    tax_zero = need_data & ~is_us & tax.notna() & tax.eq(0)
    us_tax_not_zero = is_us & need_data & tax.fillna(0).ne(0)

    all_issue_masks = [
        map_freight_fail, map_tax_fail, freight_missing, tax_missing,
        freight_zero, tax_zero, us_tax_not_zero,
    ]
    any_issue = pd.Series(False, index=df.index)
    for m in all_issue_masks:
        any_issue = any_issue | m

    def _print_issues(title: str, issue_mask: pd.Series, cols: list) -> None:
        n = int(issue_mask.sum())
        if n == 0:
            print(f'  {Color.GREEN}[OK] {title}：无异常{Color.RESET}')
            return
        print(f'  {Color.RED}[!!] {title}：{n} 行{Color.RESET}')
        sample = df.loc[issue_mask, cols].drop_duplicates().head(15)
        for _, row in sample.iterrows():
            print('    ', ' | '.join(f'{c}={row[c]}' for c in cols))
        if n > 15:
            print(f'    ... 另有 {n - 15} 行未列出')

    print(f'\n{Color.CYAN}{"=" * 60}')
    print('头程运费 / 头程税费 数据检查')
    print(f'{"=" * 60}{Color.RESET}')
    print(f'  总行数: {len(df)} | US: {is_us.sum()} | 非US: {(~is_us).sum()}')
    print(f'  ZHG仓库(允许为0): {is_zhg.sum()} | 分销SKU(允许为0): {is_fenxiao.sum()}')
    print(f'  需有数据的行(有销量且非ZHG非分销): {need_data.sum()}')

    issue_cols = ['订单号', 'SKU', '映射站点', '仓库', '仓库SKU销量',
                  '单个-头程运费', '单个-头程税费', '头程运费', '头程税费']
    issue_cols = [c for c in issue_cols if c in df.columns]

    _print_issues('头程运费-映射后仍为空', map_freight_fail, issue_cols)
    _print_issues('头程税费-映射后仍为空(非US)', map_tax_fail, issue_cols)
    _print_issues('头程运费-结果为0(可能 BTH 头程为0或未维护)', freight_zero, issue_cols)
    _print_issues('头程税费-结果为0(非US, 可能 BTH 关税为0或未维护)', tax_zero, issue_cols)
    _print_issues('US站点-头程税费应为0但非0', us_tax_not_zero, issue_cols)

    unique_issue_rows = int(any_issue.sum())
    if unique_issue_rows == 0:
        print(f'\n{Color.GREEN}检查通过：头程运费、头程税费数据齐全（US关税为0符合预期）{Color.RESET}')
    else:
        print(f'\n{Color.RED}合计异常行（去重后）: {unique_issue_rows} 行{Color.RESET}')
        print(f'{Color.YELLOW}请先核对是否分销SKU（25/SN25/207开头，头程关税可为0）{Color.RESET}')
        print(f'{Color.YELLOW}若仍缺失：先确认 SKU 是否在 BTH全部SKU明细（基础数据维护）；再联系李杨{Color.RESET}')
        print(f'{Color.YELLOW}path：{product_map_sku_path}{Color.RESET}')
    print(f'{Color.CYAN}{"=" * 60}{Color.RESET}\n')


_check_first_leg_costs(main_df_1)

# 保存结果
output_path = main_file_path.replace('已完成-2', '已完成-3')
main_df_1.to_excel(output_path, index=False)
print(f'处理完成，output_path：{output_path}')
