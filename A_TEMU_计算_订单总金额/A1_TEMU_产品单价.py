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

from A_报表.A0_设置_时间段.A0_set_date import *
from A_报表.A0_设置_时间段.A0_paths import DESKTOP_ROOT

warnings.simplefilter(action='ignore', category=FutureWarning)


def my_process(price_df, sheet):
    """
    清洗价格数据：统一日期、过滤取消订单、按店铺币种折算为 EUR。
    返回新增『产品单价（EUR）』『运费回款（EUR）』『店铺』『产品单价-识别码』列的 DataFrame。
    """
    # 空值的地方——补 0
    price_df = price_df.fillna(0)
    # 1. 统一日期精度  去掉时分秒
    price_df['日期'] = pd.to_datetime(price_df['日期']).dt.floor('d')
    # 2. 按测试区间过滤
    start_date = pd.to_datetime(test_start_date)
    end_date = pd.to_datetime(test_end_date)
    price_df = price_df[(price_df['日期'] >= start_date) & (price_df['日期'] <= end_date)]
    # 3. 剔除取消订单
    price_df = price_df[price_df['产品单价'] != '客户取消订单'].copy()
    # 4. 店铺分组：RMB→EUR  or  USD→EUR
    rmb_sheets = ['AIHOMEU', 'BathVogue_EU']
    usd_sheets = ['HAUSE_MATE', 'KR-A', 'KR-B', 'KR-C', 'HJ-A', 'HJ-B', 'HJ-C', 'NF-A', 'NF-B', 'NF-C']
    zi_niao_sheets = ['TEMU-AL', 'TEMU-BZ', 'TEMU-AQ']  # 多个不同币种
    site = ''
    if sheet in rmb_sheets:
        # RMB 定价，除以 7.3 转 EUR
        price_df['产品单价（EUR）'] = price_df['产品单价'].astype(float) / 7.3
        price_df['运费回款（EUR）'] = price_df['运费回款'].astype(float) / 7.3
        site = sheet
        # 构建识别码 = 参考号 + SKU ID
        price_df['产品单价-识别码'] = price_df["参考号"] + price_df['SKU ID'].astype('Int64').astype(str)
    elif sheet in usd_sheets:
        # USD 定价，乘以汇率转 EUR
        price_df['产品单价（EUR）'] = price_df['产品单价'].astype(float) * USD_to_EUR
        price_df['运费回款（EUR）'] = price_df['运费回款'].astype(float) * USD_to_EUR
        site = sheet
        # 构建识别码 = 参考号 + SKU ID
        price_df['产品单价-识别码'] = price_df["参考号"] + price_df['SKU ID'].astype('Int64').astype(str)
    elif sheet in zi_niao_sheets:
        # --------- 工具函数 ---------
        def to_eur(s):
            rate = {'€': 1, 'zł': zl_to_EUR, 'Ft': Ft_to_EUR, 'Kč': kc_to_EUR, 'kr': kr_to_EUR, 'Lei': Lei_to_EUR}
            # 1. 去空白
            no_space = s.astype(str).str.replace(r'\s+', '', regex=True)
            # 2. 给纯数字补 €
            no_space = no_space.str.replace(r'^(\d+(?:[.,]\d+)?)$', r'\1€', regex=True)
            # 3. 提取数字和币种
            pat = r'(\d+(?:[.,]\d+)?).*?(€|zł|Ft|Kč|kr|Lei)'
            tmp = no_space.str.extract(pat, expand=True)
            num, curr = tmp[0], tmp[1]
            # 4. 数字统一成点号并转 float
            num = pd.to_numeric(num.str.replace(',', '.'), errors='coerce')
            # 5. 转换
            return num * curr.map(rate)

        price_df['产品单价（EUR）'] = to_eur(price_df['产品单价'])
        price_df['运费回款'] = to_eur(price_df['运费回款'])  # 紫鸟店铺的“运费收入”，统一叫：运费回款
        price_df['税金收入'] = to_eur(price_df['税金收入'])
        price_df['运费税收入'] = to_eur(price_df['运费税收入'])
        # 总的收益  要加上 '预估扣除金额' 的 值（处理后的金额，是正数！）
        price_df['预估扣除金额'] = to_eur(price_df['预估扣除金额'])
        # 运费回款特殊计算   运费回款（EUR） = 运费回款 + 税金收入
        price_df['运费回款（EUR）'] = price_df['运费回款'].astype(float) + price_df['税金收入'].astype(float) + price_df[
            '运费税收入'].astype(float) + price_df['预估扣除金额'].astype(float)
        site = sheet
        # 构建识别码 = 参考号 + SKU 货号
        price_df['产品单价-识别码'] = price_df["参考号"] + price_df['SKU 货号'].astype(str)
    else:
        print(f'无法获取到对应的站点，请检查 sheet 名字，程序终止！！！')
        print(f'错误的 sheet：{sheet}')
        exit()

    # 5. 标记店铺
    price_df['店铺'] = site
    return price_df


# TODO 文件夹路径！！！
file_path = r"\\Betohow\数据报表\RPA\Temu\订单详情\TEMU-订单详情.xlsx"

# 先读+清洗，再丢掉每张表里全空的列，最后 concat
processed = []
with pd.ExcelFile(file_path) as xls:
    for name in xls.sheet_names:
        df = my_process(pd.read_excel(xls, sheet_name=name), name)
        if len(df) > 0:
            df = df.dropna(axis=1, how='all')  # 0 行时不能 dropna，否则会删掉全部列
        processed.append(df)
all_price_df = pd.concat(processed, ignore_index=True, sort=False)
if all_price_df.empty:
    raise ValueError(
        f'日期区间 {test_start_date} ~ {test_end_date} 内没有订单数据，'
        f'请检查 A0_set_date 或源文件「TEMU-订单详情.xlsx」是否已更新到该月份。'
    )
# 空值的地方——补 0
all_price_df = all_price_df.fillna(0)
# 表头没有，则创建空列
required = ['SKU 货号', '运费税收入', '税金收入', '预估扣除金额']
for col in required:
    if col not in all_price_df.columns:
        all_price_df[col] = pd.NA
# 表头 重新排序
all_price_df = all_price_df[[
    "店铺",
    "日期",
    "参考号",
    "SKU 货号",
    "SKU ID",
    "产品单价-识别码",
    "产品单价",
    "产品单价（EUR）",
    "购买数量",
    "销售回款",
    "销售冲回",
    "运费回款",
    "运费税收入",
    "税金收入",
    "预估扣除金额",
    "运费回款（EUR）",
    "运费冲回",
    "预计收入"
]]
# 检查"产品单价（EUR）" 中 是否 存在为 0 的值，存在：报错！
if (all_price_df['产品单价（EUR）'] == 0).any():
    raise ValueError('"产品单价（EUR）" 中存在为 0 的值，请检查：「紫鸟店铺」的"TEMU-订单详情.xlsx"是否有空行、是否有新币种！')

# 将处理后的数据保存到新的Excel文件
output_file_path = fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\TEMU-产品单价\(处理完成)TEMU-产品单价.xlsx'
all_price_df.to_excel(output_file_path, index=False)  # index=False表示不保存索引列
print(f"处理完成，结果已保存到{output_file_path}")
