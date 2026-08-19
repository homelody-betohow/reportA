import json
import pandas as pd
import importlib.util
from pathlib import Path
# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from config.A0_set_date import shared_date, folder_name
from common.platform_shop import map_site_to_ops_leader
from config.A0_paths import DESKTOP_ROOT

_MODULE_DIR = Path(__file__).resolve().parent
_STORE_OPERATOR_JSON = _MODULE_DIR / "map_store_operator.json"
_REPORT_PLATFORM_JSON = _MODULE_DIR / "map_report_platform.json"
_NOBODY = "nobody"
_BLANK_OWNERS = frozenset(("", "nan", "None", "NaN", "无负责人"))


def _load_json_items(path: Path) -> list[dict]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    items = payload.get("items") or []
    return [item for item in items if isinstance(item, dict)]


def _load_store_operator_map(path: Path = _STORE_OPERATOR_JSON) -> dict[tuple[str, str], str]:
    """map_store_operator.json → {(market_code, ops_owner): ops_leader}。同键后写覆盖先写。"""
    mapping: dict[tuple[str, str], str] = {}
    for item in _load_json_items(path):
        code = str(item.get("market_code") or "").strip()
        owner = str(item.get("ops_owner") or "").strip()
        leader = str(item.get("ops_leader") or "").strip()
        if not code or not owner or not leader or owner in _BLANK_OWNERS:
            continue
        mapping[(code, owner)] = leader
    return mapping


def _load_report_platform_map(path: Path = _REPORT_PLATFORM_JSON) -> dict[str, str]:
    """map_report_platform.json → {market_region: report_platform}。同键后写覆盖先写。"""
    mapping: dict[str, str] = {}
    for item in _load_json_items(path):
        site = str(item.get("market_region") or "").strip()
        plat = str(item.get("report_platform") or "").strip()
        if site and plat:
            mapping[site] = plat
    return mapping


def _map_leader_from_json(df: pd.DataFrame, mapping: dict[tuple[str, str], str]) -> pd.Series:
    """按 平台(market_code) + 销售负责人(ops_owner) 查销售经理。"""
    codes = df["平台"].astype(str).str.strip()
    owners = df["销售负责人"].astype(str).str.strip()
    keys = list(zip(codes, owners))
    return pd.Series([mapping.get(k) for k in keys], index=df.index, dtype=object)


# TODO 文件路径！！！
main_file_path = fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计\(已完成-22)订单统计-{shared_date}.xlsx"
main_df = pd.read_excel(main_file_path)
main_df_1 = main_df.copy()

# 映射 销售经理：map_store_operator.json（平台+销售负责人）→ ops_leader，再用站点 platform_shop.ops_leader 覆盖
main_df_1["销售经理"] = _map_leader_from_json(main_df_1, _load_store_operator_map())

# 优先：订单「站点」→ platform_shop.ops_leader（LM 的 -ls/-xj 后缀会先剥掉再匹配）
_shop_leader = map_site_to_ops_leader(main_df_1["站点"])
_hit_leader = _shop_leader.notna() & ~_shop_leader.astype(str).str.strip().isin(_BLANK_OWNERS)
main_df_1.loc[_hit_leader, "销售经理"] = _shop_leader.loc[_hit_leader]

# 销售负责人为 nobody / 不在映射表时，映射结果为空；与 M2/M3 口径一致，补 nobody
for _col in ('销售经理', '销售负责人'):
    _s = main_df_1[_col]
    _blank = _s.isna() | _s.astype(str).str.strip().isin(_BLANK_OWNERS)
    main_df_1.loc[_blank, _col] = _NOBODY
# CD 平台不归到具体销售经理，统一 nobody（原先写空）
main_df_1.loc[main_df_1['平台'] == 'CD', '销售经理'] = _NOBODY
# 新增列：仓租识别码，平台+产品状态
main_df_1['仓租识别码'] = main_df_1['平台'].astype(str) + main_df_1['产品状态'].astype(str)

# 表头 重新排序
main_df_1 = main_df_1[[
    "商品ID",
    "SKU",
    "站点",
    "平台",
    "平台商品ID识别码",
    "站点商品ID识别码",
    "仓租识别码",
    "产品状态",
    "二级分类",
    "三级分类",
    "销售经理",
    "销售负责人",
    "销量",
    "平台销售额",
    "退款数量",
    "重发数量",
    "退款额",
    "销售额",
    "测评费",
    "秒杀费",
    "广告费(AMZ)",
    "广告费(非AMZ)",
    "广告费合计",
    "平台费(AMZ)",
    "平台费(非AMZ)",
    "平台费合计",
    "销售税(AMZ)",
    "销售税(非AMZ)",
    "销售税合计",
    "派送费",
    "海外仓仓租费",
    "FBA仓租费",
    "仓租合计",
    "提现费",
    "月租",
    "赔偿金额",
    "其他分摊费用",
    "二次上架数量",
    "二次上架金额",
    "订单采购成本",
    "重发采购成本",
    "二次上架采购成本",
    "采购成本",
    "头程",
    "关税",
    "毛利",
    "毛利率",
    "运营模式",
    "供应商"
]]

# 空值的地方——补 0
main_df_1 = main_df_1.fillna(0)
# 下面这些列，值为 0 的地方，替换为 空（销售经理/销售负责人已是 nobody，不要再清成空）
cols = ['商品ID', 'SKU', '平台', '站点', '平台商品ID识别码', '站点商品ID识别码', '仓租识别码', '产品状态',
        '二级分类', '三级分类', '运营模式', '供应商']
main_df_1[cols] = main_df_1[cols].replace(0, '')

# 刷新一下 仓租识别码
main_df_1['仓租识别码'] = main_df_1['平台'].astype(str) + main_df_1['产品状态'].astype(str)

if folder_name == '日报':
    # 解析 shared_date（如 7.1-7.6），格式化为 mm月dd日
    start_md, end_md = shared_date.split('-')
    month_1, day_1 = map(int, start_md.split('.'))
    month_2, day_2 = map(int, end_md.split('.'))
    full_date = f"{month_1:02d}/{day_1:02d}-{month_2:02d}/{day_2:02d}"
    print(full_date)
    # 找到“商品ID”列的索引位置
    product_id_index = main_df_1.columns.get_loc('商品ID')
    # 在“商品ID”列前插入“日期”列
    main_df_1.insert(product_id_index, '日期', full_date)
    # 删除 指定列   日报 不要 '订单采购成本', '重发采购成本', '二次上架采购成本'
    main_df_1.drop(columns=['订单采购成本', '重发采购成本', '二次上架采购成本'], inplace=True)
    # 平台（报表）：map_report_platform.json 站点 → 报表平台；未命中保留原站点
    # 插在「站点」右侧，与原先 sku_mappings 插入位置一致
    _sites = main_df_1["站点"].astype(str).str.strip()
    _report_plat = _sites.map(_load_report_platform_map())
    _report_plat = _report_plat.where(_report_plat.notna(), _sites)
    main_df_1.insert(main_df_1.columns.get_loc("站点") + 1, "平台（报表）", _report_plat)
    main_df_1 = main_df_1.rename(columns={"仓租识别码": "平台（报表识别码）"})
    main_df_1["平台（报表识别码）"] = main_df_1["平台（报表）"] + main_df_1["商品ID"].astype(str)

# 去除 整张表 的前后空格
for col in main_df_1.columns:
    main_df_1[col] = main_df_1[col].apply(lambda x: x.strip() if isinstance(x, str) else x)

# 保存结果
output_path = main_file_path.rsplit('\\', 2)[0] + f'\\{shared_date}--{folder_name}.xlsx'
main_df_1.to_excel(output_path, index=False)
print(f"结果已保存到文件：{output_path}")
