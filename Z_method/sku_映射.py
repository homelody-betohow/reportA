import pandas as pd


def sku_mappings(main_df, main_sku, map_sku_path, map_old_sku, map_new_sku, map_sku_sheet, xuan_lie_2_ci=None):
    """
        映射sku
        :param main_df: 主表
        :param main_sku: 主表需要映射的列
        :param map_sku_path: 映射表的路径
        :param map_old_sku: 映射表的映射前的列
        :param map_new_sku: 映射表的映射后的列
        :param map_sku_sheet: 映射表的sheet页的名字
        :param xuan_lie_2_ci: 是否要2次选列，US的头程需要！ 默认：None
        """
    # 产品信息库的表 需要跳过没用的行和列
    if '产品信息库' in map_sku_path.rsplit('\\', 1)[-1]:
        # 跳过前4行，删除前5列
        sku_map_df = pd.read_excel(map_sku_path, sheet_name=map_sku_sheet, skiprows=4)
        sku_map_df = sku_map_df.iloc[:, 5:]  # 保留第6列及以后的所有列
    elif '欧洲平台定价表' in map_sku_path.rsplit('\\', 1)[-1]:
        # 方便定位-具体的映射行   跳过前2行
        sku_map_df = pd.read_excel(map_sku_path, sheet_name=map_sku_sheet, skiprows=2)
    else:
        sku_map_df = pd.read_excel(map_sku_path, sheet_name=map_sku_sheet)

    # US 站点 的 头程需要2次定位到下一列，也就是 us 列
    _orig_map_new_sku = map_new_sku
    if xuan_lie_2_ci == 'US':
        column_index = sku_map_df.columns.get_loc(map_new_sku)
        map_new_sku = sku_map_df.columns[column_index + 1]
    # 构建 SKU 映射字典 只保留 map_old_sku 最后一次出现的行（行号最大的）  默认情况！
    sku_mapping = {row[map_old_sku]: row[map_new_sku] for _, row in sku_map_df.iterrows()}
    # 商品ID 映射 SKU  时，则 只保留 map_old_sku 第一次出现的行（行号最小的）
    if map_old_sku == "商品ID" and map_new_sku == "产品编码":
        # 构建 SKU 映射字典   # 只保留 map_old_sku 第一次出现的行（行号最小的）
        sku_map_df_unique = sku_map_df.drop_duplicates(subset=[map_old_sku], keep='first')
        sku_mapping = {row[map_old_sku]: row[map_new_sku] for _, row in sku_map_df_unique.iterrows()}

    # ------------------- 主表处理 -------------------
    # 新建列
    new_col_name = f"映射{map_new_sku}"
    insert_pos = main_df.columns.get_loc(main_sku) + 1
    main_df.insert(insert_pos, new_col_name, None)

    main_series = main_df[main_sku]
    # 1) EAN 列：不做字符串化，也不处理 -NW
    if main_sku == 'EAN':
        main_series_no_nw = main_series
        nw_mask = False  # EAN 一定没有 -NW
    else:
        # 2) 其余列：先转成字符串并去空格
        main_series = main_series.astype(str).str.strip()

        # 根据文件名决定是否保留 -NW（现在实际上，并没有映射 库存周转明细 表）
        if '库存周转明细' in map_sku_path.rsplit('\\', 1)[-1]:
            # 库存周转明细：不剥 -NW
            main_series_no_nw = main_series
            nw_mask = False
        else:
            # 非库存周转明细：去掉 -NW
            nw_mask = main_series.str.endswith('-NW', na=False)
            main_series_no_nw = main_series.mask(
                nw_mask,
                main_series.str.replace('-NW$', '', regex=True)
            )

    # ------------------- 字典映射 -------------------
    mapped_series = main_series_no_nw.map(sku_mapping)

    # ------------------- 加回 -NW -------------------
    # 映射后的数据是'商品ID' or 映射前的数据是：'产品代码（SKU）'、'商品ID'  则：映射后的数据才带上 “-NW”
    # has_nw 被创建为一个布尔值，随后被用于 if has_nw: 的判断。
    if has_nw := (map_new_sku == '商品ID' or main_sku in {'商品ID', '产品代码（SKU）'}):
        # 下列列名不加 -NW
        no_nw_suffix = {"二级分类", "三级分类", "负责人", "销售负责人"}
        if map_new_sku not in no_nw_suffix:
            mapped_series = mapped_series.mask(nw_mask, mapped_series + '-NW')

    # ------------------- 未匹配值处理 -------------------
    # 映射不到，保留原值（默认）
    # 映射不到，置空的几种情况
    none_main_sku = {'参考号', '订单参考号', '退件号', 'SKC识别码', '店铺英文名', '店铺英文名-站点'}
    none_map_new_sku = {"二级分类", "三级分类", "运营模式", "供应商", "AMZ新老品", "本土平台新老品",
                        "产品状态", "销售负责人-平台", "销售负责人-站点", "销售负责人-SKU", "负责人",
                        "销售负责人", "原始采购价", "申报价格", "关税（含税）", "头程（RMB）", "HY-DE",
                        "销售负责人-SKU（AMAZON-EU）", "fba fees", "平台费（佣金）", "佣金比", "VAT税",
                        "国家或地区代码",
                        "运费回款（EUR）", "产品单价（EUR）", "站点", "平台", "德国发MF/FBC运费（EUR）", "德国发FBA运费（EUR）",
                        "销售经理"}

    # US 头程二次选列后列名可能为 Unnamed:*，仍按原始「头程（RMB）」规则：未匹配置空
    fill_na = (
        (main_sku in none_main_sku)
        or (map_new_sku in none_map_new_sku)
        or (_orig_map_new_sku in none_map_new_sku)
    )
    mapped_series = mapped_series.where(mapped_series.notna(),
                                        None if fill_na else main_series)

    # 写回主表
    main_df = main_df.copy()
    main_df[new_col_name] = mapped_series

    # ------------------- 日志 -------------------
    # log_df = main_df[[main_sku, new_col_name]].copy()
    # log_df['行号'] = range(2, len(log_df) + 2)
    # log_df = log_df[
    #     log_df[new_col_name].notna() & (log_df[main_sku].astype(str).str.strip() != log_df[new_col_name])]
    # print("\n" + "-" * 60 + " SKU 映射日志 " + "-" * 60)
    # if not log_df.empty:
    #     print("行号\t\t旧SKU\t\t\t\t新SKU")
    #     for _, r in log_df.iterrows():
    #         print(f"{int(r['行号'])}\t\t{r[main_sku]}\t\t\t{r[new_col_name]}")
    # else:
    #     print("（无 SKU 映射记录）")
    # print("=" * 120)
    # print(f"共映射了{len(log_df)}个sku！！！！")

    return main_df
