import pandas as pd


def sku_mappings(main_df, main_sku, map_sku_path, map_old_sku, map_new_sku, map_sku_sheet, xuan_lie_2_ci=None):
    """
        通用信息映射：用「映射表」把主表某一列的值，翻译成另一列的值，并插入到主表中。

        典型用途：旧SKU→新SKU、商品ID→产品编码、SKU→负责人/分类/头程等。

        流程概览：
          1. 按映射表文件名决定如何读 Excel（跳过表头废行等）
          2. 用映射表构建 {旧值: 新值} 字典
          3. 主表匹配列做预处理（去空格、可选剥离 -NW）
          4. 字典映射；必要时把 -NW 加回映射结果
          5. 未命中时：部分场景置空，其余保留原值
          6. 在主表 main_sku 右侧插入「映射{目标列名}」列并返回

        :param main_df: 主表 DataFrame
        :param main_sku: 主表待映射的列名（匹配键，如 SKU / 商品ID / EAN）
        :param map_sku_path: 映射表 Excel 路径
        :param map_old_sku: 映射表中「匹配前」的列名（字典的 key）
        :param map_new_sku: 映射表中「匹配后」的列名（字典的 value）
        :param map_sku_sheet: 映射表 sheet 名
        :param xuan_lie_2_ci: 二次选列标记；传 'US' 时取 map_new_sku 右侧一列（US 头程专用），默认 None
        """
    # ========== 1. 读取映射表 ==========
    # 不同源表表头结构不同，按文件名分支处理
    if '产品信息库' in map_sku_path.rsplit('\\', 1)[-1]:
        # 产品信息库：前 4 行是说明/合并表头，前 5 列是无用辅助列
        sku_map_df = pd.read_excel(map_sku_path, sheet_name=map_sku_sheet, skiprows=4)
        sku_map_df = sku_map_df.iloc[:, 5:]  # 保留第 6 列及以后
    elif '欧洲平台定价表' in map_sku_path.rsplit('\\', 1)[-1]:
        # 欧洲平台定价表：前 2 行是表头说明，跳过以便正确定位列名
        sku_map_df = pd.read_excel(map_sku_path, sheet_name=map_sku_sheet, skiprows=2)
    else:
        # 其余映射表：按默认表头直接读
        sku_map_df = pd.read_excel(map_sku_path, sheet_name=map_sku_sheet)

    # ========== 2. 构建映射字典 {old: new} ==========
    # US 站点头程：调用方传入的是「头程（RMB）」列名，实际取值需右移一列到 US 对应列
    _orig_map_new_sku = map_new_sku  # 保留原始目标列名，供后面「未匹配置空」规则使用
    if xuan_lie_2_ci == 'US':
        column_index = sku_map_df.columns.get_loc(map_new_sku)
        map_new_sku = sku_map_df.columns[column_index + 1]

    # 默认：同一 old 出现多次时，取最后一次（行号最大）——适合产品信息库等「越靠后越新」的表
    sku_mapping = {row[map_old_sku]: row[map_new_sku] for _, row in sku_map_df.iterrows()}
    # 特例：商品ID → 产品编码 时取第一次出现（行号最小），避免后写覆盖先写的主编码
    if map_old_sku == "商品ID" and map_new_sku == "产品编码":
        sku_map_df_unique = sku_map_df.drop_duplicates(subset=[map_old_sku], keep='first')
        sku_mapping = {row[map_old_sku]: row[map_new_sku] for _, row in sku_map_df_unique.iterrows()}

    # ========== 3. 主表准备：插入结果列 + 匹配键预处理 ==========
    # 结果列插在 main_sku 右侧，列名为「映射{目标列名}」
    new_col_name = f"映射{map_new_sku}"
    insert_pos = main_df.columns.get_loc(main_sku) + 1
    main_df.insert(insert_pos, new_col_name, None)

    main_series = main_df[main_sku]
    # EAN：数值型键，不转字符串、不处理 -NW 后缀
    if main_sku == 'EAN':
        main_series_no_nw = main_series
        nw_mask = False  # EAN 不可能带 -NW
    else:
        # 其余列：统一转字符串并去首尾空格，避免 Excel 类型/空格导致匹配失败
        main_series = main_series.astype(str).str.strip()

        # -NW 表示「无仓/特殊库存」等后缀：映射表一般只有无后缀的主码，故多数场景先剥再匹配
        # 库存周转明细本身可能需要保留完整 SKU，故该表不剥 -NW（当前业务实际较少用此映射）
        if '库存周转明细' in map_sku_path.rsplit('\\', 1)[-1]:
            main_series_no_nw = main_series
            nw_mask = False
        else:
            # nw_mask：标记哪些行原本带 -NW，映射成功后要按规则加回
            nw_mask = main_series.str.endswith('-NW', na=False)
            main_series_no_nw = main_series.mask(
                nw_mask,
                main_series.str.replace('-NW$', '', regex=True)
            )

    # ========== 4. 执行映射 ==========
    mapped_series = main_series_no_nw.map(sku_mapping)

    # ========== 5. 按规则加回 -NW ==========
    # 仅当「映射结果是商品ID」或「主表键是商品ID / 产品代码（SKU）」时，才把 -NW 缀回结果
    # 分类、负责人等属性列即使原 SKU 带 -NW，也不给属性值加 -NW
    if has_nw := (map_new_sku == '商品ID' or main_sku in {'商品ID', '产品代码（SKU）'}):
        no_nw_suffix = {"二级分类", "三级分类", "负责人", "销售负责人"}
        if map_new_sku not in no_nw_suffix:
            mapped_series = mapped_series.mask(nw_mask, mapped_series + '-NW')

    # ========== 6. 未匹配处理 ==========
    # 默认：映射不到则保留主表原值
    # 例外：下列主表键 / 目标列视为「查属性」，查不到应置空，避免把 SKU 本身误当成属性值
    none_main_sku = {'参考号', '订单参考号', '退件号', 'SKC识别码', '店铺英文名', '店铺英文名-站点'}
    none_map_new_sku = {"二级分类", "三级分类", "运营模式", "供应商", "AMZ新老品", "本土平台新老品",
                        "产品状态", "销售负责人-平台", "销售负责人-站点", "销售负责人-SKU", "负责人",
                        "销售负责人", "原始采购价", "申报价格", "关税（含税）", "头程（RMB）", "HY-DE",
                        "销售负责人-SKU（AMAZON-EU）", "fba fees", "平台费（佣金）", "佣金比", "VAT税",
                        "国家或地区代码",
                        "运费回款（EUR）", "产品单价（EUR）", "站点", "平台", "德国发MF/FBC运费（EUR）", "德国发FBA运费（EUR）",
                        "销售经理"}

    # US 头程二次选列后实际列名可能变成 Unnamed:*，仍用原始「头程（RMB）」判断是否置空
    fill_na = (
        (main_sku in none_main_sku)
        or (map_new_sku in none_map_new_sku)
        or (_orig_map_new_sku in none_map_new_sku)
    )
    mapped_series = mapped_series.where(mapped_series.notna(),
                                        None if fill_na else main_series)

    # ========== 7. 写回并返回 ==========
    main_df = main_df.copy()
    main_df[new_col_name] = mapped_series

    # ------------------- 日志（已关闭）-------------------
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
