import os
import sys
import glob
import argparse
import io
from contextlib import redirect_stdout, nullcontext
import pandas as pd
import pymysql.cursors
import importlib.util
from pathlib import Path
# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(p / "ensure_project_root.py" for p in Path(__file__).resolve().parents if (p / "ensure_project_root.py").is_file())
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.sku_mapping import sku_mappings
from common.platform_shop import map_shop_to_region
from config.A0_set_date import shared_date, folder_name
from common.style import Color
from config.A0_paths import DESKTOP_ROOT

from database.db_connection import get_db_manager  # noqa: E402


_COLLATE = "utf8mb4_unicode_ci"


def get_mapping_from_db(main_df):
    """
    从数据库 sales_order_returned 表获取映射信息；映射站点优先取退件表 market_region，
    为空时回退 platform_shop.market_region（关联键：platform + platform_site + shop_name_en）。

    Args:
        main_df: 主数据框，需包含 '退件号' 和 'SKU' 列

    Returns:
        映射数据的 DataFrame，包含 return_doc_no, warehouse_sku, shop_name_en, market_region, platform_sku
    """
    try:
        # 获取数据库管理器
        db_manager = get_db_manager()
        
        # 提取去重后的退件号列表
        return_doc_no_list = main_df['退件号'].dropna().unique().tolist()
        
        if not return_doc_no_list:
            return pd.DataFrame(columns=['return_doc_no', 'warehouse_sku', 'shop_name_en', 'market_region', 'platform_sku'])
        
        # 去除空值和空字符串
        return_doc_no_list = [str(x).strip() for x in return_doc_no_list if x and str(x).strip()]
        
        if not return_doc_no_list:
            return pd.DataFrame(columns=['return_doc_no', 'warehouse_sku', 'shop_name_en', 'market_region', 'platform_sku'])
        
        print(f"{Color.CYAN}[查询] 准备查询 {len(return_doc_no_list)} 个退件号的映射信息...{Color.RESET}")
        
        # 构建 SQL（使用参数化查询防止 SQL 注入）
        # 只用 return_doc_no 筛选，获取所有相关的 warehouse_sku
        placeholders_doc = ','.join(['%s'] * len(return_doc_no_list))
        
        sql = f"""
            SELECT DISTINCT
                r.return_doc_no,
                r.warehouse_sku,
                r.shop_name_en,
                COALESCE(
                    NULLIF(TRIM(r.market_region), ''),
                    NULLIF(TRIM(ps.market_region), '')
                ) AS market_region,
                r.platform_sku
            FROM sales_order_returned AS r
            LEFT JOIN platform_shop AS ps ON
                TRIM(r.platform) COLLATE {_COLLATE} = TRIM(ps.platform) COLLATE {_COLLATE}
                AND TRIM(IFNULL(r.platform_site, '')) COLLATE {_COLLATE}
                    = TRIM(IFNULL(ps.platform_site, '')) COLLATE {_COLLATE}
                AND TRIM(IFNULL(r.shop_name_en, '')) COLLATE {_COLLATE}
                    = TRIM(IFNULL(ps.shop_name_en, '')) COLLATE {_COLLATE}
            WHERE r.return_doc_no IN ({placeholders_doc})
                AND r.shop_name_en IS NOT NULL
                AND TRIM(r.shop_name_en) <> ''
        """
        
        # 执行查询
        params = tuple(return_doc_no_list)
        conn = db_manager.get_connection()
        cur = conn.cursor(pymysql.cursors.DictCursor)
        try:
            cur.execute(sql, params)
            results = cur.fetchall()
        finally:
            cur.close()
            conn.close()
        
        # 转换为 DataFrame
        if results:
            df = pd.DataFrame(results)
            print(f"{Color.GREEN}[OK] 从数据库查询到 {len(df)} 条映射数据（包含 {df['return_doc_no'].nunique()} 个退件号）{Color.RESET}")
            
            # 显示数据库中的 warehouse_sku 示例（用于调试）
            if len(df) > 0:
                sample_sku = df['warehouse_sku'].iloc[0] if len(df) > 0 else ''
                # print(f"{Color.CYAN}[示例] 数据库中的 warehouse_sku: {sample_sku}{Color.RESET}")
            
            return df
        else:
            print(f"{Color.YELLOW}[警告] 数据库中未查询到匹配的映射数据{Color.RESET}")
            return pd.DataFrame(columns=['return_doc_no', 'warehouse_sku', 'shop_name_en', 'market_region', 'platform_sku'])
            
    except Exception as e:
        print(f"{Color.RED}[错误] 从数据库查询映射数据失败: {e}{Color.RESET}")
        print(f"{Color.YELLOW}[提示] 将尝试使用原有的 RPA 文件映射方式{Color.RESET}")
        return None


def normalize_sku_for_output(df):
    """统一清洗 SKU，供 G3 映射原始采购价等下游使用（数据库/RPA 路径均执行）"""
    df = df.copy()
    df['SKU'] = df['SKU'].astype(str).str.strip().apply(
        lambda x: x.replace('900008-', '').replace('-ECO', '').replace('-BC', '').replace('-AT-01', '').strip()
    )
    df.loc[df['SKU'] == '20007-YES', 'SKU'] = 'SK20007'
    return df


def main(check_only: bool = False) -> list[str]:
    """
    正常模式：完整执行并输出过程日志 + 保存结果文件。
    检查模式（--check）：完整执行，但仅输出“映射结果统计/告警”。
    """

    def p(msg: str) -> None:
        if not check_only:
            print(msg)

    # 检查模式：静默所有过程输出（包括被调用函数里的 print），只保留最终检查结果
    _stdout_cm = redirect_stdout(io.StringIO()) if check_only else nullcontext()

    with _stdout_cm:
        # TODO 文件路径！！！
        main_file_path = fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\二次上架\鸿羽仓-二次上架明细-{shared_date}.xls'
        main_df = pd.read_excel(main_file_path)
        # 不要"实收数量" 为 0 的
        main_df = main_df[main_df['实收数量'] != 0]
        # SKU 清洗
        main_df['SKU'] = main_df['SKU'].astype(str).str.strip().str.replace(r'((-[1-8]|-AT|--5|-BC|-FB|-BTL)+)$', '', regex=True)

        # ========== 优化部分：尝试从数据库获取映射数据 ==========
        p(f"\n{Color.CYAN}[步骤 1] 尝试从数据库获取账号和站点映射信息...{Color.RESET}")
        db_mapping_df = get_mapping_from_db(main_df)

        use_db_mapping = False
        if db_mapping_df is not None and not db_mapping_df.empty:
            use_db_mapping = True
            db_mapping_df['cleaned_sku'] = db_mapping_df['warehouse_sku'].astype(str).str.strip().str.replace(
                r'((-[1-8]|-AT|--5|-BC|-FB|-BTL)+)$', '', regex=True
            )
            p(f"{Color.CYAN}[匹配] 使用退件号 + 清洗后的SKU 进行匹配...{Color.RESET}")

            main_df = main_df.merge(
                db_mapping_df[['return_doc_no', 'cleaned_sku', 'shop_name_en', 'market_region', 'platform_sku']],
                left_on=['退件号', 'SKU'],
                right_on=['return_doc_no', 'cleaned_sku'],
                how='left'
            )

            matched_by_sku = int(main_df['shop_name_en'].notna().sum())
            p(f"{Color.GREEN}[匹配结果] 通过 退件号+SKU 匹配成功: {matched_by_sku} 条{Color.RESET}")

            unmatched_mask = main_df['shop_name_en'].isna()
            unmatched_count = int(unmatched_mask.sum())
            if unmatched_count > 0:
                p(f"{Color.YELLOW}[补充匹配] 还有 {unmatched_count} 条未匹配，尝试只用退件号匹配...{Color.RESET}")
                db_mapping_first = db_mapping_df.groupby('return_doc_no').first().reset_index()
                unmatched_df = main_df[unmatched_mask].copy()
                unmatched_df = unmatched_df.merge(
                    db_mapping_first[['return_doc_no', 'shop_name_en', 'market_region', 'platform_sku']],
                    left_on='退件号',
                    right_on='return_doc_no',
                    how='left',
                    suffixes=('', '_补充')
                )
                main_df.loc[unmatched_mask, 'shop_name_en'] = unmatched_df['shop_name_en_补充'].values
                main_df.loc[unmatched_mask, 'market_region'] = unmatched_df['market_region_补充'].values
                main_df.loc[unmatched_mask, 'platform_sku'] = unmatched_df['platform_sku_补充'].values

                matched_by_doc = int(main_df['shop_name_en'].notna().sum()) - matched_by_sku
                p(f"{Color.GREEN}[补充匹配] 通过 退件号 补充匹配成功: {matched_by_doc} 条{Color.RESET}")

            main_df = main_df.rename(columns={
                'shop_name_en': '合并-映射账号',
                'market_region': '映射站点',
                'platform_sku': '平台SKU',
            })
            main_df = main_df.drop(columns=[c for c in ['return_doc_no', 'cleaned_sku'] if c in main_df.columns], errors='ignore')
        else:
            p(f"{Color.YELLOW}[提示] 数据库映射失败或无数据，使用原有的 RPA 文件映射方式{Color.RESET}")

        # 数据库失败/或映射缺口较大时，用 RPA 补充
        if (not use_db_mapping) or (('合并-映射账号' in main_df.columns) and (main_df['合并-映射账号'].isna().sum() > len(main_df) * 0.2)):
            p(f"\n{Color.CYAN}[步骤 2] 使用 RPA 文件进行补充映射...{Color.RESET}")

            folder_path = r'\\Betohow\数据报表\RPA\二次上架-数据查询\订单管理'
            file_list = glob.glob(os.path.join(folder_path, '*.csv'))
            columns_to_keep = ['店铺账号', '销售参考号', 'SKU', '产品数量', '产品SKU']
            merged_df = pd.DataFrame(columns=columns_to_keep)
            for file in file_list:
                temp_df = pd.read_csv(file, low_memory=False)
                temp_df = temp_df[columns_to_keep]
                merged_df = pd.concat([merged_df, temp_df], ignore_index=True)

            merged_df_cleaned = merged_df.astype(str).apply(lambda col: col.map(lambda x: x.strip('="')))
            output_path = folder_path + '\\all-订单管理查询.xlsx'
            merged_df_cleaned.to_excel(output_path, index=False)
            p(f"表格合并完成，结果已保存到：{output_path}")

            product_map_sku_path = output_path
            main_df_1 = sku_mappings(
                main_df=main_df,
                main_sku='订单参考号',
                map_sku_path=product_map_sku_path,
                map_old_sku="销售参考号",
                map_new_sku="店铺账号",
                map_sku_sheet="Sheet1"
            )

            main_df_1 = main_df_1.rename(columns={'映射店铺账号': '映射店铺账号-1'})
            main_df_2 = sku_mappings(
                main_df=main_df_1,
                main_sku='参考号',
                map_sku_path=product_map_sku_path,
                map_old_sku="销售参考号",
                map_new_sku="店铺账号",
                map_sku_sheet="Sheet1"
            )
            main_df_2 = main_df_2.rename(columns={'映射店铺账号': '映射店铺账号-2'})

            product_map_sku_path = r"\\Betohow\数据报表\RPA\二次上架-数据查询\自发货\自发货-订单查询.xlsx"
            main_df_3 = sku_mappings(
                main_df=main_df_2,
                main_sku='参考号',
                map_sku_path=product_map_sku_path,
                map_old_sku="服务号",
                map_new_sku="店铺账号",
                map_sku_sheet="Worksheet 1"
            )

            if use_db_mapping and ('合并-映射账号' in main_df_3.columns):
                main_df_3['RPA-映射账号'] = main_df_3['映射店铺账号'].combine_first(main_df_3['映射店铺账号-1']).combine_first(
                    main_df_3['映射店铺账号-2'])
                main_df_3['RPA-映射账号'] = main_df_3['RPA-映射账号'].apply(
                    lambda x: x.split('(')[0].strip() if x and '(' in x else x)
                main_df_3['合并-映射账号'] = main_df_3['合并-映射账号'].fillna(main_df_3['RPA-映射账号'])
                main_df_3 = main_df_3.drop(columns=['RPA-映射账号', '映射店铺账号', '映射店铺账号-1', '映射店铺账号-2'], errors='ignore')
            else:
                main_df_3['合并-映射账号'] = main_df_3['映射店铺账号'].combine_first(main_df_3['映射店铺账号-1']).combine_first(
                    main_df_3['映射店铺账号-2'])
                main_df_3['合并-映射账号'] = main_df_3['合并-映射账号'].apply(
                    lambda x: x.split('(')[0].strip() if x and '(' in x else x)

            if (not use_db_mapping) or ('映射站点' not in main_df_3.columns) or (main_df_3['映射站点'].isna().sum() > 0):
                # 店铺账号 → 映射站点（数据源：platform_shop，替代原「站点-匹配表」）
                main_df_4 = map_shop_to_region(main_df_3, shop_col='合并-映射账号')
                if use_db_mapping and '映射站点' in main_df_3.columns:
                    main_df_4['映射站点'] = main_df_3['映射站点'].fillna(main_df_4['映射站点'])
            else:
                main_df_4 = main_df_3
        else:
            p(f"{Color.GREEN}[跳过] 数据库映射成功，跳过 RPA 文件映射步骤{Color.RESET}")
            main_df_4 = main_df

        main_df_4 = normalize_sku_for_output(main_df_4)
        if '平台SKU' not in main_df_4.columns:
            main_df_4['平台SKU'] = ''

        main_df_4 = main_df_4[
            ['退件号', '订单号', '参考号', '订单参考号', '合并-映射账号', '映射站点', '平台SKU', 'SKU', '实收数量', '良品',
             '退件费用(RMB)', '退件类型']
        ]

        # 检查模式：不写文件（避免权限/占用导致失败）
        if not check_only:
            output_file_path = (main_file_path.rsplit('\\', 1)[0] + '\\(已完成-1)' +
                                main_file_path.rsplit('\\', 1)[1].replace('xls', 'xlsx'))
            main_df_4.to_excel(output_file_path, index=False)
            p(f"处理完成，结果已保存到{output_file_path}")

        # ===== 检查结果（最终只看这一段）=====
        final_mapped_count = int(main_df_4['合并-映射账号'].notna().sum())
        final_site_mapped_count = int(main_df_4['映射站点'].notna().sum())
        final_total_count = int(len(main_df_4))

    result_lines = [
        f'{Color.CYAN}========== 映射结果统计 =========={Color.RESET}',
        f'合并-映射账号: {final_mapped_count}/{final_total_count} ({final_mapped_count/final_total_count*100:.1f}%)',
        f'映射站点: {final_site_mapped_count}/{final_total_count} ({final_site_mapped_count/final_total_count*100:.1f}%)',
    ]
    if final_mapped_count < final_total_count or final_site_mapped_count < final_total_count:
        result_lines += [
            f'{Color.YELLOW}~~~[注意]请检查，合并-映射账号、映射站点 是否都有了！！！--- ====== ---{Color.RESET}',
            '[查询SQL] SELECT id,shop_name_en,warehouse_name,return_doc_no,orig_ref_no,return_type,warehouse_sku,product_sku,cs_remark FROM `sales_order_returned` WHERE shop_name_en IS NULL;',
            '询问：惠成 ,保存 shop_name_en 后， 执行 ：reportPRA/scripts/handle/upReturnedShop.py ;  再次执行本脚本',
        ]
    else:
        result_lines.append(f'{Color.GREEN}[成功] 所有数据都已完成映射！{Color.RESET}')

    if check_only:
        print("\n".join(result_lines))
    else:
        print("\n".join(result_lines))
    return result_lines


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--check', action='store_true', help='仅输出检查结果（映射结果统计/告警）')
    args = parser.parse_args()
    main(check_only=bool(args.check))
