import argparse
import io
from contextlib import nullcontext, redirect_stdout
import importlib.util
from pathlib import Path

import numpy as np
import pandas as pd

# 须在 import A_报表 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from A_报表.Z_method.style import Color  # noqa: E402
from A_报表.Z_method.sku_映射 import sku_mappings  # noqa: E402
from A_报表.Z_method.platform_shop import apply_lm_fr_region_suffix, map_region_to_platform  # noqa: E402
from A_报表.A0_设置_时间段.A0_set_date import folder_name, shared_date  # noqa: E402
from A_报表.A0_设置_时间段.A0_paths import BTH_ALL_SKU_DETAIL_PATH, DESKTOP_ROOT  # noqa: E402


def main(check_only: bool = False) -> list[str]:
    """
    正常模式：完整执行并输出过程日志 + 保存结果文件。
    检查模式（--check）：静默过程输出、且不写文件，仅输出“映射原始采购价”的检查结果。
    """

    def p(msg: str) -> None:
        if not check_only:
            print(msg)

    # 检查模式：静默所有过程输出（包含被调用函数里的 print），只保留最终检查结果
    _stdout_cm = redirect_stdout(io.StringIO()) if check_only else nullcontext()

    with _stdout_cm:
        # TODO 文件路径！！！
        main_file_path = fr'{DESKTOP_ROOT}\{folder_name}{shared_date}\二次上架\(已完成-1)鸿羽仓-二次上架明细-{shared_date}.xlsx'
        main_df = pd.read_excel(main_file_path)

        # 映射 平台（数据源：platform_shop）
        main_df_1 = map_region_to_platform(main_df, site_col='映射站点')


        # 通用处理函数：处理 LM_BC_FR 、 LM_RP_FR 的二次上架订单映射及站点后缀
        def process_lm_orders(df: pd.DataFrame, account_name: str, order_file_path: str) -> None:
            mask = df['合并-映射账号'] == account_name
            p(f"\n{account_name}的二次上架订单:")
            order_df = pd.read_excel(order_file_path)
            for col in ['销售参考号', 'SKU']:
                order_df[col] = order_df[col].str.replace(r'^="(.*)"$', r'\1', regex=True)
            mapping = order_df.set_index('销售参考号')['SKU'].to_dict()

            orig_order = df.loc[mask, '订单参考号']
            orig_ref = df.loc[mask, '参考号'] if '参考号' in df.columns else pd.Series(np.nan, index=orig_order.index)
            mapped_from_order = orig_order.map(mapping)
            mapped_from_ref = orig_ref.map(mapping)
            mapped_sku = mapped_from_order.combine_first(mapped_from_ref)

            found_mask = mapped_sku.notna()
            matched_by_order = mapped_from_order.notna()
            matched_by_ref = mapped_from_ref.notna() & ~matched_by_order

            # 优先保留 G2 从数据库带回的 平台SKU，RPA 仅补充空值
            existing_sku = df.loc[mask, '平台sku'].fillna('').astype(str).str.strip()
            df.loc[mask, '平台sku'] = existing_sku.mask(existing_sku.isin(['', 'nan', 'None']), mapped_sku.astype(str))

            p(
                f"[OK] 已映射 {int(found_mask.sum())} 行"
                f"（订单参考号 {int(matched_by_order.sum())}，参考号 {int(matched_by_ref.sum())}），"
                f"未找到 {int((~found_mask).sum())} 行"
            )

            apply_lm_fr_region_suffix(
                df,
                shop_col='合并-映射账号',
                shops=(account_name,),
                inplace=True,
            )


        # 执行处理：对两个法国账号进行同样的操作
        path = r'\\Betohow\数据报表\RPA\二次上架-数据查询\订单管理\all-订单管理查询.xlsx'
        if '平台SKU' in main_df_1.columns:
            main_df_1['平台sku'] = main_df_1['平台SKU'].fillna('').astype(str).str.strip().replace({'nan': '', 'None': ''})
        else:
            main_df_1['平台sku'] = ''
        lm_bc = 'LM_BC_FR' in main_df_1['合并-映射账号'].values
        lm_rp = 'LM_RP_FR' in main_df_1['合并-映射账号'].values
        for account in ['LM_BC_FR', 'LM_RP_FR']:
            if account in main_df_1['合并-映射账号'].values:
                process_lm_orders(main_df_1, account, path)

        # 重命名
        main_df_1 = main_df_1.rename(columns={'映射站点': '站点'})
        main_df_1 = main_df_1.rename(columns={'映射平台': '平台'})
        # 构建识别码
        main_df_1['SKU-站点识别码'] = main_df_1['站点'] + main_df_1['SKU']
        main_df_1['SKU-平台识别码'] = main_df_1['平台'] + main_df_1['SKU']

        # TODO 处理——退件费用(EUR) 按 实收数量 占比分摊
        # 0. 确保分摊运费列存在且先置 0
        main_df_1['分摊运费(EUR)'] = 0.0

        # ---------- 1. OTTO 运费分摊 ----------
        otto_mask = main_df_1['站点'].str.contains('OTTO', na=False)
        otto_df = main_df_1[otto_mask].copy()
        if not otto_df.empty:
            otto_df['退件类型'] = otto_df['退件类型'].fillna('')

            def get_otto_fee(return_type: str) -> float:
                if pd.isna(return_type):
                    raise ValueError("退件类型为 NaN，无法处理。")
                if return_type.startswith(('买家退件', '认领')):
                    return 6.1
                if return_type.startswith('物流退件'):
                    return 5.5
                raise ValueError(f"不支持的退件类型: '{return_type}'")

            fee_per_return = (
                otto_df.groupby('退件号')['退件类型']
                .first()
                .apply(get_otto_fee)
                .rename('total_fee')
            )
            otto_df = otto_df.merge(fee_per_return, left_on='退件号', right_index=True)
            otto_df['qty_sum'] = otto_df.groupby('退件号')['实收数量'].transform('sum')
            otto_df['ratio'] = otto_df['实收数量'] / otto_df['qty_sum']
            otto_df['分摊运费(EUR)'] = otto_df['total_fee'] * otto_df['ratio']
            main_df_1.loc[otto_mask, '分摊运费(EUR)'] = otto_df['分摊运费(EUR)'].values

        # ---------- 2. 实际退件费用（RMB）分摊 ----------
        fee_first = main_df_1.groupby('退件号')['退件费用(RMB)'].transform('first')
        qty_sum_all = main_df_1.groupby('退件号')['实收数量'].transform('sum')
        ratio = main_df_1['实收数量'] / qty_sum_all
        main_df_1['实际-退件费用(RMB)'] = np.where(qty_sum_all == 0, 0, ratio * fee_first)

        # ---------- 3. 币种转换 ----------
        main_df_1['实际-退件费用(EUR)'] = main_df_1['实际-退件费用(RMB)'] / 7.3
        main_df_1['退件费用(EUR)'] = main_df_1['实际-退件费用(EUR)'] + main_df_1['分摊运费(EUR)']

        # ---------- 4. 列名整理 + 映射原始采购价 ----------
        main_df_1.rename(columns={'退件费用(RMB)': '原-退件费用(RMB)'}, inplace=True)
        product_map_sku_path = BTH_ALL_SKU_DETAIL_PATH
        main_df_2 = sku_mappings(
            main_df=main_df_1,
            main_sku='SKU',
            map_sku_path=product_map_sku_path,
            map_old_sku="SKU",
            map_new_sku="原始采购价",
            map_sku_sheet='基础数据维护',
        )
        main_df_2['映射原始采购价'] = pd.to_numeric(main_df_2['映射原始采购价'], errors='coerce')

        # ===== 检查点：仅检查“映射原始采购价” =====
        total_cnt = int(len(main_df_2))
        mapped_cnt = int(main_df_2['映射原始采购价'].notna().sum())
        missing_cnt = total_cnt - mapped_cnt
        mapped_rate = (mapped_cnt / total_cnt * 100) if total_cnt else 0.0

        # 检查模式：到这里就结束（避免后续成本计算/手动回填/写文件的噪音与依赖）
        if check_only:
            pass
        else:
            # ====== 以下保持原有业务流程 ======
            print(f"{Color.YELLOW}二次上架采购成本 计算规则：{Color.RESET}")
            print(f"{Color.GREEN}1. OTTO 平台：")
            print("1.1 良品：二次上架采购成本（RMB） = 映射原始采购价 * 实收数量")
            print("1.2 次品：二次上架采购成本（RMB） = 0")
            print("1.3 NW后缀的SKU：二次上架采购成本（RMB） = 0")
            print("2. 其它平台：")
            print("2.1 NW后缀的SKU：二次上架采购成本（RMB） = 0")
            print("2.2 非NW后缀的SKU：二次上架采购成本（RMB） = 映射原始采购价 * 实收数量")
            print(f"{Color.RESET}")

            main_df_2['二次上架采购成本（RMB）'] = 0.0
            main_df_2['实收数量'] = pd.to_numeric(main_df_2['实收数量'], errors='coerce').fillna(0)
            main_df_2['良品'] = pd.to_numeric(main_df_2['良品'], errors='coerce').fillna(0)
            main_df_2['映射原始采购价'] = main_df_2['映射原始采购价'].fillna(0)

            otto_platform_mask = main_df_2['平台'].str.contains('OTTO', na=False)
            otto_good_mask = otto_platform_mask & (main_df_2['良品'] >= 1)
            main_df_2.loc[otto_good_mask, '二次上架采购成本（RMB）'] = (
                main_df_2.loc[otto_good_mask, '映射原始采购价'] * main_df_2.loc[otto_good_mask, '良品']
            )
            otto_defect_mask = otto_platform_mask & (main_df_2['良品'] <= 0)
            main_df_2.loc[otto_defect_mask, '二次上架采购成本（RMB）'] = 0

            nw_suffix_mask = otto_platform_mask & main_df_2['SKU'].astype(str).str.upper().str.endswith('NW')
            main_df_2.loc[nw_suffix_mask, '二次上架采购成本（RMB）'] = 0

            other_platform_mask = ~otto_platform_mask
            nw_suffix_mask = other_platform_mask & main_df_2['SKU'].astype(str).str.upper().str.endswith('NW')
            main_df_2.loc[nw_suffix_mask, '二次上架采购成本（RMB）'] = 0
            non_nw_suffix_mask = other_platform_mask & ~main_df_2['SKU'].astype(str).str.upper().str.endswith('NW')
            main_df_2.loc[non_nw_suffix_mask, '二次上架采购成本（RMB）'] = (
                main_df_2.loc[non_nw_suffix_mask, '映射原始采购价'] * main_df_2.loc[non_nw_suffix_mask, '良品']
            )

            print(f'\n{Color.YELLOW}=== 二次上架采购成本计算完成 ==={Color.RESET}')

            mask = (main_df_2['合并-映射账号'].isin(['LM_BC_FR', 'LM_RP_FR'])) & (main_df_2['平台sku'] == '')
            main_df_2.loc[mask, ['SKU-站点识别码', '站点']] = np.nan

            # ========= LM_BC_FR 手动映射回填（暂时注释）=========
            # 如果未来再次出现 LM_BC_FR 平台sku 为空、需要手动补 “站点/SKU-站点识别码”，
            # 再把这段恢复即可。
            #
            # manual_map_file_path = fr"{DESKTOP_ROOT}\手动-二次映射.xlsx"
            # manual_map_sheet = "二次上架-LM-BC-自发货"
            # try:
            #     manual_map_df = pd.read_excel(manual_map_file_path, sheet_name=manual_map_sheet, usecols="A:I")
            #
            #     def _norm_key(x) -> str:
            #         s = "" if pd.isna(x) else str(x)
            #         s = s.strip()
            #         s = pd.Series([s]).str.replace(r'^="(.*)"$', r"\1", regex=True).iloc[0]
            #         if s.endswith(".0") and s.replace(".", "", 1).isdigit():
            #             s = s[:-2]
            #         return s.strip()
            #
            #     key_series = manual_map_df.iloc[:, 0].map(_norm_key)
            #     site_series = manual_map_df.iloc[:, 6]
            #     child_site_code_series = manual_map_df.iloc[:, 8]
            #     manual_site_map = dict(zip(key_series, site_series))
            #     manual_child_site_code_map = dict(zip(key_series, child_site_code_series))
            #
            #     lm_bc_fill_mask = (
            #         (main_df_2['合并-映射账号'] == 'LM_BC_FR')
            #         & (main_df_2['SKU-站点识别码'].isna() | (main_df_2['SKU-站点识别码'].astype(str).str.strip() == ''))
            #     )
            #
            #     key_from_rma = (
            #         main_df_2.loc[lm_bc_fill_mask, '退件号'].map(_norm_key)
            #         if '退件号' in main_df_2.columns
            #         else pd.Series(index=main_df_2.loc[lm_bc_fill_mask].index, dtype=str)
            #     )
            #     key_from_order_ref = main_df_2.loc[lm_bc_fill_mask, '订单参考号']
            #     key_from_order_ref = (
            #         key_from_order_ref.astype(str).str.split('——', n=1).str[0].map(_norm_key)
            #         if '订单参考号' in main_df_2.columns
            #         else pd.Series(index=main_df_2.loc[lm_bc_fill_mask].index, dtype=str)
            #     )
            #     key_from_ref = (
            #         main_df_2.loc[lm_bc_fill_mask, '参考号'].map(_norm_key)
            #         if '参考号' in main_df_2.columns
            #         else pd.Series(index=main_df_2.loc[lm_bc_fill_mask].index, dtype=str)
            #     )
            #
            #     filled_site = (
            #         key_from_rma.map(manual_site_map)
            #         .combine_first(key_from_order_ref.map(manual_site_map))
            #         .combine_first(key_from_ref.map(manual_site_map))
            #     )
            #     filled_child_site_code = (
            #         key_from_rma.map(manual_child_site_code_map)
            #         .combine_first(key_from_order_ref.map(manual_child_site_code_map))
            #         .combine_first(key_from_ref.map(manual_child_site_code_map))
            #     )
            #
            #     main_df_2.loc[lm_bc_fill_mask, '站点'] = main_df_2.loc[lm_bc_fill_mask, '站点'].combine_first(filled_site)
            #     main_df_2.loc[lm_bc_fill_mask, 'SKU-站点识别码'] = main_df_2.loc[lm_bc_fill_mask, 'SKU-站点识别码'].combine_first(
            #         filled_child_site_code
            #     )
            #
            #     filled_cnt = int(filled_child_site_code.notna().sum())
            #     remain_cnt = int(main_df_2.loc[lm_bc_fill_mask, 'SKU-站点识别码'].isna().sum())
            #     print(f'{Color.YELLOW}LM_BC_FR 手动映射回填：成功 {filled_cnt} 行，仍为空 {remain_cnt} 行（通常是手动表里没有对应 key）。{Color.RESET}')
            # except FileNotFoundError:
            #     print(f'{Color.YELLOW}未找到手动映射文件：{manual_map_file_path}，将跳过 LM_BC_FR 的手动VLOOKUP自动回填。{Color.RESET}')
            # except ValueError as e:
            #     print(f'{Color.YELLOW}读取手动映射文件失败（sheet/列范围可能不对）：{e}，将跳过 LM_BC_FR 的手动VLOOKUP自动回填。{Color.RESET}')
            # ==============================================

            main_df_2 = main_df_2[
                ['退件号', '映射原始采购价', '订单号', '参考号', '订单参考号', '合并-映射账号', '站点', '平台', 'SKU-站点识别码',
                 'SKU-平台识别码', '平台sku', 'SKU', '实收数量', '良品', '原-退件费用(RMB)', '实际-退件费用(RMB)',
                 '实际-退件费用(EUR)', '分摊运费(EUR)', '退件费用(EUR)', '二次上架采购成本（RMB）', '退件类型']
            ]

            output_file_path = main_file_path.replace('已完成-1', '已完成-2')
            try:
                main_df_2.to_excel(output_file_path, index=False)
            except PermissionError:
                print(f'{Color.RED}保存失败：无法写入文件（权限被拒绝）{Color.RESET}')
                print(f'目标路径：{output_file_path}')
                print('请检查：')
                print('  1. 是否已在 Excel 中打开了该文件？请先关闭后重新运行。')
                print('  2. 文件是否为只读，或文件夹无写入权限。')
                raise SystemExit(1)
            print(f"处理完成，结果已保存到{output_file_path}")
            print('-' * 100)
            # print(f"{Color.RED}Amazon-退仓数据：二次上架采购成本（RMB） 设置为0{Color.RESET}")

    # 只输出检查结果
    result_lines = [
        f'{Color.CYAN}========== 映射原始采购价 检查结果 =========={Color.RESET}',
        f'映射原始采购价: {mapped_cnt}/{total_cnt} ({mapped_rate:.1f}%)',
    ]
    if missing_cnt > 0:
        result_lines.append(f'{Color.YELLOW}[注意] 未映射数量: {missing_cnt}{Color.RESET}')
    else:
        result_lines.append(f'{Color.GREEN}[成功] 映射原始采购价 全部都有！{Color.RESET}')

    print("\n".join(result_lines))
    return result_lines


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--check', action='store_true', help='仅输出“映射原始采购价”检查结果（不写文件、静默过程输出）')
    args = parser.parse_args()
    main(check_only=bool(args.check))
