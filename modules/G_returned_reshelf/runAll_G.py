from __future__ import annotations

import argparse
import importlib.util
import sys
from pathlib import Path

_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.style import Color
from common.runall_utils import (
    setup_console_encoding,
    run_script,
    get_scripts_in_order,
    extract_output_file_path,
)

# 设置 Windows 控制台编码
setup_console_encoding()


def _is_g1(script_path: Path) -> bool:
    # 判断是否是 G1 这一步（输出需要手动 RPA 查询的内容）
    return script_path.name.startswith("G1_")


def _is_g2(script_path: Path) -> bool:
    # 判断是否是 G2 这一步（映射 RPA 查询结果，需要检查映射是否完整）
    return script_path.name.startswith("G2_")


def _is_g3(script_path: Path) -> bool:
    # 判断是否是 G3 这一步（计算金额和成本，需要检查映射结果）
    return script_path.name.startswith("G3_")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="按顺序执行本目录下所有 G*.py 脚本（排除 runAll_G.py）"
    )
    parser.add_argument(
        "--continue-on-error",
        action="store_true",
        help="某一步失败后继续执行后续脚本（默认：失败即停止）",
    )
    parser.add_argument(
        "--stop-after-g1",
        action="store_true",
        help="执行到 G1_*.py 完成后立即停止（用于你手动 RPA 查询）",
    )
    parser.add_argument(
        "--pause-after-g1",
        action="store_true",
        help="执行到 G1_*.py 完成后暂停，按回车后继续执行后续脚本",
    )
    parser.add_argument(
        "--stop-after-g2",
        action="store_true",
        help="执行到 G2_*.py 完成后立即停止（用于你检查 RPA 查询结果）",
    )
    parser.add_argument(
        "--pause-after-g2",
        action="store_true",
        help="执行到 G2_*.py 完成后暂停，按回车后继续执行后续脚本",
    )
    parser.add_argument(
        "--stop-after-g3",
        action="store_true",
        help="执行到 G3_*.py 完成后立即停止（用于你检查映射结果）",
    )
    parser.add_argument(
        "--pause-after-g3",
        action="store_true",
        help="执行到 G3_*.py 完成后暂停，按回车后继续执行后续脚本",
    )
    args, unknown = parser.parse_known_args(argv)
    if unknown:
        print(f"[runAll_G] 忽略多余参数：{unknown}")

    folder = Path(__file__).resolve().parent
    scripts = get_scripts_in_order(folder, Path(__file__).name)

    if not scripts:
        print(f"[runAll_G] 未找到可执行脚本：{folder}")
        return 0

    print(f"[runAll_G] 将执行 {len(scripts)} 个脚本，目录：{folder}")

    failed: list[tuple[str, int]] = []

    for idx, script in enumerate(scripts, start=1):
        print(f"[runAll_G] ({idx}/{len(scripts)}) 开始：{Color.YELLOW} {script.name} {Color.RESET}")
        code, output = run_script(script)
        if code == 0:
            print(f"[runAll_G] ({idx}/{len(scripts)}) 完成：{script.name}")

            if _is_g1(script):
                print(f"\n{Color.YELLOW}{'=' * 80}{Color.RESET}")
                print(f"{Color.YELLOW}[runAll_G] 重要提示：G1 已完成，请按照上面的提示操作：{Color.RESET}")
                print(f"{Color.YELLOW}{'=' * 80}{Color.RESET}")
                print(f"{Color.YELLOW}【需要手动 RPA 查询的位置】{Color.RESET}")
                print(f"{Color.YELLOW}  - 自发货查询：{Color.RESET}")
                print(f"    复制上面的「参考号_str」到自发货系统查询")
                print(f"{Color.YELLOW}  - 订单管理查询：{Color.RESET}")
                print(f"    复制上面的「订单参考号_str」到订单管理系统查询")
                print(f"\n{Color.YELLOW}【查询结果保存位置】{Color.RESET}")
                print(f"  \\\\Betohow\\数据报表\\RPA\\二次上架-数据查询\\自发货")
                print(f"  \\\\Betohow\\数据报表\\RPA\\二次上架-数据查询\\订单管理")
                print(f"\n{Color.YELLOW}【操作步骤】{Color.RESET}")
                print(f"{Color.YELLOW}  1. 复制上面黄色的参考号（自发货）到 RPA 系统查询{Color.RESET}")
                print(f"{Color.YELLOW}  2. 复制上面黄色的订单参考号（订单管理）到 RPA 系统查询{Color.RESET}")
                print(f"{Color.YELLOW}  3. 将查询结果保存到上述对应的文件夹中{Color.RESET}")
                print(f"{Color.YELLOW}  4. 确认文件保存成功后，按回车键继续{Color.RESET}")
                print(f"{Color.YELLOW}{'=' * 80}{Color.RESET}\n")

                if args.stop_after_g1:
                    print("[runAll_G] 已执行到 G1，按参数要求停止（请先手动 RPA 查询并保存结果）")
                    return 0
                if args.pause_after_g1:
                    input("[runAll_G] >>> 按回车键继续执行后续步骤...")
                else:
                    # 默认行为：暂停等待用户确认
                    input("[runAll_G] >>> 按回车键继续执行后续步骤...")
            
            if _is_g2(script):
                # 从输出中提取文件路径
                output_file = extract_output_file_path(output, "(已完成-1)鸿羽仓-二次上架明细-*.xlsx")
                
                print(f"\n{Color.YELLOW}{'=' * 80}{Color.RESET}")
                print(f"{Color.YELLOW}[runAll_G] 重要提示：G2 已完成（RPA 查询结果映射）{Color.RESET}")
                print(f"{Color.YELLOW}{'=' * 80}{Color.RESET}")
                print(f"{Color.YELLOW}【需要检查的文件】{Color.RESET}")
                print(f"  {output_file}")
                print(f"\n{Color.YELLOW}【检查内容】{Color.RESET}")
                print(f"{Color.YELLOW}  1. 打开上述 Excel 文件{Color.RESET}")
                print(f"{Color.YELLOW}  2. 检查「合并-映射账号」列是否都有数据{Color.RESET}")
                print(f"{Color.YELLOW}  3. 检查「映射站点」列是否都有数据{Color.RESET}")
                print(f"{Color.YELLOW}  - 如果「合并-映射账号」没有 → 询问惠成{Color.RESET}")
                print(f"\n{Color.YELLOW}【操作】{Color.RESET}")
                print(f"{Color.YELLOW}  - 确认映射结果无误后，按回车键继续{Color.RESET}")
                print(f"{Color.YELLOW}{'=' * 80}{Color.RESET}\n")

                if args.stop_after_g2:
                    print("[runAll_G] 已执行到 G2，按参数要求停止（请先检查 RPA 映射结果）")
                    return 0
                if args.pause_after_g2:
                    input("[runAll_G] >>> 按回车键继续执行后续步骤...")
                else:
                    # 默认行为：暂停等待用户确认
                    input("[runAll_G] >>> 按回车键继续执行后续步骤...")
            
            if _is_g3(script):
                # 从输出中提取文件路径
                output_file = extract_output_file_path(output, "(已完成-2)鸿羽仓-二次上架明细-*.xlsx")
                
                print(f"\n{Color.YELLOW}{'=' * 80}{Color.RESET}")
                print(f"{Color.YELLOW}[runAll_G] 重要提示：G3 已完成（计算金额和成本）{Color.RESET}")
                print(f"{Color.YELLOW}{'=' * 80}{Color.RESET}")
                print(f"{Color.YELLOW}【需要检查的文件】{Color.RESET}")
                print(f"  {output_file}")
                print(f"\n{Color.YELLOW}【检查内容】{Color.RESET}")
                print(f"{Color.YELLOW}  1. 打开上述 Excel 文件{Color.RESET}")
                print(f"{Color.YELLOW}  2. 检查「映射原始采购价」列是否都有数据{Color.RESET}")
                print(f"{Color.YELLOW}  3. 检查 LM_BC_FR 的「平台sku」列是否都已映射{Color.RESET}")
                print(f"{Color.YELLOW}  4. 检查「站点」和「SKU-站点识别码」列{Color.RESET}")
                print(f"\n{Color.YELLOW}【如果发现 LM_BC_FR 平台SKU为空】{Color.RESET}")
                print(f"{Color.YELLOW}  - 脚本已自动尝试 VLOOKUP 回填（从手动-二次映射.xlsx）{Color.RESET}")
                print(f"{Color.YELLOW}  - 如仍有空值，需手动在 Excel 中使用 VLOOKUP 补充{Color.RESET}")
                print(f"    SKU-站点识别码: =VLOOKUP(A列,[手动-二次映射.xlsx]二次上架-LM-BC-自发货!$A:$I,9,FALSE)")
                print(f"    站点: =VLOOKUP(A列,[手动-二次映射.xlsx]二次上架-LM-BC-自发货!$A:$I,7,FALSE)")
                print(f"\n{Color.YELLOW}【操作】{Color.RESET}")
                print(f"{Color.YELLOW}  - 确认数据完整后，按回车键继续{Color.RESET}")
                print(f"{Color.YELLOW}{'=' * 80}{Color.RESET}\n")

                if args.stop_after_g3:
                    print("[runAll_G] 已执行到 G3，按参数要求停止（请先检查映射和计算结果）")
                    return 0
                if args.pause_after_g3:
                    input("[runAll_G] >>> 按回车键继续执行后续步骤...")
                else:
                    # 默认行为：暂停等待用户确认
                    input("[runAll_G] >>> 按回车键继续执行后续步骤...")
            continue

        print(f"{Color.RED}[runAll_G] ({idx}/{len(scripts)}) 失败：{script.name}（exit={code}）{Color.RESET}")
        failed.append((script.name, code))
        if not args.continue_on_error:
            break

    if failed:
        print(f"\n{Color.RED}[runAll_G] 执行结束：存在失败步骤：{Color.RESET}")
        for name, code in failed:
            print(f"  {Color.RED}- {name}（exit={code}）{Color.RESET}")
        return 1

    print(f"\n{Color.GREEN}[runAll_G] 执行结束：全部成功{Color.RESET}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
