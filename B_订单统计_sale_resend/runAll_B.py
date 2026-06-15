from __future__ import annotations

import argparse
import importlib.util
import locale
import os
import re
import subprocess
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

from A_报表.Z_method.style import Color
from A_报表.Z_method.runall_utils import (
    setup_console_encoding,
    run_script,
    get_scripts_in_order,
    extract_output_file_path,
)

# 设置 Windows 控制台编码
setup_console_encoding()


def _is_b1_1(script_path: Path) -> bool:
    # 判断是否是 B1_1 这一步（输出需要手动查询的订单号）
    return script_path.name.startswith("B1_1_")


def _is_b5(script_path: Path) -> bool:
    # 判断是否是 B5 这一步（映射MF尾程，需要检查是否有空值）
    return script_path.name.startswith("B5_")


def _is_b6(script_path: Path) -> bool:
    # 判断是否是 B6 这一步（映射非MF尾程，需要检查映射结果）
    return script_path.name.startswith("B6_")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="按顺序执行本目录下所有 B*.py 脚本（排除 runAll_B.py）"
    )
    parser.add_argument(
        "--continue-on-error",
        action="store_true",
        help="某一步失败后继续执行后续脚本（默认：失败即停止）",
    )
    parser.add_argument(
        "--stop-after-b1-1",
        action="store_true",
        help="执行到 B1_1_*.py 完成后立即停止（用于你手动去 ERP 查询订单）",
    )
    parser.add_argument(
        "--pause-after-b1-1",
        action="store_true",
        help="执行到 B1_1_*.py 完成后暂停，按回车后继续执行后续脚本",
    )
    parser.add_argument(
        "--stop-after-b5",
        action="store_true",
        help="执行到 B5_*.py 完成后立即停止（用于你手动检查 MF-派送费 是否有空值）",
    )
    parser.add_argument(
        "--pause-after-b5",
        action="store_true",
        help="执行到 B5_*.py 完成后暂停，按回车后继续执行后续脚本",
    )
    parser.add_argument(
        "--stop-after-b6",
        action="store_true",
        help="执行到 B6_*.py 完成后立即停止（用于你手动检查 非MF-尾程 映射是否有空值）",
    )
    parser.add_argument(
        "--pause-after-b6",
        action="store_true",
        help="执行到 B6_*.py 完成后暂停，按回车后继续执行后续脚本",
    )
    args, unknown = parser.parse_known_args(argv)
    if unknown:
        print(f"[runAll_B] 忽略多余参数：{unknown}")

    folder = Path(__file__).resolve().parent
    scripts = get_scripts_in_order(folder, Path(__file__).name)

    if not scripts:
        print(f"[runAll_B] 未找到可执行脚本：{folder}")
        return 0

    print(f"[runAll_B] 将执行 {len(scripts)} 个脚本，目录：{folder}")

    failed: list[tuple[str, int]] = []

    for idx, script in enumerate(scripts, start=1):
        print(f"[runAll_B] ({idx}/{len(scripts)}) 开始：{Color.YELLOW} {script.name} {Color.RESET}")
        code, output = run_script(script)
        if code == 0:
            print(f"[runAll_B] ({idx}/{len(scripts)}) 完成：{script.name}")

            if _is_b1_1(script):
                print(f"\n{Color.YELLOW}{'=' * 80}{Color.RESET}")
                print(f"{Color.YELLOW}[runAll_B] 重要提示：B1_1 已完成，请按照上面的提示操作：{Color.RESET}")
                print(f"{Color.YELLOW}{'=' * 80}{Color.RESET}")
                print(f"{Color.YELLOW}【需要保存 ERP 查询结果的位置】{Color.RESET}")
                print(f"{Color.YELLOW}  - REAL-FB 无站点订单：{Color.RESET}")
                print(f"    \\\\Betohow\\数据报表\\RPA\\报表-无站点-订单查询\\REAL-FB")
                print(f"{Color.YELLOW}  - LM-BC 重发订单：{Color.RESET}")
                print(f"    \\\\Betohow\\数据报表\\RPA\\报表-无站点-订单查询\\LM-BC-重发")
                print(f"{Color.YELLOW}  - LM-RP 重发订单：{Color.RESET}")
                print(f"    \\\\Betohow\\数据报表\\RPA\\报表-无站点-订单查询\\LM-RP-重发")
                print(f"\n{Color.YELLOW}【操作步骤】{Color.RESET}")
                print(f"{Color.YELLOW}  1. 复制上面黄色的订单号列表（Ctrl+C）{Color.RESET}")
                print(f"{Color.YELLOW}  2. 到 ERP 中批量查询这些订单{Color.RESET}")
                print(f"{Color.YELLOW}  3. 将查询结果保存到上述对应的文件夹中{Color.RESET}")
                print(f"{Color.YELLOW}  4. 确认文件保存成功后，按回车键继续{Color.RESET}")
                print(f"{Color.YELLOW}{'=' * 80}{Color.RESET}\n")

                if args.stop_after_b1_1:
                    print("[runAll_B] 已执行到 B1_1，按参数要求停止（请先手动查询 ERP 并保存结果）")
                    return 0
                if args.pause_after_b1_1:
                    input("[runAll_B] >>> 按回车键继续执行后续步骤...")
                else:
                    # 默认行为：暂停等待用户确认
                    input("[runAll_B] >>> 按回车键继续执行后续步骤...")
            
            if _is_b5(script):
                # 从输出中提取文件路径
                output_file = extract_output_file_path(output, "(已完成-5)订单统计-*.xlsx")
                
                print(f"\n{Color.YELLOW}{'=' * 80}{Color.RESET}")
                print(f"{Color.YELLOW}[runAll_B] 重要提示：B5 已完成（MF 站点尾程映射）{Color.RESET}")
                print(f"{Color.YELLOW}{'=' * 80}{Color.RESET}")
                print(f"{Color.YELLOW}【需要检查的文件】{Color.RESET}")
                print(f"  {output_file}")
                print(f"\n{Color.YELLOW}【检查内容】{Color.RESET}")
                print(f"{Color.YELLOW}  1. 打开上述 Excel 文件{Color.RESET}")
                print(f"{Color.YELLOW}  2. 筛选「派送费-映射分类」列，找到包含 'MF' 的所有行{Color.RESET}")
                print(f"{Color.YELLOW}  3. 检查这些行的「MF-派送费」列是否有空值{Color.RESET}")
                print(f"{Color.YELLOW}  4. 注意查看「仓库SKU销量」，确认数量合理{Color.RESET}")
                print(f"\n{Color.YELLOW}【如果发现空值】{Color.RESET}")
                print(f"{Color.YELLOW}  - 找王园芳补充基础表：桌面\\MANO-MF 尾程.xlsx{Color.RESET}")
                print(f"{Color.YELLOW}  - COMMF 和 OHPAMF 是一样的，需要同时补充{Color.RESET}")
                print(f"\n{Color.YELLOW}【操作】{Color.RESET}")
                print(f"{Color.YELLOW}  - 确认数据无误后，按回车键继续{Color.RESET}")
                print(f"{Color.YELLOW}{'=' * 80}{Color.RESET}\n")

                if args.stop_after_b5:
                    print("[runAll_B] 已执行到 B5，按参数要求停止（请先手动检查 MF-派送费）")
                    return 0
                if args.pause_after_b5:
                    input("[runAll_B] >>> 按回车键继续执行后续步骤...")
                else:
                    # 默认行为：暂停等待用户确认
                    input("[runAll_B] >>> 按回车键继续执行后续步骤...")

            if _is_b6(script):
                # 从输出中提取文件路径
                output_file = extract_output_file_path(output, "(已完成-6)订单统计-*.xlsx")
                
                print(f"\n{Color.YELLOW}{'=' * 80}{Color.RESET}")
                print(f"{Color.YELLOW}[runAll_B] 重要提示：B6 已完成（非MF 站点尾程映射）{Color.RESET}")
                print(f"{Color.YELLOW}{'=' * 80}{Color.RESET}")
                print(f"{Color.YELLOW}【需要检查的文件】{Color.RESET}")
                print(f"  {output_file}")
                print(f"\n{Color.YELLOW}【检查内容】{Color.RESET}")
                print(f"{Color.YELLOW}  1. 打开上述 Excel 文件{Color.RESET}")
                print(f"{Color.YELLOW}  2. 检查非MF站点的尾程映射结果{Color.RESET}")
                print(f"{Color.YELLOW}  3. 查看相关列是否有空值或异常值{Color.RESET}")
                print(f"{Color.YELLOW}  4. 重点检查「派送费」相关字段{Color.RESET}")
                print(f"\n{Color.YELLOW}【如果发现空值】{Color.RESET}")
                print(f"{Color.YELLOW}  - 需要补充对应的映射表{Color.RESET}")
                print(f"{Color.YELLOW}  - 补充后重新运行 B6 脚本{Color.RESET}")
                print(f"\n{Color.YELLOW}【操作】{Color.RESET}")
                print(f"{Color.YELLOW}  - 确认数据无误后，按回车键继续{Color.RESET}")
                print(f"{Color.YELLOW}{'=' * 80}{Color.RESET}\n")

                if args.stop_after_b6:
                    print("[runAll_B] 已执行到 B6，按参数要求停止（请先手动检查 非MF-尾程 映射结果）")
                    return 0
                if args.pause_after_b6:
                    input("[runAll_B] >>> 按回车键继续执行后续步骤...")
                else:
                    # 默认行为：暂停等待用户确认
                    input("[runAll_B] >>> 按回车键继续执行后续步骤...")
            continue

        print(f"[runAll_B] ({idx}/{len(scripts)}) 失败：{script.name}（exit={code}）")
        failed.append((script.name, code))
        if not args.continue_on_error:
            break

    if failed:
        print("[runAll_B] 执行结束：存在失败步骤：")
        for name, code in failed:
            print(f"  - {name}（exit={code}）")
        return 1

    print("[runAll_B] 执行结束：全部成功")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
