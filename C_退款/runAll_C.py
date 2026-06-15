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


def _is_c1_1(script_path: Path) -> bool:
    # 判断是否是 C1_1 这一步（输出需要手动查询的订单号）
    return script_path.name.startswith("C1_1_")




def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="按顺序执行本目录下所有 C*.py 脚本（排除 runAll_C.py）"
    )
    parser.add_argument(
        "--continue-on-error",
        action="store_true",
        help="某一步失败后继续执行后续脚本（默认：失败即停止）",
    )
    parser.add_argument(
        "--stop-after-c1-1",
        action="store_true",
        help="执行到 C1_1_*.py 完成后立即停止（用于你手动去 ERP 查询订单）",
    )
    parser.add_argument(
        "--pause-after-c1-1",
        action="store_true",
        help="执行到 C1_1_*.py 完成后暂停，按回车后继续执行后续脚本",
    )
    args, unknown = parser.parse_known_args(argv)
    if unknown:
        print(f"[runAll_C] 忽略多余参数：{unknown}")

    folder = Path(__file__).resolve().parent
    scripts = get_scripts_in_order(folder, Path(__file__).name)

    if not scripts:
        print(f"[runAll_C] 未找到可执行脚本：{folder}")
        return 0

    print(f"[runAll_C] 将执行 {len(scripts)} 个脚本，目录：{folder}")

    failed: list[tuple[str, int]] = []

    for idx, script in enumerate(scripts, start=1):
        print(f"[runAll_C] ({idx}/{len(scripts)}) 开始：{Color.YELLOW} {script.name} {Color.RESET}")
        code, output = run_script(script)
        if code == 0:
            print(f"[runAll_C] ({idx}/{len(scripts)}) 完成：{script.name}")

            if _is_c1_1(script):
                print(f"\n{Color.YELLOW}{'=' * 80}{Color.RESET}")
                print(f"{Color.YELLOW}[runAll_C] 重要提示：C1_1 已完成，请按照上面的提示操作：{Color.RESET}")
                print(f"{Color.YELLOW}{'=' * 80}{Color.RESET}")
                print(f"{Color.YELLOW}【需要检查的位置】{Color.RESET}")
                print(f"{Color.YELLOW}  - LM-BC 退款订单查询结果保存到：{Color.RESET}")
                print(f"    \\\\Betohow\\数据报表\\RPA\\报表-无站点-订单查询\\LM-BC-退款")
                print(f"{Color.YELLOW}  - LM-RP 退款订单查询结果保存到：{Color.RESET}")
                print(f"    \\\\Betohow\\数据报表\\RPA\\报表-无站点-订单查询\\LM-RP-退款")
                print(f"\n{Color.YELLOW}【操作步骤】{Color.RESET}")
                print(f"{Color.YELLOW}  1. 复制上面黄色的订单号列表（Ctrl+C）{Color.RESET}")
                print(f"{Color.YELLOW}  2. 到 ERP 中批量查询这些退款订单{Color.RESET}")
                print(f"{Color.YELLOW}  3. 将查询结果保存到上述对应的文件夹中{Color.RESET}")
                print(f"{Color.YELLOW}  4. 确认文件保存成功后，按回车键继续{Color.RESET}")
                print(f"{Color.YELLOW}{'=' * 80}{Color.RESET}\n")

                if args.stop_after_c1_1:
                    print("[runAll_C] 已执行到 C1_1，按参数要求停止（请先手动查询 ERP 并保存结果）")
                    return 0
                if args.pause_after_c1_1:
                    input("[runAll_C] >>> 按回车键继续执行后续步骤...")
                else:
                    # 默认行为：暂停等待用户确认
                    input("[runAll_C] >>> 按回车键继续执行后续步骤...")
            
            continue

        print(f"[runAll_C] ({idx}/{len(scripts)}) 失败：{script.name}（exit={code}）")
        failed.append((script.name, code))
        if not args.continue_on_error:
            break

    if failed:
        print("[runAll_C] 执行结束：存在失败步骤：")
        for name, code in failed:
            print(f"  - {name}（exit={code}）")
        return 1

    print("[runAll_C] 执行结束：全部成功")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
