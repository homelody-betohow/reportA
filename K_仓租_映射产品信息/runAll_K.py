"""
按顺序执行 K_仓租_映射产品信息 目录下的 K1 → K6 全流程脚本。

执行顺序：
  K1 HY仓租 → K2 4PX仓租 → K3 合并仓租表站点商品ID识别码
  → K4 合并分摊订单统计 → K5 映射产品信息（手动检查） → K6 处理分销分摊仓租

特别说明：
  - K5 执行完成后会自动打开文件/文件夹，方便你检查 MANO-MF 仓租映射结果
  - 默认行为：K5 完成后自动打开文件和文件夹，等待你按回车继续

用法：
  python "A_报表/K_仓租_映射产品信息/runAll_K.py"
  
  # 执行到 K5 后立即停止（手动检查后再运行后续步骤）
  python "A_报表/K_仓租_映射产品信息/runAll_K.py" --stop-after-k5
  
  # 不自动打开文件
  python "A_报表/K_仓租_映射产品信息/runAll_K.py" --open-after-k5=none
"""
from __future__ import annotations

import argparse
import importlib.util
import os
import re
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
)

# 设置 Windows 控制台编码
setup_console_encoding()


def _is_k5(script_path: Path) -> bool:
    return script_path.name.startswith("K5_")


_SAVED_AS_RE = re.compile(r"文件另存为[：:]\s*(?P<path>.+?\.xlsx)\s*$")


def _extract_saved_excel_path(output: str) -> Path | None:
    for raw_line in reversed(output.splitlines()):
        line = raw_line.strip()
        m = _SAVED_AS_RE.search(line)
        if m:
            p = m.group("path").strip().strip('"').strip("'")
            try:
                return Path(p)
            except Exception:
                return None
    return None


def _open_after_k5(target: Path | None, mode: str) -> None:
    """
    K5 执行后自动打开文件/文件夹，强提醒用户检查。
    mode: none | folder | file | both
    """
    if mode == "none":
        return
    if target is None:
        return

    try:
        if mode in ("folder", "both"):
            os.startfile(str(target.parent))
        if mode in ("file", "both") and target.is_file():
            os.startfile(str(target))
    except Exception as e:
        print(f"{Color.RED}[runAll_K] 打开检查文件/文件夹失败：{e}{Color.RESET}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="按顺序执行本目录下所有 K*.py 脚本（排除 runAll_K.py）"
    )
    parser.add_argument(
        "--continue-on-error",
        action="store_true",
        help="某一步失败后继续执行后续脚本（默认：失败即停止）",
    )
    parser.add_argument(
        "--stop-after-k5",
        action="store_true",
        help="执行到 K5_*.py 完成后立即停止（用于你手动检查映射结果）",
    )
    parser.add_argument(
        "--pause-after-k5",
        action="store_true",
        help="执行到 K5_*.py 完成后暂停，按回车后继续执行后续脚本",
    )
    parser.add_argument(
        "--open-after-k5",
        choices=["none", "folder", "file", "both"],
        default="both",
        help="执行完 K5 后自动打开（强提醒检查）：none=不打开 folder=打开文件夹 file=打开文件 both=都打开",
    )
    args, unknown = parser.parse_known_args(argv)
    if unknown:
        print(f"[runAll_K] 忽略多余参数：{unknown}")

    folder = Path(__file__).resolve().parent
    scripts = get_scripts_in_order(folder, Path(__file__).name)

    if not scripts:
        print(f"[runAll_K] 未找到可执行脚本：{folder}")
        return 0

    print(f"[runAll_K] 将执行 {len(scripts)} 个脚本，目录：{folder}")

    failed: list[tuple[str, int]] = []

    for idx, script in enumerate(scripts, start=1):
        print(f"[runAll_K] ({idx}/{len(scripts)}) 开始：{Color.YELLOW} {script.name} {Color.RESET}")
        code, output = run_script(script)
        if code == 0:
            print(f"[runAll_K] ({idx}/{len(scripts)}) 完成：{script.name}")

            if _is_k5(script):
                saved_excel = _extract_saved_excel_path(output)
                
                print(f"\n{Color.YELLOW}{'=' * 80}{Color.RESET}")
                print(f"{Color.YELLOW}[runAll_K] 重要提示：K5 已完成（产品信息映射）{Color.RESET}")
                print(f"{Color.YELLOW}{'=' * 80}{Color.RESET}")
                
                if saved_excel is not None:
                    print(f"{Color.YELLOW}【生成的文件】{Color.RESET}")
                    print(f"  {saved_excel}")
                    print(f"\n{Color.YELLOW}【检查内容】{Color.RESET}")
                    print(f"{Color.YELLOW}  1. 检查 MANO-MF 仓租映射是否完整{Color.RESET}")
                    print(f"{Color.YELLOW}  2. 查看产品信息字段是否有空值或异常值{Color.RESET}")
                    print(f"{Color.YELLOW}  3. 核对仓租金额计算是否正确{Color.RESET}")
                    print(f"\n{Color.YELLOW}【如果发现空值】{Color.RESET}")
                    print(f"{Color.YELLOW}  - 需要补充对应的产品信息映射表{Color.RESET}")
                    print(f"{Color.YELLOW}  - 补充后重新运行 K5 脚本{Color.RESET}")
                    print(f"\n{Color.YELLOW}【操作】{Color.RESET}")
                    print(f"{Color.YELLOW}  - 系统将自动打开文件供你检查{Color.RESET}")
                    print(f"{Color.YELLOW}  - 确认数据无误后，按回车键继续{Color.RESET}")
                    print(f"{Color.YELLOW}{'=' * 80}{Color.RESET}\n")
                    
                    _open_after_k5(saved_excel, args.open_after_k5)
                else:
                    print(f"{Color.YELLOW}[runAll_K] 未能从 K5 输出中识别保存的 Excel 路径（无法自动打开）{Color.RESET}")

                if args.stop_after_k5:
                    print(f"{Color.CYAN}[runAll_K] 已执行到 K5，按参数要求停止（请先手动检查映射字段是否齐全）{Color.RESET}")
                    return 0
                if args.pause_after_k5:
                    input("[runAll_K] >>> 按回车键继续执行后续步骤...")
                else:
                    # 默认行为：暂停等待用户确认
                    input("[runAll_K] >>> 按回车键继续执行后续步骤...")
            continue

        print(f"{Color.RED}[runAll_K] ({idx}/{len(scripts)}) 失败：{script.name}（exit={code}）{Color.RESET}")
        failed.append((script.name, code))
        if not args.continue_on_error:
            break

    if failed:
        print(f"\n{Color.RED}[runAll_K] 执行结束：存在失败步骤：{Color.RESET}")
        for name, code in failed:
            print(f"  {Color.RED}- {name}（exit={code}）{Color.RESET}")
        return 1

    print(f"\n{Color.GREEN}[runAll_K] 执行结束：全部成功{Color.RESET}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
