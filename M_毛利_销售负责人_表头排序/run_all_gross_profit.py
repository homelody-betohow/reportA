"""
按顺序执行 M_毛利_销售负责人_表头排序 目录下的 M1 → M4 全流程脚本。

执行顺序：
  M1 计算毛利 → M2 映射销售负责人(非AMZ)
  → M3 映射销售负责人(AMZ) → M4 映射销售经理表头排序

特点：
  - 全自动流程，无需手动干预
  - 计算毛利并映射销售负责人信息
  - 按销售经理对表头进行排序

用法：
  python "A_报表/M_毛利_销售负责人_表头排序/run_all_gross_profit.py"
"""
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

from A_报表.Z_method.style import Color
from A_报表.Z_method.runall_utils import (
    setup_console_encoding,
    run_script,
)

# 设置 Windows 控制台编码
setup_console_encoding()

# 执行步骤必须按依赖顺序
_PIPELINE: tuple[str, ...] = (
    "M1_计算毛利.py",
    "M2_映射_销售负责人_非AMZ.py",
    "M3_映射_销售负责人_AMZ.py",
    "M4_映射_销售经理_表头排序.py",
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="按顺序执行毛利计算和销售负责人映射脚本（M1 → M4）"
    )
    parser.add_argument(
        "--continue-on-error",
        action="store_true",
        help="某一步失败后继续执行后续脚本（默认：失败即停止）",
    )
    args, unknown = parser.parse_known_args(argv)
    if unknown:
        print(f"[run_all_gross_profit] 忽略多余参数：{unknown}")

    folder = Path(__file__).resolve().parent
    
    # 验证所有脚本是否存在
    scripts = []
    for name in _PIPELINE:
        script_path = folder / name
        if not script_path.is_file():
            print(f"{Color.RED}[run_all_gross_profit] 错误：找不到脚本 {name}{Color.RESET}")
            return 1
        scripts.append(script_path)

    print(f"[run_all_gross_profit] 将执行 {len(scripts)} 个脚本，目录：{folder}")

    failed: list[tuple[str, int]] = []

    for idx, script in enumerate(scripts, start=1):
        print(f"[run_all_gross_profit] ({idx}/{len(scripts)}) 开始：{Color.YELLOW} {script.name} {Color.RESET}")
        code, output = run_script(script)
        
        if code == 0:
            print(f"[run_all_gross_profit] ({idx}/{len(scripts)}) 完成：{script.name}")
            continue

        print(f"{Color.RED}[run_all_gross_profit] ({idx}/{len(scripts)}) 失败：{script.name}（exit={code}）{Color.RESET}")
        failed.append((script.name, code))
        
        if not args.continue_on_error:
            break

    if failed:
        print(f"\n{Color.RED}[run_all_gross_profit] 执行结束：存在失败步骤：{Color.RESET}")
        for name, code in failed:
            print(f"  {Color.RED}- {name}（exit={code}）{Color.RESET}")
        return 1

    print(f"\n{Color.GREEN}[run_all_gross_profit] 执行结束：全部成功{Color.RESET}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
