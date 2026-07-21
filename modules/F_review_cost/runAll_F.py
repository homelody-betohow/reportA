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
)

# 设置 Windows 控制台编码
setup_console_encoding()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="按顺序执行本目录下所有 F*.py 脚本（排除 runAll_F.py）"
    )
    parser.add_argument(
        "--continue-on-error",
        action="store_true",
        help="某一步失败后继续执行后续脚本（默认：失败即停止）",
    )
    args, unknown = parser.parse_known_args(argv)
    if unknown:
        print(f"[runAll_F] 忽略多余参数：{unknown}")

    folder = Path(__file__).resolve().parent
    scripts = get_scripts_in_order(folder, Path(__file__).name)

    if not scripts:
        print(f"[runAll_F] 未找到可执行脚本：{folder}")
        return 0

    print(f"[runAll_F] 将执行 {len(scripts)} 个脚本，目录：{folder}")

    failed: list[tuple[str, int]] = []

    for idx, script in enumerate(scripts, start=1):
        print(f"[runAll_F] ({idx}/{len(scripts)}) 开始：{Color.YELLOW} {script.name} {Color.RESET}")
        code, output = run_script(script)
        if code == 0:
            print(f"[runAll_F] ({idx}/{len(scripts)}) 完成：{script.name}")
            continue

        print(f"{Color.RED}[runAll_F] ({idx}/{len(scripts)}) 失败：{script.name}（exit={code}）{Color.RESET}")
        failed.append((script.name, code))
        if not args.continue_on_error:
            break

    if failed:
        print(f"\n{Color.RED}[runAll_F] 执行结束：存在失败步骤：{Color.RESET}")
        for name, code in failed:
            print(f"  {Color.RED}- {name}（exit={code}）{Color.RESET}")
        return 1

    print(f"\n{Color.GREEN}[runAll_F] 执行结束：全部成功{Color.RESET}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
