from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def _scripts_in_order(folder: Path) -> list[Path]:
    # 约定：按文件名排序即可得到步骤顺序（A1 -> A2 -> A3 ...）
    # 排除自身 runAll_A.py
    scripts = [
        p
        for p in folder.glob("*.py")
        if p.is_file() and p.name != Path(__file__).name
    ]
    scripts.sort(key=lambda p: p.name)
    return scripts


def _run_one(script_path: Path) -> int:
    # 用当前 Python 解释器去跑子脚本，确保环境一致
    proc = subprocess.run(
        [sys.executable, str(script_path)],
        cwd=str(script_path.parent),
        check=False,
    )
    return int(proc.returncode)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="按顺序执行本目录下所有 A*.py 脚本（排除 runAll_A.py）"
    )
    parser.add_argument(
        "--continue-on-error",
        action="store_true",
        help="某一步失败后继续执行后续脚本（默认：失败即停止）",
    )
    args = parser.parse_args(argv)

    folder = Path(__file__).resolve().parent
    scripts = _scripts_in_order(folder)

    if not scripts:
        print(f"[runAll_A] 未找到可执行脚本：{folder}")
        return 0

    print(f"[runAll_A] 将执行 {len(scripts)} 个脚本，目录：{folder}")

    failed: list[tuple[str, int]] = []

    for idx, script in enumerate(scripts, start=1):
        print(f"[runAll_A] ({idx}/{len(scripts)}) 开始：{script.name}")
        code = _run_one(script)
        if code == 0:
            print(f"[runAll_A] ({idx}/{len(scripts)}) 完成：{script.name}")
            continue

        print(f"[runAll_A] ({idx}/{len(scripts)}) 失败：{script.name}（exit={code}）")
        failed.append((script.name, code))
        if not args.continue_on_error:
            break

    if failed:
        print("[runAll_A] 执行结束：存在失败步骤：")
        for name, code in failed:
            print(f"  - {name}（exit={code}）")
        return 1

    print("[runAll_A] 执行结束：全部成功")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
