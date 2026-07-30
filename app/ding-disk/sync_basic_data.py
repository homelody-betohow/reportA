"""
按顺序从钉钉在线表格拉取基础数据并回写数据库。

执行顺序：
  platform_shop_pull → warehouse_pull → product_sku_pull

默认某步失败会跳过并继续下一步；可用 ``--stop-on-error`` 改为失败即停止。

用法（项目根目录）::

    python app/ding-disk/sync_basic_data.py
    python app/ding-disk/sync_basic_data.py --stop-on-error
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
assert _spec.loader is not None
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from common.style import Color  # noqa: E402
from common.runall_utils import (  # noqa: E402
    setup_console_encoding,
    run_script,
)

setup_console_encoding()

_PIPELINE: tuple[str, ...] = (
    "platform_shop_pull.py",
    "warehouse_pull.py",
    "product_sku_pull.py",
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="按顺序拉取钉钉基础数据：店铺 → 仓库 → 产品SKU",
    )
    parser.add_argument(
        "--stop-on-error",
        action="store_true",
        help="某一步失败后立即停止（默认：跳过失败步骤，继续执行下一步）",
    )
    args = parser.parse_args(argv)

    folder = Path(__file__).resolve().parent

    scripts: list[Path] = []
    for name in _PIPELINE:
        script_path = folder / name
        if not script_path.is_file():
            print(f"{Color.RED}[sync_basic_data] 错误：找不到脚本 {name}{Color.RESET}")
            return 1
        scripts.append(script_path)

    print(f"[sync_basic_data] 将执行 {len(scripts)} 个脚本，目录：{folder}")

    failed: list[tuple[str, int]] = []

    for idx, script in enumerate(scripts, start=1):
        print(
            f"[sync_basic_data] ({idx}/{len(scripts)}) 开始："
            f"{Color.YELLOW} {script.name} {Color.RESET}"
        )
        code, _output = run_script(script)

        if code == 0:
            print(f"[sync_basic_data] ({idx}/{len(scripts)}) 完成：{script.name}")
            continue

        print(
            f"{Color.RED}[sync_basic_data] ({idx}/{len(scripts)}) "
            f"失败，跳过：{script.name}（exit={code}）{Color.RESET}"
        )
        failed.append((script.name, code))

        if args.stop_on_error:
            break

    if failed:
        print(f"\n{Color.RED}[sync_basic_data] 执行结束：存在失败步骤：{Color.RESET}")
        for name, code in failed:
            print(f"  {Color.RED}- {name}（exit={code}）{Color.RESET}")
        return 1

    print(f"\n{Color.GREEN}[sync_basic_data] 执行结束：全部成功{Color.RESET}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
