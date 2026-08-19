"""
按顺序执行 D_ads 目录下的 D1 → D6 全流程脚本。

执行顺序：
  D1 OTTO → D2 REAL → D3 MANO
  → D4_1 shopping 美元合并 → D4_2 创建分摊表 → D4_3 分摊美元 → D4_4 合并欧元
  → D0 Cdiscount（fee_advertising platform=cdiscount）
  → D5 合并各平台广告 → D6 合并订单统计

前置条件（由其它流程生成，本脚本不检查）：
  - D4_3 需要：订单统计目录下「(已完成-6)订单统计-{日期}.xlsx」
  - D6 需要：订单统计目录下「(已完成-8)订单统计-{日期}.xlsx」

用法：
  python "modules/D_ads/runAll_D.py"
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

from common.style import Color
from common.runall_utils import (
    setup_console_encoding,
    run_script,
)

# 设置 Windows 控制台编码
setup_console_encoding()

# D4 子步骤必须按依赖顺序；D4_4 文件名在仓库中为双后缀 .py.py
_PIPELINE: tuple[str, ...] = (
    "D0_Cdiscount.py",
    "D1_OTTO.py",
    "D2_REAL.py",
    "D3_MANO.py",
    "D4_1_DLZ_shopping_美元_合并_儿子-站点识别码.py",
    "D4_2_DLZ_创建分摊表.py",
    "D4_3_DLZ_分摊_美元.py",
    "D4_4_DLZ_shopping_分摊_合并_欧元.py.py",
    "D5_合并_儿子-站点识别码_所有平台广告.py",
    "D6_合并_订单统计.py",
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="按顺序执行 D_ads 目录下所有步骤脚本（D1 → D6）"
    )
    parser.add_argument(
        "--continue-on-error",
        action="store_true",
        help="某一步失败后继续执行后续脚本（默认：失败即停止）",
    )
    args, unknown = parser.parse_known_args(argv)
    if unknown:
        print(f"[runAll_D] 忽略多余参数：{unknown}")

    folder = Path(__file__).resolve().parent
    
    # 验证所有脚本是否存在
    scripts = []
    for name in _PIPELINE:
        script_path = folder / name
        if not script_path.is_file():
            print(f"{Color.RED}[runAll_D] 错误：找不到脚本 {name}{Color.RESET}")
            return 1
        scripts.append(script_path)

    print(f"[runAll_D] 将执行 {len(scripts)} 个脚本，目录：{folder}")
    print(f"[runAll_D] 前置条件：")
    print(f"  - 确保订单统计目录下有「(已完成-6)订单统计.xlsx」（D4_3 需要）")
    print(f"  - 确保订单统计目录下有「(已完成-8)订单统计.xlsx」（D6 需要）")
    print()

    failed: list[tuple[str, int]] = []

    for idx, script in enumerate(scripts, start=1):
        print(f"[runAll_D] ({idx}/{len(scripts)}) 开始：{Color.YELLOW} {script.name} {Color.RESET}")
        code, output = run_script(script)
        
        if code == 0:
            print(f"[runAll_D] ({idx}/{len(scripts)}) 完成：{script.name}")
            continue

        print(f"{Color.RED}[runAll_D] ({idx}/{len(scripts)}) 失败：{script.name}（exit={code}）{Color.RESET}")
        failed.append((script.name, code))
        
        if not args.continue_on_error:
            break

    if failed:
        print(f"\n{Color.RED}[runAll_D] 执行结束：存在失败步骤：{Color.RESET}")
        for name, code in failed:
            print(f"  {Color.RED}- {name}（exit={code}）{Color.RESET}")
        return 1

    print(f"\n{Color.GREEN}[runAll_D] 执行结束：全部成功{Color.RESET}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
