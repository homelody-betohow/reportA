"""
按推荐顺序执行报表流水线（A→M）。

用法（在项目根目录）:
  python -m app.run_pipeline
  python -m app.run_pipeline --only A,B,C
  python app/run_pipeline.py
"""
from __future__ import annotations

import argparse
import importlib.util
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

# 将项目根加入 sys.path
_epr = ROOT / "ensure_project_root.py"
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr)
_mod = importlib.util.module_from_spec(_spec)
assert _spec.loader is not None
_spec.loader.exec_module(_mod)
_mod.bootstrap(__file__)

from common.runall_utils import run_script, setup_console_encoding  # noqa: E402
from common.style import Color  # noqa: E402

# (短名, 相对 modules 的入口脚本)
PIPELINE: list[tuple[str, str]] = [
    ("A", "A_temu_order_amount/runAll_A.py"),
    ("B", "B_order_stats_sale_resend/runAll_B.py"),
    ("C", "C_refund/runAll_C.py"),
    ("D", "D_ads/runAll_D.py"),
    ("F", "F_review_cost/runAll_F.py"),
    ("G", "G_returned_reshelf/runAll_G.py"),
    ("H", "H_amz_profit_otto_manager_fee/runAll_H.py"),
    ("K", "K_storage_fee_product_info/runAll_K.py"),
    ("M", "M_gross_profit_owner_headers/run_all_gross_profit.py"),
]


def main(argv: list[str] | None = None) -> int:
    setup_console_encoding()
    parser = argparse.ArgumentParser(description="执行报表流水线")
    parser.add_argument(
        "--only",
        help="只跑指定模块，逗号分隔，如 A,B,C（默认全部有 runAll 的模块）",
        default="",
    )
    args = parser.parse_args(argv)

    only = {x.strip().upper() for x in args.only.split(",") if x.strip()}
    steps = [(k, p) for k, p in PIPELINE if not only or k in only]

    if not steps:
        print(Color.RED + f"无匹配模块: {args.only!r}" + Color.RESET)
        return 2

    modules_root = ROOT / "modules"
    failed: list[str] = []
    for key, rel in steps:
        script = modules_root / rel
        print(Color.CYAN + f"\n===== [{key}] {rel} =====" + Color.RESET)
        if not script.is_file():
            print(Color.RED + f"脚本不存在: {script}" + Color.RESET)
            failed.append(key)
            continue
        code, _ = run_script(script)
        if code != 0:
            failed.append(key)
            print(Color.RED + f"[{key}] 失败 exit={code}" + Color.RESET)
        else:
            print(Color.GREEN + f"[{key}] 完成" + Color.RESET)

    if failed:
        print(Color.RED + f"\n失败模块: {', '.join(failed)}" + Color.RESET)
        return 1
    print(Color.GREEN + "\n全部完成" + Color.RESET)
    return 0


if __name__ == "__main__":
    sys.exit(main())
