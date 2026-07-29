"""按顺序执行退件四步：create_return → update_return → download_label → refresh_return。

需要 ``runtime_config``：统一 UTF-8 控制台、项目根 / ``config/*.json`` 路径
（与单步脚本相同；本文件只做编排，业务配置仍由各步自行加载）。

控制台输出同时写入日志（冻结：``dist/logs/returned_yyyymmdd.log``）。

每步可单独配置「执行前等待秒数」（上一成功结束后再等）。优先级::

    CLI ``--delay STEP=SEC``  >  ``returned_config.json`` 的 ``step_delays``
    > 内置默认（见 DEFAULT_STEP_DELAYS）

用法（项目根目录）::

    python app/hongyu/returned/run_task.py
    python app/hongyu/returned/run_task.py --delay update_return=35 --delay download_label=10
    python app/hongyu/returned/run_task.py --dry-run
    python app/hongyu/returned/run_task.py --continue-on-error
    python app/hongyu/returned/run_task.py --skip download_label
    python app/hongyu/returned/run_task.py --only refresh_return
    python app/hongyu/returned/run_task.py --no-log

Windows 可执行文件::

    .\\dist\\returned\\run_task.exe
    .\\dist\\returned\\run_task.exe --delay update_return=40 --dry-run
    # 配置：dist\\config\\returned_config.json → step_delays
    # 日志：dist\\logs\\returned_yyyymmdd.log
"""

from __future__ import annotations

import argparse
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Mapping, Optional, Sequence, Tuple

_RETURNED_DIR = Path(__file__).resolve().parent
if str(_RETURNED_DIR) not in sys.path:
    sys.path.insert(0, str(_RETURNED_DIR))

import runtime_config as _returned_rt  # noqa: E402

_PROJECT_ROOT, _CFG = _returned_rt.init_script(__file__)

from common.dist_paths import daily_log_path  # noqa: E402
from tee_io import TeeTextIO  # noqa: E402

# 各步脚本在同目录；导入时会再 init_script / 读 returned_config
import create_return as _step_create  # noqa: E402
import download_label as _step_download  # noqa: E402
import refresh_return as _step_refresh  # noqa: E402
import update_return as _step_update  # noqa: E402

StepFn = Callable[[Optional[Sequence[str]]], int]

STEPS: Tuple[Tuple[str, StepFn], ...] = (
    ("create_return", _step_create.main),
    ("update_return", _step_update.main),
    ("download_label", _step_download.main),
    ("refresh_return", _step_refresh.main),
)

STEP_NAMES = tuple(name for name, _ in STEPS)

# 各步执行前等待（秒）；create 为流水线首步，默认 0
DEFAULT_STEP_DELAYS: Dict[str, float] = {
    "create_return": 0.0,
    "update_return": 35.0,
    "download_label": 3.0,
    "refresh_return": 5.0,
}

LOG_PREFIX = "returned"


def _parse_delay_arg(raw: str) -> Tuple[str, float]:
    """``STEP=SEC`` / ``STEP:SEC`` → (step_name, seconds)。"""
    text = str(raw or "").strip()
    sep = "=" if "=" in text else (":" if ":" in text else None)
    if not sep:
        raise argparse.ArgumentTypeError(
            f"--delay 格式应为 STEP=SEC，收到: {raw!r}"
        )
    name, sec_s = text.split(sep, 1)
    name = name.strip()
    sec_s = sec_s.strip()
    if name not in STEP_NAMES:
        raise argparse.ArgumentTypeError(
            f"--delay 未知步骤 {name!r}，可选: {', '.join(STEP_NAMES)}"
        )
    try:
        sec = float(sec_s)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            f"--delay 秒数非法: {sec_s!r}"
        ) from exc
    if sec < 0:
        raise argparse.ArgumentTypeError(f"--delay 秒数不能为负: {sec}")
    return name, sec


def resolve_step_delays(
    *,
    config_delays: Mapping[str, float] | None = None,
    cli_delays: Mapping[str, float] | None = None,
) -> Dict[str, float]:
    """合并各步延迟：内置默认 ← config ``step_delays`` ← CLI ``--delay``。"""
    out = {name: float(DEFAULT_STEP_DELAYS.get(name, 0.0)) for name in STEP_NAMES}
    for key, value in (config_delays or {}).items():
        name = str(key).strip()
        if name not in out:
            continue
        out[name] = max(0.0, float(value))
    for key, value in (cli_delays or {}).items():
        name = str(key).strip()
        if name in out:
            out[name] = max(0.0, float(value))
    return out


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "顺序执行 create_return → update_return → download_label → refresh_return"
        )
    )
    parser.add_argument(
        "--delay",
        action="append",
        dest="delays",
        type=_parse_delay_arg,
        default=None,
        metavar="STEP=SEC",
        help=(
            "单步执行前等待秒数，可重复。"
            f"步骤: {', '.join(STEP_NAMES)}。"
            "例: --delay update_return=35 --delay download_label=10"
        ),
    )
    parser.add_argument(
        "--continue-on-error",
        action="store_true",
        help="某步失败（非 0）后仍继续后续步骤；默认遇错即停",
    )
    parser.add_argument(
        "--skip",
        action="append",
        dest="skips",
        choices=list(STEP_NAMES),
        default=None,
        help="跳过指定步骤（可重复）",
    )
    parser.add_argument(
        "--only",
        action="append",
        dest="only",
        choices=list(STEP_NAMES),
        default=None,
        help="仅执行指定步骤（可重复；与 --skip 同时用时以 --only 为准）",
    )
    parser.add_argument(
        "--no-log",
        action="store_true",
        help="不写 dist/logs（或仓库 logs）下的按日日志文件",
    )
    # 透传给各步（各步 argparse 均支持的常用项）
    parser.add_argument("--workbook-id", default=None, help="透传各步")
    parser.add_argument("--dry-run", action="store_true", help="透传各步")
    parser.add_argument("--force", action="store_true", help="透传各步")
    parser.add_argument("--no-write-back", action="store_true", help="透传各步")
    parser.add_argument("--limit", type=int, default=None, help="透传各步")
    parser.add_argument(
        "--operation-desc",
        dest="operation_desc",
        default=None,
        help="透传 create_return：OMS operation_desc",
    )
    return parser


def _child_argv(args: argparse.Namespace) -> List[str]:
    argv: List[str] = []
    if args.workbook_id:
        argv.extend(["--workbook-id", str(args.workbook_id)])
    if args.dry_run:
        argv.append("--dry-run")
    if args.force:
        argv.append("--force")
    if args.no_write_back:
        argv.append("--no-write-back")
    if args.limit is not None:
        argv.extend(["--limit", str(args.limit)])
    if args.operation_desc:
        argv.extend(["--operation-desc", str(args.operation_desc)])
    return argv


def _selected_steps(args: argparse.Namespace) -> List[Tuple[str, StepFn]]:
    if args.only:
        wanted = set(args.only)
        return [(n, fn) for n, fn in STEPS if n in wanted]
    skips = set(args.skips or ())
    return [(n, fn) for n, fn in STEPS if n not in skips]


def _cli_delay_map(args: argparse.Namespace) -> Dict[str, float]:
    pairs = args.delays or []
    return {name: sec for name, sec in pairs}


def _run_pipeline(args: argparse.Namespace) -> int:
    delays = resolve_step_delays(
        config_delays=_CFG.step_delays,
        cli_delays=_cli_delay_map(args),
    )
    child = _child_argv(args)
    steps = _selected_steps(args)

    if not steps:
        print("[FAIL] 没有可执行步骤（检查 --only / --skip）", file=sys.stderr)
        return 2

    delay_summary = ",".join(f"{n}={delays.get(n, 0):g}" for n, _ in steps)
    print(
        f"[RUN_TASK] start={datetime.now().isoformat(timespec='seconds')} "
        f"steps={[n for n, _ in steps]} delays={{ {delay_summary} }} "
        f"child_argv={child or '(defaults)'}"
    )

    worst = 0
    for i, (name, fn) in enumerate(steps):
        wait = max(0.0, float(delays.get(name, 0.0)))
        if i > 0 and wait > 0:
            prev_name = steps[i - 1][0]
            print(f"[WAIT] {wait:g}s before {name} (after {prev_name}) ...")
            time.sleep(wait)
        elif i == 0 and wait > 0:
            print(f"[WAIT] {wait:g}s before {name} (startup) ...")
            time.sleep(wait)

        print(f"[STEP] >>> {name}")
        started = time.perf_counter()
        try:
            code = int(fn(child))
        except SystemExit as exc:
            code = int(exc.code) if isinstance(exc.code, int) else 1
        except Exception as exc:
            print(f"[STEP] <<< {name} EXCEPTION: {exc}", file=sys.stderr)
            code = 1
        elapsed = time.perf_counter() - started
        print(f"[STEP] <<< {name} exit={code} elapsed={elapsed:.1f}s")

        if code != 0:
            worst = code if worst == 0 else worst
            if not args.continue_on_error:
                print(f"[RUN_TASK] stop on error from {name}", file=sys.stderr)
                return code

    print(
        f"[RUN_TASK] done={datetime.now().isoformat(timespec='seconds')} "
        f"worst_exit={worst}"
    )
    return worst


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    if args.no_log:
        return _run_pipeline(args)

    log_path = daily_log_path(LOG_PREFIX)
    log_fp = log_path.open("a", encoding="utf-8", newline="\n")
    old_out, old_err = sys.stdout, sys.stderr
    sys.stdout = TeeTextIO(old_out, log_fp)
    sys.stderr = TeeTextIO(old_err, log_fp)
    try:
        print(f"[LOG] file={log_path}")
        return _run_pipeline(args)
    finally:
        sys.stdout = old_out
        sys.stderr = old_err
        log_fp.close()


if __name__ == "__main__":
    raise SystemExit(main())
