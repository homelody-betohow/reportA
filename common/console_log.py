"""终端彩色日志工具，供 snapshot 等脚本共用。

用法（在脚本入口将项目根加入 sys.path 后）::

    from common.console_log import (
        enable_windows_ansi,
        log,
        log_error,
        log_success,
        log_warning,
    )

    enable_windows_ansi()
    log("INFO", "开始")
    log_success("完成")
"""

from __future__ import annotations

import os
import sys
from datetime import datetime

_RESET = "\033[0m"
_ANSI = {
    "RED": "\033[91m",
    "GREEN": "\033[92m",
    "YELLOW": "\033[93m",
    "CYAN": "\033[96m",
    "BOLD": "\033[1m",
    "BG_RED": "\033[41m",
    "WHITE": "\033[97m",
}

_ansi_enabled = False


def enable_windows_ansi() -> None:
    """Windows 控制台启用 ANSI 转义（幂等，脚本入口调用一次即可）。"""
    global _ansi_enabled
    if _ansi_enabled:
        return
    if os.name == "nt":
        try:
            import ctypes

            kernel32 = ctypes.windll.kernel32
            for handle_id in (-11, -12):
                handle = kernel32.GetStdHandle(handle_id)
                mode = ctypes.c_uint32()
                if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
                    kernel32.SetConsoleMode(handle, mode.value | 0x0004)
        except Exception:
            pass
    _ansi_enabled = True


def use_color() -> bool:
    if os.getenv("NO_COLOR"):
        return False
    if not hasattr(sys.stdout, "isatty") or not sys.stdout.isatty():
        return False
    return True


def colorize(text: str, *styles: str) -> str:
    """给文本加 ANSI 样式（styles: RED/GREEN/YELLOW/CYAN/BOLD/BG_RED/WHITE）。"""
    if not use_color():
        return text
    codes = "".join(_ANSI.get(s, s) for s in styles)
    return f"{codes}{text}{_RESET}"


_c = colorize


def _timestamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(level: str, msg: str, *, flush: bool = True) -> None:
    print(f"[{_timestamp()}] [{level}] {msg}", flush=flush)


def log_info(msg: str, *, flush: bool = True) -> None:
    log("INFO", msg, flush=flush)


def log_success(msg: str, *, flush: bool = True) -> None:
    ts = _timestamp()
    print(
        f"[{ts}] [{colorize('OK', 'GREEN', 'BOLD')}] {colorize(msg, 'GREEN', 'BOLD')}",
        flush=flush,
    )


def log_error(msg: str, *, flush: bool = True) -> None:
    ts = _timestamp()
    print(
        f"[{ts}] [{colorize('ERROR', 'RED', 'BOLD')}] {colorize(msg, 'RED', 'BOLD')}",
        flush=flush,
    )


def log_warning(msg: str, *, flush: bool = True) -> None:
    ts = _timestamp()
    print(
        f"[{ts}] [{colorize('WARN', 'YELLOW', 'BOLD')}] {colorize(msg, 'YELLOW', 'BOLD')}",
        flush=flush,
    )
