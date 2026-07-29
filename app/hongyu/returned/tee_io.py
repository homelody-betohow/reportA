"""控制台与日志文件双写（tee）。"""
from __future__ import annotations

from typing import TextIO


class TeeTextIO:
    """把写入同时送到控制台与日志文件。"""

    def __init__(self, primary: TextIO, secondary: TextIO) -> None:
        self._primary = primary
        self._secondary = secondary

    def write(self, data: str) -> int:
        n = self._primary.write(data)
        self._secondary.write(data)
        self._primary.flush()
        self._secondary.flush()
        return n

    def flush(self) -> None:
        self._primary.flush()
        self._secondary.flush()

    def isatty(self) -> bool:
        return bool(getattr(self._primary, "isatty", lambda: False)())

    @property
    def encoding(self) -> str:
        return getattr(self._primary, "encoding", None) or "utf-8"

    def reconfigure(self, **kwargs) -> None:
        fn = getattr(self._primary, "reconfigure", None)
        if callable(fn):
            fn(**kwargs)
