from __future__ import annotations

from typing import Any


class EccangConfigError(RuntimeError):
    """易仓 ERP 配置缺失或无效。"""


class EccangApiError(RuntimeError):
    """易仓 ERP 接口返回失败或 HTTP 异常。"""

    def __init__(
        self,
        message: str,
        *,
        code: str | None = None,
        ask: str | None = None,
        method: str | None = None,
        raw_payload: dict[str, Any] | None = None,
        raw_text: str | None = None,
    ) -> None:
        super().__init__(message)
        self.code = code
        self.ask = ask
        self.method = method
        self.raw_payload = raw_payload
        self.raw_text = raw_text
