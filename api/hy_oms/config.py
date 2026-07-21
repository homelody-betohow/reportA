"""鸿羽 OMS 连接配置。

直接在此填写 appToken / appKey（OMS 后台 → API 密钥）。
"""

from __future__ import annotations

from dataclasses import dataclass

# ---------------------------------------------------------------------------
# 在此配置凭证
# ---------------------------------------------------------------------------
APP_TOKEN = "a8bac7475e212cb720a1cf26a31946f3"  # 填写鸿羽 OMS appToken
APP_KEY = "6daf9ae6ed2759a00dfb31c5b978d8be"    # 填写鸿羽 OMS appKey

BASE_URL = "http://oms.gindalogistik.com"
LANGUAGE = "zh_CN"
TIMEOUT = 60.0


@dataclass(frozen=True)
class HyOmsConfig:
    app_token: str = APP_TOKEN
    app_key: str = APP_KEY
    base_url: str = BASE_URL
    language: str = LANGUAGE
    timeout: float = TIMEOUT

    @property
    def service_url(self) -> str:
        return f"{self.base_url.rstrip('/')}/default/svc/web-service"

    @property
    def wsdl_url(self) -> str:
        return f"{self.base_url.rstrip('/')}/default/svc/wsdl"

    def validate(self) -> "HyOmsConfig":
        if not (self.app_token or "").strip() or not (self.app_key or "").strip():
            raise ValueError(
                "未配置鸿羽 OMS 凭证：请在 api/hy_oms/config.py 中填写 APP_TOKEN / APP_KEY。"
            )
        return self

    @classmethod
    def default(cls) -> "HyOmsConfig":
        """使用本文件顶部的 APP_TOKEN / APP_KEY。"""
        return cls().validate()

    # 兼容旧调用名
    @classmethod
    def from_env(cls, **_kwargs) -> "HyOmsConfig":
        return cls.default()
