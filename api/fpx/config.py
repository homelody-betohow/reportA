"""4PX（递四方 / FPX）开放平台连接配置。

直接在此填写 AppKey / AppSecret（开放平台 → 接入管理 → 我的应用）。
文档：https://open.4px.com/v2/doc/detail?ids=55,88,214
接入说明：https://open.au.4px.com/apiInfo/merchant
"""

from __future__ import annotations

from dataclasses import dataclass

# ---------------------------------------------------------------------------
# 在此配置凭证
# ---------------------------------------------------------------------------
# 填写 4PX 开放平台 AppKey
APP_KEY = "ceff21b8-939c-4919-858b-70582fa358d2"
# 填写 4PX 开放平台 AppSecret
APP_SECRET = "22f42ee7-d227-483c-b758-c330cea58056"
# 合作伙伴/ISV 需要；4PX 商家（B 类）通常可不填
ACCESS_TOKEN = ""

# 正式：https://open.4px.com ；沙箱：https://open-test.4px.com
BASE_URL = "https://open.4px.com"
API_VERSION = "1.0.0"
LANGUAGE = "cn"
FORMAT = "json"
TIMEOUT = 60.0
SANDBOX = False


@dataclass(frozen=True)
class FpxConfig:
    app_key: str = APP_KEY
    app_secret: str = APP_SECRET
    access_token: str = ACCESS_TOKEN
    base_url: str = BASE_URL
    api_version: str = API_VERSION
    language: str = LANGUAGE
    format: str = FORMAT
    timeout: float = TIMEOUT
    sandbox: bool = SANDBOX

    @property
    def service_url(self) -> str:
        if self.sandbox:
            host = "https://open-test.4px.com"
        else:
            host = self.base_url.rstrip("/") or "https://open.4px.com"
        return f"{host}/router/api/service"

    def validate(self) -> "FpxConfig":
        if not (self.app_key or "").strip() or not (self.app_secret or "").strip():
            raise ValueError(
                "未配置 4PX 凭证：请在 api/fpx/config.py 中填写 APP_KEY / APP_SECRET。"
            )
        return self

    @classmethod
    def default(cls) -> "FpxConfig":
        return cls().validate()

    @classmethod
    def from_env(cls, **_kwargs) -> "FpxConfig":
        return cls.default()
