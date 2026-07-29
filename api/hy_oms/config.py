"""鸿羽 OMS 连接配置。

凭证优先从 ``config/secrets.json``（或环境变量 ``HY_OMS_*``）读取；
也可在本文件填写兜底默认值（打包进 exe 时不建议写入真实密钥）。
"""

from __future__ import annotations

from dataclasses import dataclass

from common.secrets_config import hy_oms_secrets

# ---------------------------------------------------------------------------
# 可选兜底（真实凭证请写在 config/secrets.json）
# ---------------------------------------------------------------------------
APP_TOKEN = ""
APP_KEY = ""

USER_ACCOUNT = ""
USER_PASSWORD = ""
COMPANY_CODE = ""

BASE_URL = "http://oms.gindalogistik.com"
LANGUAGE = "zh_CN"
TIMEOUT = 60.0


@dataclass(frozen=True)
class HyOmsConfig:
    app_token: str = APP_TOKEN
    app_key: str = APP_KEY
    user_account: str = USER_ACCOUNT
    user_password: str = USER_PASSWORD
    company_code: str = COMPANY_CODE
    base_url: str = BASE_URL
    language: str = LANGUAGE
    timeout: float = TIMEOUT

    @property
    def service_url(self) -> str:
        return f"{self.base_url.rstrip('/')}/default/svc/web-service"

    @property
    def wsdl_url(self) -> str:
        return f"{self.base_url.rstrip('/')}/default/svc/wsdl"

    @property
    def quick_login_path(self) -> str:
        return f"{self.base_url.rstrip('/')}/default/index/quick-login"

    def validate(self) -> "HyOmsConfig":
        if not (self.app_token or "").strip() or not (self.app_key or "").strip():
            raise ValueError(
                "未配置鸿羽 OMS 凭证：请在 config/secrets.json 的 hy_oms 中填写 "
                "app_token / app_key（或设置环境变量 HY_OMS_APP_TOKEN / HY_OMS_APP_KEY）。"
            )
        return self

    @classmethod
    def default(cls) -> "HyOmsConfig":
        merged = hy_oms_secrets(
            defaults={
                "app_token": APP_TOKEN,
                "app_key": APP_KEY,
                "user_account": USER_ACCOUNT,
                "user_password": USER_PASSWORD,
                "company_code": COMPANY_CODE,
                "base_url": BASE_URL,
                "language": LANGUAGE,
                "timeout": TIMEOUT,
            }
        )
        return cls(
            app_token=str(merged["app_token"]),
            app_key=str(merged["app_key"]),
            user_account=str(merged["user_account"]),
            user_password=str(merged["user_password"]),
            company_code=str(merged["company_code"]),
            base_url=str(merged["base_url"]),
            language=str(merged["language"]),
            timeout=float(merged["timeout"]),
        ).validate()

    @classmethod
    def from_env(cls, **_kwargs) -> "HyOmsConfig":
        return cls.default()
