"""易仓 ERP 开放平台连接配置。

直接在此填写 APP_KEY / APP_SECRET / SERVICE_ID（开放平台 → 应用管理）。
文档：https://open.eccang.com
"""

from __future__ import annotations

from dataclasses import dataclass

from .exceptions import EccangConfigError

_PROD_BASE_URL = "http://openapi-web.eccang.com/openApi/api/unity"

# ---------------------------------------------------------------------------
# 在此配置凭证
# ---------------------------------------------------------------------------
APP_KEY = "c1591fbbf3f344d3"
APP_SECRET = "b8f25afe805d48e0"
SERVICE_ID = "ERP2009189VG"  # 服务ID，需在易仓开放平台授权后获取

# 正式环境默认见 _PROD_BASE_URL；留空则使用默认
BASE_URL = ""
CHARSET = "UTF-8"
SIGN_TYPE = "MD5"
VERSION = "1.0.0"
TIMEOUT = 60


@dataclass(frozen=True)
class EccangConfig:
    """易仓 ERP 开放平台配置（凭证见本文件顶部常量）。"""

    app_key: str = APP_KEY
    app_secret: str = APP_SECRET
    service_id: str = SERVICE_ID
    base_url: str = _PROD_BASE_URL
    charset: str = CHARSET
    sign_type: str = SIGN_TYPE
    version: str = VERSION
    timeout: int = TIMEOUT

    def validate(self) -> "EccangConfig":
        if (
            not (self.app_key or "").strip()
            or not (self.app_secret or "").strip()
            or not (self.service_id or "").strip()
        ):
            raise EccangConfigError(
                "缺少易仓 ERP 凭证：请在 api/eccang/config.py 中填写 "
                "APP_KEY、APP_SECRET 与 SERVICE_ID。\n"
                "获取方式：\n"
                "1. 访问 https://open.eccang.com 登录\n"
                "2. 进入应用管理 -> 新增应用（选择服务商应用）\n"
                "3. 填写应用信息并设置 IP 白名单\n"
                "4. 审核通过后获取 APP_KEY 和 APP_SECRET\n"
                "5. 联系易仓客户授权获取 SERVICE_ID"
            )
        return self

    @classmethod
    def default(cls) -> "EccangConfig":
        """从模块常量加载默认配置。"""
        base_url = (BASE_URL or "").strip() or _PROD_BASE_URL
        try:
            timeout = max(5, int(TIMEOUT))
        except (TypeError, ValueError):
            timeout = 60

        return cls(
            app_key=(APP_KEY or "").strip(),
            app_secret=(APP_SECRET or "").strip(),
            service_id=(SERVICE_ID or "").strip(),
            base_url=base_url.rstrip("/"),
            charset=(CHARSET or "UTF-8").strip() or "UTF-8",
            sign_type=(SIGN_TYPE or "MD5").strip() or "MD5",
            version=(VERSION or "1.0.0").strip() or "1.0.0",
            timeout=timeout,
        ).validate()

    @classmethod
    def from_env(cls) -> "EccangConfig":
        """兼容旧调用，实际读取 config.py 中的常量。"""
        return cls.default()
