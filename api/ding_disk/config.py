"""钉钉文档表格（Workbook）连接配置。

凭证优先从 ``config/secrets.json``（或环境变量 ``DING_DISK_*``）读取；
也可在本文件填写兜底默认值（打包进 exe 时不建议写入真实密钥）。

文档：https://open.dingtalk.com/document/development/overview-of-document-tables
"""

from __future__ import annotations

from dataclasses import dataclass

from common.secrets_config import ding_disk_secrets

# ---------------------------------------------------------------------------
# 可选兜底（真实凭证请写在 config/secrets.json）
# ---------------------------------------------------------------------------
APP_KEY = ""
APP_SECRET = ""
OPERATOR_ID = ""
WORKBOOK_ID = ""

API_HOST = "https://api.dingtalk.com"
OAPI_HOST = "https://oapi.dingtalk.com"
TIMEOUT = 60.0
TOKEN_REFRESH_SKEW = 120.0


@dataclass(frozen=True)
class DingDiskConfig:
    app_key: str = APP_KEY
    app_secret: str = APP_SECRET
    operator_id: str = OPERATOR_ID
    workbook_id: str = WORKBOOK_ID
    api_host: str = API_HOST
    oapi_host: str = OAPI_HOST
    timeout: float = TIMEOUT
    token_refresh_skew: float = TOKEN_REFRESH_SKEW

    def validate(self, *, require_operator: bool = False) -> "DingDiskConfig":
        if not (self.app_key or "").strip() or not (self.app_secret or "").strip():
            raise ValueError(
                "未配置钉钉凭证：请在 config/secrets.json 的 ding_disk 中填写 "
                "app_key / app_secret（或设置环境变量 DING_DISK_APP_KEY / DING_DISK_APP_SECRET）。"
            )
        if require_operator and not (self.operator_id or "").strip():
            raise ValueError(
                "未配置操作人：请在 config/secrets.json 的 ding_disk 中填写 operator_id"
                "（或设置 DING_DISK_OPERATOR_ID）。"
            )
        return self

    @classmethod
    def default(cls) -> "DingDiskConfig":
        merged = ding_disk_secrets(
            defaults={
                "app_key": APP_KEY,
                "app_secret": APP_SECRET,
                "operator_id": OPERATOR_ID,
                "workbook_id": WORKBOOK_ID,
                "api_host": API_HOST,
                "oapi_host": OAPI_HOST,
                "timeout": TIMEOUT,
                "token_refresh_skew": TOKEN_REFRESH_SKEW,
            }
        )
        return cls(
            app_key=str(merged["app_key"]),
            app_secret=str(merged["app_secret"]),
            operator_id=str(merged["operator_id"]),
            workbook_id=str(merged["workbook_id"]),
            api_host=str(merged["api_host"]),
            oapi_host=str(merged["oapi_host"]),
            timeout=float(merged["timeout"]),
            token_refresh_skew=float(merged["token_refresh_skew"]),
        ).validate()

    @classmethod
    def from_env(cls, **_kwargs) -> "DingDiskConfig":
        return cls.default()
