"""钉钉文档表格（Workbook）连接配置。

在开放平台创建企业内部应用后，填写 Client ID / Client Secret，并开通：
- Document.Workbook.Read（读）
- Document.Workbook.Write（写）

文档：https://open.dingtalk.com/document/development/overview-of-document-tables
获取 accessToken：https://open.dingtalk.com/document/orgapp/obtain-the-access-token-of-an-internal-app
"""

from __future__ import annotations

from dataclasses import dataclass

# ---------------------------------------------------------------------------
# 在此配置凭证
# ---------------------------------------------------------------------------
# 企业内部应用 Client ID（原 AppKey）
APP_KEY = "dinggnu4oe6cebpi2iuv"
# 企业内部应用 Client Secret（原 AppSecret）
APP_SECRET = "3PS72YfveWzVG8-cbxLpyOFXFjKQovQABquTOPWDeBDtMXPGY8BYGCJir1n923tq"
# 操作人：可填 userId 或 unionId（表格接口需要 unionId；填 userId 时客户端会经通讯录自动换取）
OPERATOR_ID = "016067253334-1323510411"
# OPERATOR_ID = "04GTaRiP1OHfGdHCvH1ndjwiEiE"
# 可选：默认表格文件 ID（知识库 nodeId / dentryUuid，即 workbookId）
# 业务脚本也可在 app 层自行指定，不必写在这里
WORKBOOK_ID = ""

API_HOST = "https://api.dingtalk.com"
# 旧版 topapi（通讯录 user/get 等）主机
OAPI_HOST = "https://oapi.dingtalk.com"
TIMEOUT = 60.0
# accessToken 提前刷新缓冲（秒）；官方有效期约 7200 秒
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
                "未配置钉钉凭证：请在 api/ding_disk/config.py 中填写 APP_KEY / APP_SECRET。"
            )
        if require_operator and not (self.operator_id or "").strip():
            raise ValueError(
                "未配置操作人：请在 api/ding_disk/config.py 中填写 OPERATOR_ID（unionId）。"
            )
        return self

    @classmethod
    def default(cls) -> "DingDiskConfig":
        return cls().validate()

    @classmethod
    def from_env(cls, **_kwargs) -> "DingDiskConfig":
        return cls.default()
