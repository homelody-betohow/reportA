"""易仓 ERP 开放平台 API 客户端。

文档：https://open.eccang.com/#/documentCenter
协议：HTTP POST ``/openApi/api/unity`` + MD5 签名

凭证直接写在 ``api/eccang/config.py`` 的 ``APP_KEY`` / ``APP_SECRET`` / ``SERVICE_ID``。
"""

from .client import EccangClient
from .config import EccangConfig
from .exceptions import EccangApiError, EccangConfigError
from .methods import EccangMethods, EccangService

__all__ = [
    "EccangClient",
    "EccangConfig",
    "EccangApiError",
    "EccangConfigError",
    "EccangMethods",
    "EccangService",
]
