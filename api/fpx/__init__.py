"""4PX（递四方 / FPX）开放平台 API 客户端。

文档：https://open.4px.com/v2/doc/detail?ids=55,88,214
协议：HTTP POST ``/router/api/service`` + MD5 签名

凭证直接写在 ``api/fpx/config.py`` 的 ``APP_KEY`` / ``APP_SECRET``。
"""

from .client import FpxClient, build_sign
from .exceptions import FpxAuthError, FpxError, FpxResponseError

__all__ = [
    "FpxClient",
    "FpxError",
    "FpxAuthError",
    "FpxResponseError",
    "build_sign",
]
