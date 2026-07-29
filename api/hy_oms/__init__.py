"""鸿羽（HY）海外仓 OMS API 客户端。

文档：http://oms.gindalogistik.com/api-doc/index.php
协议：易仓系 SOAP `callService`（WSDL: ``{base}/default/svc/wsdl``）

凭证直接写在 ``api/hy_oms/config.py`` 的 ``APP_TOKEN`` / ``APP_KEY``。
"""

from .client import HyOmsClient
from .exceptions import HyOmsAuthError, HyOmsError, HyOmsResponseError
from .web_session import DownloadResult, HyOmsWebSession

__all__ = [
    "HyOmsClient",
    "HyOmsWebSession",
    "DownloadResult",
    "HyOmsError",
    "HyOmsAuthError",
    "HyOmsResponseError",
]
