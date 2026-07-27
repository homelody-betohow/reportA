"""钉钉文档表格（Workbook）与通讯录 API 客户端。

文档：https://open.dingtalk.com/document/development/overview-of-document-tables
协议：HTTP ``/v1.0/doc/workbooks/{workbookId}/...`` + ``x-acs-dingtalk-access-token``
通讯录：``oapi.dingtalk.com/topapi/v2/user/get``（userId → unionId）；
全员拉取见 ``api.ding_disk.getUsers``。

凭证写在 ``api/ding_disk/config.py`` 的 ``APP_KEY`` / ``APP_SECRET`` / ``OPERATOR_ID``。
``OPERATOR_ID`` 可填 userId 或 unionId；调用表格接口时会自动解析为 unionId。

高层读写（DataFrame / 分块）见 ``api.ding_disk.workbook.Workbook``。
"""

from .client import (
    DingDiskClient,
    build_a1_range,
    parse_a1_cell,
    workbook_id_from_url,
)
from .exceptions import DingDiskAuthError, DingDiskError, DingDiskResponseError
from .workbook import (
    MAX_RANGE_CELLS,
    Workbook,
    append_dataframe,
    list_sheets,
    read_all_sheets,
    read_sheet,
    write_dataframe,
)

__all__ = [
    "DingDiskClient",
    "DingDiskError",
    "DingDiskAuthError",
    "DingDiskResponseError",
    "MAX_RANGE_CELLS",
    "Workbook",
    "append_dataframe",
    "build_a1_range",
    "list_sheets",
    "parse_a1_cell",
    "read_all_sheets",
    "read_sheet",
    "workbook_id_from_url",
    "write_dataframe",
]
