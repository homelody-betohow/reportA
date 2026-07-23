"""钉钉文档表格 API 异常。"""


class DingDiskError(Exception):
    """钉钉文档 / 表格对接基础异常。"""


class DingDiskAuthError(DingDiskError):
    """AppKey / AppSecret / accessToken / operatorId 无效或未配置。"""


class DingDiskResponseError(DingDiskError):
    """接口返回非 2xx 或无法解析响应。"""

    def __init__(self, message: str, *, err_code=None, http_status=None, raw=None):
        super().__init__(message)
        self.err_code = err_code
        self.http_status = http_status
        self.raw = raw
