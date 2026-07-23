"""4PX 开放平台 API 异常。"""


class FpxError(Exception):
    """4PX 对接基础异常。"""


class FpxAuthError(FpxError):
    """AppKey / AppSecret / access_token 无效或未配置。"""


class FpxResponseError(FpxError):
    """接口返回 result!=1 或无法解析响应。"""

    def __init__(self, message: str, *, err_code=None, raw=None):
        super().__init__(message)
        self.err_code = err_code
        self.raw = raw
