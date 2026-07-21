"""鸿羽 OMS API 异常。"""


class HyOmsError(Exception):
    """鸿羽 OMS 对接基础异常。"""


class HyOmsAuthError(HyOmsError):
    """appToken / appKey 无效或未配置。"""


class HyOmsResponseError(HyOmsError):
    """接口返回 ask=Failure 或无法解析响应。"""

    def __init__(self, message: str, *, err_code=None, raw=None):
        super().__init__(message)
        self.err_code = err_code
        self.raw = raw
