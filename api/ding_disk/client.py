"""钉钉文档表格（Workbook）HTTP 客户端。

协议：新版服务端 API ``https://api.dingtalk.com/v1.0/doc/workbooks/...``
鉴权：``x-acs-dingtalk-access-token``（企业内部应用 oauth2 accessToken）

文档入口::
    https://open.dingtalk.com/document/development/overview-of-document-tables

``workbookId`` 为知识库节点 ``nodeId``（``dentryUuid``），也可从表格 URL 取得。
``sheetId`` 可为工作表 ID 或工作表名称。
``rangeAddress`` 使用 A1 表示法，如 ``A1:C10``。
"""

from __future__ import annotations

import json
import re
import time
from typing import Any, Dict, List, Mapping, MutableMapping, Optional, Sequence, Union
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

from .config import DingDiskConfig
from .exceptions import DingDiskAuthError, DingDiskError, DingDiskResponseError

JsonDict = Dict[str, Any]
Params = Optional[Mapping[str, Any]]
CellMatrix = Sequence[Sequence[Any]]

_TOKEN_PATH = "/v1.0/oauth2/accessToken"
_SHEETS_PATH = "/v1.0/doc/workbooks/{workbookId}/sheets"
_SHEET_PATH = "/v1.0/doc/workbooks/{workbookId}/sheets/{sheetId}"
_RANGE_PATH = "/v1.0/doc/workbooks/{workbookId}/sheets/{sheetId}/ranges/{rangeAddress}"
_APPEND_ROWS_PATH = "/v1.0/doc/workbooks/{workbookId}/sheets/{sheetId}/appendRows"
_CLEAR_DATA_PATH = (
    "/v1.0/doc/workbooks/{workbookId}/sheets/{sheetId}/ranges/{rangeAddress}/clearData"
)
_CLEAR_ALL_PATH = (
    "/v1.0/doc/workbooks/{workbookId}/sheets/{sheetId}/ranges/{rangeAddress}/clear"
)


def _col_to_a1(col_index: int) -> str:
    """0-based column index → A1 列标（0→A, 25→Z, 26→AA）。"""
    if col_index < 0:
        raise ValueError(f"列索引不能为负: {col_index}")
    n = col_index + 1
    letters: List[str] = []
    while n:
        n, rem = divmod(n - 1, 26)
        letters.append(chr(65 + rem))
    return "".join(reversed(letters))


def parse_a1_cell(cell: str) -> tuple[int, int]:
    """解析单个单元格地址，返回 (row_index_0based, col_index_0based)。"""
    text = (cell or "").strip().upper()
    m = re.fullmatch(r"([A-Z]+)(\d+)", text)
    if not m:
        raise ValueError(f"非法 A1 单元格地址: {cell!r}")
    col_letters, row_s = m.group(1), m.group(2)
    col = 0
    for ch in col_letters:
        col = col * 26 + (ord(ch) - 64)
    return int(row_s) - 1, col - 1


def build_a1_range(
    values: CellMatrix,
    *,
    start_cell: str = "A1",
) -> str:
    """根据二维 values 与起点单元格生成 A1 range（如 ``A1:C3``）。"""
    if not values:
        raise ValueError("values 不能为空")
    rows = len(values)
    cols = max((len(r) for r in values), default=0)
    if cols <= 0:
        raise ValueError("values 至少需要一列")
    start_row, start_col = parse_a1_cell(start_cell)
    end_row = start_row + rows - 1
    end_col = start_col + cols - 1
    start = f"{_col_to_a1(start_col)}{start_row + 1}"
    end = f"{_col_to_a1(end_col)}{end_row + 1}"
    return start if start == end else f"{start}:{end}"


def workbook_id_from_url(url: str) -> str:
    """从钉钉表格 URL 尽量提取 workbookId（nodeId / dentryUuid）。

    常见形态含 ``/i/nodes/<id>``、``dentryUuid=``、``nodeId=``、``workbookId=``。
    提取失败时返回空字符串。
    """
    text = (url or "").strip()
    if not text:
        return ""
    for pat in (
        r"/i/nodes/([A-Za-z0-9_-]+)",
        r"[?&#](?:dentryUuid|nodeId|workbookId)=([A-Za-z0-9_-]+)",
        r"/workbook[s]?/([A-Za-z0-9_-]+)",
    ):
        m = re.search(pat, text, flags=re.I)
        if m:
            return m.group(1)
    return ""


class DingDiskClient:
    """钉钉文档表格 + 通讯录客户端。

    任意新版接口可 ``request("GET", "/v1.0/doc/...")``；
    通讯录 topapi 见 ``get_user`` / ``get_userid_by_mobile``。
    表格类接口的 ``operatorId`` 需要 unionId：若配置的是 userId，会自动换取。
    """

    def __init__(self, config: DingDiskConfig):
        self.config = config
        self._access_token: Optional[str] = None
        self._token_expire_at: float = 0.0
        self._union_id_cache: Dict[str, str] = {}

    @classmethod
    def from_config(cls) -> "DingDiskClient":
        return cls(DingDiskConfig.default())

    @classmethod
    def from_env(cls, **_kwargs) -> "DingDiskClient":
        return cls.from_config()

    # ------------------------------------------------------------------
    # Token
    # ------------------------------------------------------------------

    def get_access_token(self, *, force_refresh: bool = False) -> str:
        """获取（并缓存）企业内部应用 accessToken。"""
        now = time.time()
        if (
            not force_refresh
            and self._access_token
            and now < self._token_expire_at - self.config.token_refresh_skew
        ):
            return self._access_token

        payload = {
            "appKey": self.config.app_key,
            "appSecret": self.config.app_secret,
        }
        data = self._http_json(
            "POST",
            _TOKEN_PATH,
            body=payload,
            with_token=False,
        )
        token = str(data.get("accessToken") or "").strip()
        if not token:
            raise DingDiskAuthError(f"未返回 accessToken: {data}")
        expire_in = data.get("expireIn", 7200)
        try:
            expire_sec = float(expire_in)
        except (TypeError, ValueError):
            expire_sec = 7200.0
        self._access_token = token
        self._token_expire_at = now + max(expire_sec, 60.0)
        return token

    # ------------------------------------------------------------------
    # 通讯录（旧版 topapi，可用应用 accessToken）
    # ------------------------------------------------------------------

    def get_user(self, userid: str, *, language: str = "zh_CN") -> JsonDict:
        """查询用户详情 ``topapi/v2/user/get``，返回含 ``userid`` / ``unionid`` / ``name`` 等。

        文档：https://open.dingtalk.com/document/orgapp/query-user-details
        权限：成员信息读权限（qyapi_get_member）。
        """
        uid = (userid or "").strip()
        if not uid:
            raise ValueError("userid 不能为空")
        data = self._oapi_json(
            "/topapi/v2/user/get",
            body={"userid": uid, "language": language},
        )
        result = data.get("result")
        if not isinstance(result, dict):
            raise DingDiskResponseError("get_user 未返回 result", raw=data)
        return result

    def get_userid_by_mobile(self, mobile: str) -> str:
        """根据手机号查 userId ``topapi/v2/user/getbymobile``。

        文档：https://open.dingtalk.com/document/orgapp/query-users-by-phone-number
        权限：根据手机号获取成员基本信息。
        """
        mobile = (mobile or "").strip()
        if not mobile:
            raise ValueError("mobile 不能为空")
        data = self._oapi_json(
            "/topapi/v2/user/getbymobile",
            body={"mobile": mobile},
        )
        result = data.get("result") or {}
        userid = str(result.get("userid") or "").strip() if isinstance(result, dict) else ""
        if not userid:
            raise DingDiskResponseError(f"未找到手机号对应用户: {mobile}", raw=data)
        return userid

    def get_unionid(self, userid: str) -> str:
        """由 userId 取 unionId。"""
        detail = self.get_user(userid)
        unionid = str(detail.get("unionid") or detail.get("unionId") or "").strip()
        if not unionid:
            raise DingDiskResponseError(f"用户无 unionid: userid={userid}", raw=detail)
        return unionid

    def resolve_operator_union_id(self, operator_id: Optional[str] = None) -> str:
        """将配置/传入的操作人标识解析为表格接口所需的 unionId。

        - 若传入值本身已是 unionId（通讯录查不到），原样返回
        - 若是 userId，则换取 unionId 并缓存
        """
        op = (operator_id if operator_id is not None else self.config.operator_id or "").strip()
        if not op:
            raise ValueError(
                "缺少 operatorId：请传入 operator_id 或在 config.py 填写 OPERATOR_ID。"
            )
        cached = self._union_id_cache.get(op)
        if cached:
            return cached
        try:
            unionid = self.get_unionid(op)
        except DingDiskResponseError:
            # 已是 unionId，或无通讯录权限时：按原值使用
            self._union_id_cache[op] = op
            return op
        self._union_id_cache[op] = unionid
        # 也缓存反向，避免重复查询
        self._union_id_cache[unionid] = unionid
        return unionid

    def _oapi_json(self, path: str, *, body: Mapping[str, Any]) -> JsonDict:
        """调用 ``oapi.dingtalk.com`` topapi，access_token 放 query。"""
        host = (self.config.oapi_host or "https://oapi.dingtalk.com").rstrip("/")
        token = self.get_access_token()
        url = f"{host}{path}?{urlencode({'access_token': token})}"
        data_bytes = json.dumps(body, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        req = Request(
            url,
            data=data_bytes,
            headers={
                "Accept": "application/json",
                "Content-Type": "application/json; charset=utf-8",
            },
            method="POST",
        )
        try:
            with urlopen(req, timeout=self.config.timeout) as resp:
                raw_text = resp.read().decode("utf-8", errors="replace")
        except HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace") if exc.fp else ""
            raise DingDiskResponseError(
                detail or str(exc.reason),
                http_status=exc.code,
                raw=detail,
            ) from exc
        except URLError as exc:
            raise DingDiskError(f"无法连接钉钉 oapi: {exc.reason}") from exc

        try:
            data = json.loads(raw_text)
        except json.JSONDecodeError as exc:
            raise DingDiskResponseError(
                f"oapi 响应不是合法 JSON: {raw_text[:200]}",
                raw=raw_text,
            ) from exc
        if not isinstance(data, dict):
            raise DingDiskResponseError("oapi 响应不是对象", raw=data)

        errcode = data.get("errcode", 0)
        try:
            err_n = int(errcode)
        except (TypeError, ValueError):
            err_n = -1
        if err_n != 0:
            errmsg = data.get("errmsg") or f"errcode={errcode}"
            if err_n in {88, 40014, 40001} or "access_token" in str(errmsg).lower():
                raise DingDiskAuthError(str(errmsg))
            raise DingDiskResponseError(str(errmsg), err_code=errcode, raw=data)
        return data

    # ------------------------------------------------------------------
    # 底层 HTTP（新版 api.dingtalk.com）
    # ------------------------------------------------------------------

    def request(
        self,
        method: str,
        path: str,
        *,
        query: Params = None,
        body: Any = None,
        operator_id: Optional[str] = None,
        with_token: bool = True,
        require_operator: bool = False,
        resolve_operator: bool = True,
    ) -> JsonDict:
        """通用请求。``path`` 以 ``/v1.0/...`` 开头。"""
        q: MutableMapping[str, Any] = dict(query or {})
        op = (operator_id if operator_id is not None else self.config.operator_id or "").strip()
        if require_operator and not op:
            raise ValueError(
                "缺少 operatorId：请传入 operator_id 或在 config.py 填写 OPERATOR_ID。"
            )
        if op and "operatorId" not in q:
            if resolve_operator and require_operator:
                op = self.resolve_operator_union_id(op)
            q["operatorId"] = op
        return self._http_json(
            method,
            path,
            query=q,
            body=body,
            with_token=with_token,
        )

    def _http_json(
        self,
        method: str,
        path: str,
        *,
        query: Params = None,
        body: Any = None,
        with_token: bool = True,
    ) -> JsonDict:
        host = self.config.api_host.rstrip("/")
        url = f"{host}{path}"
        if query:
            # 过滤 None，保留 False/0/"" 以外的有意义空串（operatorId 不应为空）
            flat = {k: v for k, v in query.items() if v is not None}
            if flat:
                url = f"{url}?{urlencode(flat, doseq=True)}"

        headers = {
            "Accept": "application/json",
            "Content-Type": "application/json; charset=utf-8",
        }
        if with_token:
            headers["x-acs-dingtalk-access-token"] = self.get_access_token()

        data_bytes: Optional[bytes] = None
        if body is not None:
            if isinstance(body, (bytes, bytearray)):
                data_bytes = bytes(body)
            elif isinstance(body, str):
                data_bytes = body.encode("utf-8")
            else:
                data_bytes = json.dumps(body, ensure_ascii=False, separators=(",", ":")).encode(
                    "utf-8"
                )
        elif method.upper() in {"POST", "PUT", "PATCH"}:
            data_bytes = b"{}"

        req = Request(url, data=data_bytes, headers=headers, method=method.upper())
        try:
            with urlopen(req, timeout=self.config.timeout) as resp:
                raw_text = resp.read().decode("utf-8", errors="replace")
                status = getattr(resp, "status", 200)
        except HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace") if exc.fp else ""
            parsed: Any = detail
            err_code = None
            message = detail or exc.reason
            try:
                parsed = json.loads(detail) if detail else {}
                if isinstance(parsed, dict):
                    err_code = parsed.get("code") or parsed.get("errorCode")
                    message = (
                        parsed.get("message")
                        or parsed.get("msg")
                        or parsed.get("errorMsg")
                        or message
                    )
            except json.JSONDecodeError:
                parsed = detail
            auth_hints = ("accessToken", "access_token", "Forbidden.AccessDenied", "unauthorized")
            if exc.code in {401, 403} or any(h.lower() in str(message).lower() for h in auth_hints):
                raise DingDiskAuthError(str(message)) from exc
            raise DingDiskResponseError(
                str(message),
                err_code=err_code,
                http_status=exc.code,
                raw=parsed,
            ) from exc
        except URLError as exc:
            raise DingDiskError(f"无法连接钉钉开放平台: {exc.reason}") from exc

        if not raw_text.strip():
            return {}
        try:
            data = json.loads(raw_text)
        except json.JSONDecodeError as exc:
            raise DingDiskResponseError(
                f"响应不是合法 JSON: {raw_text[:200]}",
                http_status=status,
                raw=raw_text,
            ) from exc
        if isinstance(data, dict):
            return data
        return {"data": data}

    @staticmethod
    def _encode_path_segment(value: str, *, keep_colon: bool = False) -> str:
        safe = ":" if keep_colon else ""
        return quote(str(value), safe=safe)

    def _workbook(self, workbook_id: Optional[str]) -> str:
        wid = (workbook_id if workbook_id is not None else self.config.workbook_id or "").strip()
        if not wid:
            raise ValueError(
                "缺少 workbookId：请传入 workbook_id 或在 config.py 填写 WORKBOOK_ID。"
            )
        return wid

    def _sheet_path(self, workbook_id: str, sheet_id: str) -> str:
        return _SHEET_PATH.format(
            workbookId=self._encode_path_segment(workbook_id),
            sheetId=self._encode_path_segment(sheet_id),
        )

    def _range_path(self, workbook_id: str, sheet_id: str, range_address: str) -> str:
        return _RANGE_PATH.format(
            workbookId=self._encode_path_segment(workbook_id),
            sheetId=self._encode_path_segment(sheet_id),
            rangeAddress=self._encode_path_segment(range_address, keep_colon=True),
        )

    # ------------------------------------------------------------------
    # 工作表
    # ------------------------------------------------------------------

    def list_sheets(
        self,
        workbook_id: Optional[str] = None,
        *,
        operator_id: Optional[str] = None,
    ) -> JsonDict:
        """获取所有工作表 ``GET .../workbooks/{id}/sheets``。"""
        wid = self._workbook(workbook_id)
        path = _SHEETS_PATH.format(workbookId=self._encode_path_segment(wid))
        return self.request("GET", path, operator_id=operator_id, require_operator=True)

    def get_sheet(
        self,
        sheet_id: str,
        workbook_id: Optional[str] = None,
        *,
        operator_id: Optional[str] = None,
    ) -> JsonDict:
        """获取工作表属性 ``GET .../sheets/{sheetId}``。"""
        wid = self._workbook(workbook_id)
        return self.request(
            "GET",
            self._sheet_path(wid, sheet_id),
            operator_id=operator_id,
            require_operator=True,
        )

    def create_sheet(
        self,
        name: str,
        workbook_id: Optional[str] = None,
        *,
        operator_id: Optional[str] = None,
    ) -> JsonDict:
        """创建工作表 ``POST .../sheets``。"""
        wid = self._workbook(workbook_id)
        path = _SHEETS_PATH.format(workbookId=self._encode_path_segment(wid))
        return self.request(
            "POST",
            path,
            body={"name": name},
            operator_id=operator_id,
            require_operator=True,
        )

    def update_sheet(
        self,
        sheet_id: str,
        workbook_id: Optional[str] = None,
        *,
        name: Optional[str] = None,
        visibility: Optional[str] = None,
        frozen_row_count: Optional[int] = None,
        frozen_column_count: Optional[int] = None,
        operator_id: Optional[str] = None,
        **extra: Any,
    ) -> JsonDict:
        """更新工作表 ``PUT .../sheets/{sheetId}``。"""
        wid = self._workbook(workbook_id)
        body: Dict[str, Any] = dict(extra)
        if name is not None:
            body["name"] = name
        if visibility is not None:
            body["visibility"] = visibility
        if frozen_row_count is not None:
            body["frozenRowCount"] = frozen_row_count
        if frozen_column_count is not None:
            body["frozenColumnCount"] = frozen_column_count
        return self.request(
            "PUT",
            self._sheet_path(wid, sheet_id),
            body=body,
            operator_id=operator_id,
            require_operator=True,
        )

    def delete_sheet(
        self,
        sheet_id: str,
        workbook_id: Optional[str] = None,
        *,
        operator_id: Optional[str] = None,
    ) -> JsonDict:
        """删除工作表 ``DELETE .../sheets/{sheetId}``。"""
        wid = self._workbook(workbook_id)
        return self.request(
            "DELETE",
            self._sheet_path(wid, sheet_id),
            operator_id=operator_id,
            require_operator=True,
        )

    # ------------------------------------------------------------------
    # 单元格区域
    # ------------------------------------------------------------------

    def get_range(
        self,
        sheet_id: str,
        range_address: str,
        workbook_id: Optional[str] = None,
        *,
        select: Optional[Union[str, Sequence[str]]] = None,
        operator_id: Optional[str] = None,
    ) -> JsonDict:
        """获取单元格区域 ``GET .../ranges/{rangeAddress}``。

        ``select`` 如 ``"values"`` 或 ``["values", "formulas"]``，不传则返回全部字段。
        """
        wid = self._workbook(workbook_id)
        query: Dict[str, Any] = {}
        if select is not None:
            if isinstance(select, str):
                query["select"] = select
            else:
                query["select"] = ",".join(select)
        return self.request(
            "GET",
            self._range_path(wid, sheet_id, range_address),
            query=query,
            operator_id=operator_id,
            require_operator=True,
        )

    def get_values(
        self,
        sheet_id: str,
        range_address: str,
        workbook_id: Optional[str] = None,
        *,
        operator_id: Optional[str] = None,
    ) -> List[List[Any]]:
        """读取区域 ``values`` 二维数组。"""
        data = self.get_range(
            sheet_id,
            range_address,
            workbook_id,
            select="values",
            operator_id=operator_id,
        )
        values = data.get("values") or []
        if not isinstance(values, list):
            raise DingDiskResponseError("values 不是列表", raw=data)
        return values

    def update_range(
        self,
        sheet_id: str,
        range_address: str,
        workbook_id: Optional[str] = None,
        *,
        values: Optional[CellMatrix] = None,
        operator_id: Optional[str] = None,
        **extra: Any,
    ) -> JsonDict:
        """更新单元格区域 ``PUT .../ranges/{rangeAddress}``。

        可写 ``values`` / ``backgroundColors`` / ``fontSizes`` 等，见官方文档。
        """
        wid = self._workbook(workbook_id)
        body: Dict[str, Any] = dict(extra)
        if values is not None:
            body["values"] = [list(row) for row in values]
        return self.request(
            "PUT",
            self._range_path(wid, sheet_id, range_address),
            body=body,
            operator_id=operator_id,
            require_operator=True,
        )

    def write_values(
        self,
        sheet_id: str,
        values: CellMatrix,
        workbook_id: Optional[str] = None,
        *,
        start_cell: str = "A1",
        operator_id: Optional[str] = None,
        **extra: Any,
    ) -> JsonDict:
        """按起点写入二维数据（自动计算 range）。"""
        range_address = build_a1_range(values, start_cell=start_cell)
        return self.update_range(
            sheet_id,
            range_address,
            workbook_id,
            values=values,
            operator_id=operator_id,
            **extra,
        )

    def append_rows(
        self,
        sheet_id: str,
        values: CellMatrix,
        workbook_id: Optional[str] = None,
        *,
        operator_id: Optional[str] = None,
    ) -> JsonDict:
        """在数据末尾追加行 ``POST .../appendRows``。"""
        wid = self._workbook(workbook_id)
        path = _APPEND_ROWS_PATH.format(
            workbookId=self._encode_path_segment(wid),
            sheetId=self._encode_path_segment(sheet_id),
        )
        return self.request(
            "POST",
            path,
            body={"values": [list(row) for row in values]},
            operator_id=operator_id,
            require_operator=True,
        )

    def clear_data(
        self,
        sheet_id: str,
        range_address: str,
        workbook_id: Optional[str] = None,
        *,
        operator_id: Optional[str] = None,
    ) -> JsonDict:
        """清除区域内数据（保留格式）``POST .../clearData``。"""
        wid = self._workbook(workbook_id)
        path = _CLEAR_DATA_PATH.format(
            workbookId=self._encode_path_segment(wid),
            sheetId=self._encode_path_segment(sheet_id),
            rangeAddress=self._encode_path_segment(range_address, keep_colon=True),
        )
        return self.request("POST", path, operator_id=operator_id, require_operator=True)

    def clear_all(
        self,
        sheet_id: str,
        range_address: str,
        workbook_id: Optional[str] = None,
        *,
        operator_id: Optional[str] = None,
    ) -> JsonDict:
        """清除区域内所有内容（含格式）``POST .../clear``。"""
        wid = self._workbook(workbook_id)
        path = _CLEAR_ALL_PATH.format(
            workbookId=self._encode_path_segment(wid),
            sheetId=self._encode_path_segment(sheet_id),
            rangeAddress=self._encode_path_segment(range_address, keep_colon=True),
        )
        return self.request("POST", path, operator_id=operator_id, require_operator=True)
