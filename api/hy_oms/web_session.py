"""鸿羽 OMS Web 会话：模拟登录拿 Cookie，再访问后台页面接口。

流程与手工登录一致：
1. SOAP ``logOn`` / ``getSsoToken`` → 快捷登录 URL
2. GET 该 URL，由 ``http.cookiejar`` 保存会话 Cookie
3. 带 Cookie 调用后台接口（如退件标签下载）
"""

from __future__ import annotations

import mimetypes
import re
import uuid
from dataclasses import dataclass
from http.cookiejar import CookieJar
from typing import Any, Dict, Mapping, Optional, Tuple
from urllib.error import HTTPError, URLError
from urllib.parse import unquote
from urllib.request import HTTPCookieProcessor, Request, build_opener

from .client import HyOmsClient
from .config import HyOmsConfig
from .exceptions import HyOmsAuthError, HyOmsError


_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36"
)


@dataclass
class DownloadResult:
    """标签下载结果。"""

    return_code: str
    content: bytes
    content_type: str
    filename: str
    status_code: int
    headers: Mapping[str, str]

    @property
    def is_binary(self) -> bool:
        ct = (self.content_type or "").lower()
        if any(x in ct for x in ("zip", "pdf", "octet-stream", "msdownload")):
            return True
        name = (self.filename or "").lower()
        return name.endswith((".zip", ".pdf", ".rar", ".7z"))


def _multipart_form(fields: Mapping[str, str]) -> Tuple[bytes, str]:
    """组装 multipart/form-data（仅普通字段，无文件）。"""
    boundary = f"----HyOmsBoundary{uuid.uuid4().hex}"
    lines: list[bytes] = []
    for name, value in fields.items():
        lines.append(f"--{boundary}".encode("ascii"))
        lines.append(
            f'Content-Disposition: form-data; name="{name}"'.encode("utf-8")
        )
        lines.append(b"")
        lines.append(str(value).encode("utf-8"))
    lines.append(f"--{boundary}--".encode("ascii"))
    lines.append(b"")
    body = b"\r\n".join(lines)
    content_type = f"multipart/form-data; boundary={boundary}"
    return body, content_type


def _filename_from_headers(
    headers: Mapping[str, str],
    *,
    return_code: str,
    content_type: str,
) -> str:
    """从 Content-Disposition / Content-Type 推断保存文件名。"""
    cd = headers.get("Content-Disposition") or headers.get("content-disposition") or ""
    # filename*=UTF-8''xxx 或 filename="xxx"
    m = re.search(r"filename\*\s*=\s*UTF-8''([^;]+)", cd, flags=re.I)
    if m:
        return unquote(m.group(1).strip().strip('"'))
    m = re.search(r'filename\s*=\s*"([^"]+)"', cd, flags=re.I)
    if m:
        return m.group(1).strip()
    m = re.search(r"filename\s*=\s*([^;]+)", cd, flags=re.I)
    if m:
        return unquote(m.group(1).strip().strip('"'))

    ct = (content_type or "").split(";")[0].strip().lower()
    ext = mimetypes.guess_extension(ct) or ""
    if "zip" in ct:
        ext = ".zip"
    elif "pdf" in ct:
        ext = ".pdf"
    if not ext:
        ext = ".bin"
    safe = re.sub(r"[^\w.\-]+", "_", str(return_code).strip()) or "label"
    return f"{safe}{ext}"


class HyOmsWebSession:
    """带 Cookie 的 OMS 后台会话（依赖 ``HyOmsClient.simulate_login``）。"""

    def __init__(
        self,
        client: Optional[HyOmsClient] = None,
        *,
        config: Optional[HyOmsConfig] = None,
    ):
        if client is not None:
            self.client = client
        elif config is not None:
            self.client = HyOmsClient(config)
        else:
            self.client = HyOmsClient.from_config()
        self.config = self.client.config
        self.cookie_jar = CookieJar()
        self.opener = build_opener(HTTPCookieProcessor(self.cookie_jar))
        self.login_url: str = ""
        self._logged_in = False

    @classmethod
    def from_config(cls) -> "HyOmsWebSession":
        return cls(HyOmsClient.from_config())

    @property
    def base_url(self) -> str:
        return self.config.base_url.rstrip("/")

    def get_cookies(self) -> Dict[str, str]:
        return {c.name: c.value for c in self.cookie_jar}

    def login(
        self,
        *,
        user_account: Optional[str] = None,
        user_password: Optional[str] = None,
        company_code: Optional[str] = None,
        mode: str = "logOn",
        **extra: Any,
    ) -> Dict[str, Any]:
        """模拟登录并访问快捷登录 URL，建立 Web Cookie。"""
        result = self.client.simulate_login(
            mode=mode,
            user_account=user_account,
            user_password=user_password,
            company_code=company_code,
            **extra,
        )
        login_url = str(result.get("login_url") or "").strip()
        if not login_url:
            raise HyOmsAuthError("模拟登录成功但未返回 login_url")

        req = Request(
            login_url,
            headers={
                "User-Agent": _UA,
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            },
            method="GET",
        )
        try:
            with self.opener.open(req, timeout=self.config.timeout) as resp:
                # 读完响应以完成重定向与 Set-Cookie
                _ = resp.read()
                final_url = resp.geturl()
        except HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace") if exc.fp else ""
            raise HyOmsAuthError(
                f"访问快捷登录 URL 失败 HTTP {exc.code}: {detail[:300] or exc.reason}"
            ) from exc
        except URLError as exc:
            raise HyOmsAuthError(f"无法访问快捷登录 URL: {exc.reason}") from exc

        if not self.get_cookies():
            raise HyOmsAuthError(
                f"快捷登录后未获得 Cookie（final_url={final_url}）；请检查账号权限或 login_url"
            )

        self.login_url = login_url
        self._logged_in = True
        out = dict(result)
        out["final_url"] = final_url
        out["cookies"] = self.get_cookies()
        return out

    def ensure_login(self, **kwargs: Any) -> None:
        if not self._logged_in:
            self.login(**kwargs)

    def download_label(
        self,
        return_code: str,
        *,
        auto_login: bool = True,
        **login_kwargs: Any,
    ) -> DownloadResult:
        """下载退件标签 ``POST /order/special-orders/download-label``。

        表单字段 ``code`` = 退件单号（与 OMS 后台一致）。
        """
        code = str(return_code or "").strip()
        if not code:
            raise ValueError("download_label 的 return_code 不能为空")
        if auto_login:
            self.ensure_login(**login_kwargs)

        url = f"{self.base_url}/order/special-orders/download-label"
        body, content_type = _multipart_form({"code": code})
        headers = {
            "User-Agent": _UA,
            "Content-Type": content_type,
            "X-Requested-With": "XMLHttpRequest",
            "Referer": f"{self.base_url}/order/special-orders/list",
            "Accept": "*/*",
            "Origin": self.base_url,
        }
        req = Request(url, data=body, headers=headers, method="POST")
        try:
            with self.opener.open(req, timeout=self.config.timeout) as resp:
                content = resp.read()
                status = getattr(resp, "status", None) or resp.getcode() or 200
                hdrs = {k: v for k, v in resp.headers.items()}
        except HTTPError as exc:
            detail = exc.read() if exc.fp else b""
            raise HyOmsError(
                f"下载标签失败 HTTP {exc.code}: "
                f"{detail[:500].decode('utf-8', errors='replace') or exc.reason}"
            ) from exc
        except URLError as exc:
            raise HyOmsError(f"下载标签无法连接: {exc.reason}") from exc

        resp_ct = hdrs.get("Content-Type") or hdrs.get("content-type") or ""
        filename = _filename_from_headers(hdrs, return_code=code, content_type=resp_ct)
        result = DownloadResult(
            return_code=code,
            content=content,
            content_type=resp_ct,
            filename=filename,
            status_code=int(status),
            headers=hdrs,
        )
        if not result.is_binary:
            # 常见失败：返回 HTML/JSON 错误页
            preview = content[:800].decode("utf-8", errors="replace")
            raise HyOmsError(
                f"下载标签未返回文件（Content-Type={resp_ct!r}）: {preview}"
            )
        return result
