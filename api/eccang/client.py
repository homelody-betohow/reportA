from __future__ import annotations

import hashlib
import json
import random
import string
import time
import urllib.error
import urllib.request
from typing import Any

from .config import EccangConfig
from .exceptions import EccangApiError

# 官方默认签名字段顺序（docId=313）
# https://open.eccang.com/#/documentCenter?docId=313&catId=0-173-173,0-171
DEFAULT_SIGN_ORDER: tuple[str, ...] = (
    "app_key",
    "biz_content",
    "charset",
    "interface_method",
    "nonce_str",
    "service_id",
    "sign_type",
    "timestamp",
    "version",
)


class EccangClient:
    """易仓 ERP 开放平台 HTTP 客户端。

    签名规则（官方文档 docId=313）：
    1. 取参与签名的参数（不含 sign），按参数名 ASCII 升序排序
    2. 拼接为 key1=value1&key2=value2&...
    3. 在末尾直接追加 app_secret（不加 &）
    4. 对拼接串做 MD5，取 32 位小写十六进制作为 sign

    参考：https://open.eccang.com/#/documentCenter?docId=313&catId=0-173-173,0-171
    """

    def __init__(self, config: EccangConfig | None = None) -> None:
        self.config = config or EccangConfig.default()

    @staticmethod
    def generate_nonce_str(length: int = 16) -> str:
        """生成随机字符串。"""
        return "".join(random.choices(string.ascii_letters + string.digits, k=length))

    @staticmethod
    def compact_json(body: dict[str, Any] | list[Any] | None) -> str:
        """将请求体转为紧凑 JSON 字符串。"""
        payload = body if body is not None else {}
        return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))

    @classmethod
    def build_sign(
        cls,
        params: dict[str, str],
        app_secret: str,
        *,
        sign_order: list[str] | None = None,
        uppercase: bool = False,
    ) -> str:
        """构建 MD5 签名。"""
        if sign_order:
            keys = [k for k in sign_order if k in params and k != "sign"]
        else:
            keys = sorted(k for k in params.keys() if k != "sign")

        auth_str = "&".join(f"{key}={params[key]}" for key in keys)
        auth_str += app_secret
        digest = hashlib.md5(auth_str.encode("utf-8")).hexdigest()
        return digest.upper() if uppercase else digest.lower()

    @staticmethod
    def _parse_biz_content(payload: dict[str, Any]) -> None:
        """将响应中的 biz_content JSON 字符串解析为 data 字段。"""
        biz = payload.get("biz_content")
        if not isinstance(biz, str) or not biz.strip():
            return
        try:
            parsed = json.loads(biz)
        except json.JSONDecodeError:
            return
        payload["data"] = parsed

    @staticmethod
    def _is_success(payload: dict[str, Any]) -> bool:
        code = str(payload.get("code", ""))
        ask = str(payload.get("ask", ""))
        message = str(payload.get("message") or payload.get("msg") or "")
        if code in {"0", "200"}:
            return True
        if ask.lower() == "success":
            return True
        # 部分接口成功时 message 为「操作成功」
        return message in {"操作成功", "请求成功"}

    def call(
        self,
        method: str,
        body: dict[str, Any] | list[Any] | None = None,
        *,
        version: str | None = None,
        charset: str | None = None,
        sign_order: list[str] | None = None,
    ) -> dict[str, Any]:
        """调用任意易仓 ERP 接口，返回解析后的响应字典。"""
        body_json = self.compact_json(body)
        nonce_str = self.generate_nonce_str()
        timestamp = str(int(time.time() * 1000))
        api_version = version or self.config.version
        api_charset = charset or self.config.charset

        params: dict[str, str] = {
            "app_key": self.config.app_key,
            "biz_content": body_json,
            "charset": api_charset,
            "interface_method": method,
            "nonce_str": nonce_str,
            "service_id": self.config.service_id,
            "sign_type": self.config.sign_type,
            "timestamp": timestamp,
            "version": api_version,
        }

        sign = self.build_sign(
            params,
            self.config.app_secret,
            sign_order=sign_order or list(DEFAULT_SIGN_ORDER),
        )
        params["sign"] = sign

        req = urllib.request.Request(
            self.config.base_url,
            data=json.dumps(params, ensure_ascii=False).encode("utf-8"),
            headers={
                "Content-Type": "application/json; charset=UTF-8",
                "Accept": "application/json",
            },
            method="POST",
        )

        try:
            with urllib.request.urlopen(req, timeout=self.config.timeout) as resp:
                raw = resp.read().decode("utf-8")
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")
            raise EccangApiError(
                f"易仓 ERP HTTP {exc.code}: {detail[:500]}",
                method=method,
            ) from exc
        except urllib.error.URLError as exc:
            raise EccangApiError(
                f"易仓 ERP 网络错误：{exc.reason}",
                method=method,
            ) from exc

        try:
            payload = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise EccangApiError(
                f"易仓 ERP 响应非 JSON：{raw[:500]}",
                method=method,
            ) from exc

        self._parse_biz_content(payload)

        if not self._is_success(payload):
            message = str(payload.get("message") or payload.get("msg") or "易仓 ERP 接口调用失败")
            raise EccangApiError(
                message,
                code=str(payload.get("code", "")),
                ask=str(payload.get("ask", "")),
                method=method,
                raw_payload=payload,
                raw_text=raw,
            )

        return payload
