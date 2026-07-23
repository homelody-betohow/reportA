"""4PX（递四方 / FPX）开放平台 HTTP 客户端。

协议：POST ``{host}/router/api/service?method=...&app_key=...&sign=...``
业务参数放 JSON body；签名为 MD5(公共参数按名字母序拼接 + body + app_secret)。

文档入口（直发服务 → 费用查询 → 查询订单费用信息）::
    https://open.4px.com/v2/doc/detail?ids=55,88,214
    method = ds.xms.order.getFreight
"""

from __future__ import annotations

import hashlib
import json
import time
from typing import Any, Dict, Mapping, MutableMapping, Optional
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from .config import FpxConfig
from .exceptions import FpxAuthError, FpxError, FpxResponseError

Params = Optional[Mapping[str, Any]]
JsonDict = Dict[str, Any]

# 签名时排除的公共参数（官方：不含 access_token、language）
_SIGN_EXCLUDE = frozenset({"sign", "access_token", "language"})


def build_sign(
    *,
    app_key: str,
    app_secret: str,
    method: str,
    timestamp: str,
    body: str,
    version: str,
    fmt: str = "json",
) -> str:
    """按官方规则生成 32 位小写 MD5 签名。

    拼接顺序：公共参数名按字母升序 ``key+value``，再接 body JSON，再接 app_secret。
    """
    params = {
        "app_key": app_key,
        "format": fmt,
        "method": method,
        "timestamp": str(timestamp),
        "v": version,
    }
    parts = "".join(f"{k}{params[k]}" for k in sorted(params))
    raw = f"{parts}{body}{app_secret}"
    return hashlib.md5(raw.encode("utf-8")).hexdigest()


def _compact_json(payload: Any) -> str:
    if payload is None:
        return "{}"
    if isinstance(payload, str):
        return payload if payload.strip() else "{}"
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def _normalize_data(data: Any) -> Any:
    """部分接口文档示例里 data 为 JSON 字符串，运行时也可能直接返回对象。"""
    if isinstance(data, str):
        text = data.strip()
        if text.startswith("{") or text.startswith("["):
            try:
                return json.loads(text)
            except json.JSONDecodeError:
                return data
    return data


class FpxClient:
    """4PX 开放平台通用客户端。

    任意接口均可 ``call("ds.xms.order.getFreight", {"request_no": "..."})``；
    直发费用查询另有语义化封装 ``get_freight``。
    """

    def __init__(self, config: FpxConfig):
        self.config = config

    @classmethod
    def from_config(cls) -> "FpxClient":
        return cls(FpxConfig.default())

    @classmethod
    def from_env(cls, **_kwargs) -> "FpxClient":
        return cls.from_config()

    # ------------------------------------------------------------------
    # 底层调用
    # ------------------------------------------------------------------

    def call(
        self,
        method: str,
        params: Params = None,
        *,
        version: Optional[str] = None,
        raise_on_failure: bool = True,
    ) -> JsonDict:
        """调用任意 4PX method，返回解析后的 JSON dict。"""
        cfg = self.config
        v = version or cfg.api_version
        body = _compact_json(params if params is not None else {})
        timestamp = str(int(time.time() * 1000))
        sign = build_sign(
            app_key=cfg.app_key,
            app_secret=cfg.app_secret,
            method=method,
            timestamp=timestamp,
            body=body,
            version=v,
            fmt=cfg.format,
        )
        query: MutableMapping[str, str] = {
            "method": method,
            "app_key": cfg.app_key,
            "v": v,
            "timestamp": timestamp,
            "format": cfg.format,
            "sign": sign,
            "language": cfg.language,
        }
        if (cfg.access_token or "").strip():
            query["access_token"] = cfg.access_token.strip()

        url = f"{cfg.service_url}?{urlencode(query)}"
        req = Request(
            url,
            data=body.encode("utf-8"),
            headers={
                "Content-Type": "application/json; charset=utf-8",
                "Accept": "application/json",
            },
            method="POST",
        )
        try:
            with urlopen(req, timeout=cfg.timeout) as resp:
                raw_text = resp.read().decode("utf-8", errors="replace")
        except HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace") if exc.fp else ""
            raise FpxError(f"HTTP {exc.code}: {detail or exc.reason}") from exc
        except URLError as exc:
            raise FpxError(f"无法连接 4PX 开放平台: {exc.reason}") from exc

        try:
            data = json.loads(raw_text)
        except json.JSONDecodeError as exc:
            raise FpxResponseError(f"响应不是合法 JSON: {raw_text[:200]}", raw=raw_text) from exc

        if not isinstance(data, dict):
            raise FpxResponseError("响应 JSON 不是对象", raw=data)

        if "data" in data:
            data = dict(data)
            data["data"] = _normalize_data(data.get("data"))

        result = str(data.get("result", "")).strip()
        if raise_on_failure and result != "1":
            errors = data.get("errors") or []
            err_code = None
            err_msg = None
            if isinstance(errors, list) and errors:
                first = errors[0] if isinstance(errors[0], dict) else {}
                err_code = first.get("error_code") or first.get("errorCode")
                err_msg = first.get("error_msg") or first.get("errorMsg")
            message = err_msg or data.get("msg") or f"result={result or 'empty'}"
            auth_hints = ("app_key", "app_secret", "sign", "access_token", "签名", "鉴权", "授权")
            if any(h in str(message).lower() for h in auth_hints) or str(err_code) in {
                "100001",
                "100002",
                "100003",
            }:
                raise FpxAuthError(str(message))
            raise FpxResponseError(str(message), err_code=err_code, raw=data)
        return data

    # ------------------------------------------------------------------
    # 直发服务 · 费用查询（ids=55,88,214）
    # ------------------------------------------------------------------

    def get_freight(self, request_no: str, **extra) -> JsonDict:
        """查询订单费用信息 ``ds.xms.order.getFreight``。

        ``request_no`` 可为 4PX 单号、服务商单号或客户单号。
        成功时 ``data`` 含 total_fee / charge_weight / currency / subs 等。
        """
        params: Dict[str, Any] = {"request_no": request_no, **extra}
        return self.call("ds.xms.order.getFreight", params, version="1.0.0")

    # ------------------------------------------------------------------
    # 直发服务 · 委托单 / 面单 / 轨迹
    # ------------------------------------------------------------------

    def get_order(
        self,
        request_no: Optional[str] = None,
        *,
        start_time_of_create_consignment: Optional[int] = None,
        end_time_of_create_consignment: Optional[int] = None,
        consignment_status: Optional[str] = None,
        **extra,
    ) -> JsonDict:
        """查询直发委托单 ``ds.xms.order.get``。"""
        params: Dict[str, Any] = dict(extra)
        if request_no is not None:
            params["request_no"] = request_no
        if start_time_of_create_consignment is not None:
            params["start_time_of_create_consignment"] = start_time_of_create_consignment
        if end_time_of_create_consignment is not None:
            params["end_time_of_create_consignment"] = end_time_of_create_consignment
        if consignment_status is not None:
            params["consignment_status"] = consignment_status
        return self.call("ds.xms.order.get", params, version="1.1.0")

    def cancel_order(self, request_no: str, cancel_reason: str, **extra) -> JsonDict:
        """取消直发委托单 ``ds.xms.order.cancel``。"""
        params = {"request_no": request_no, "cancel_reason": cancel_reason, **extra}
        return self.call("ds.xms.order.cancel", params, version="1.0.0")

    def get_label(
        self,
        request_no: str,
        *,
        response_label_format: Optional[str] = None,
        **extra,
    ) -> JsonDict:
        """获取面单标签 ``ds.xms.label.get``。"""
        params: Dict[str, Any] = {"request_no": request_no, **extra}
        if response_label_format is not None:
            params["response_label_format"] = response_label_format
        return self.call("ds.xms.label.get", params, version="1.1.0")

    def get_tracking(self, delivery_order_no: str, **extra) -> JsonDict:
        """物流轨迹查询 ``tr.order.tracking.get``。"""
        params = {"deliveryOrderNo": delivery_order_no, **extra}
        return self.call("tr.order.tracking.get", params, version="1.0.0")

    # ------------------------------------------------------------------
    # 公共服务 · 费用试算（ids=54,73,144）
    # ------------------------------------------------------------------

    def price_calculator(
        self,
        *,
        service_code: str,
        warehouse_code: str,
        weight: float,
        length: float,
        width: float,
        height: float,
        destination: Mapping[str, Any],
        billing_time: Optional[int] = None,
        product_codes: Optional[list] = None,
        cargo_units: Optional[list] = None,
        sku_list: Optional[list] = None,
        **extra,
    ) -> JsonDict:
        """费用试算 ``com.css.price_calculator``。

        文档：https://open.4px.com/v2/doc/detail?ids=54,73,144

        必填：``service_code``（目前仅 FB4）、仓库、重量(g)、长宽高(cm)、目的地。
        ``billing_time`` 为毫秒时间戳，默认当前时间。
        仓内按 SKU 试算时传 ``sku_list``（``sku_code`` / ``sku_qty``）。
        """
        params: Dict[str, Any] = {
            "service_code": service_code,
            "warehouse_code": warehouse_code,
            "weight": weight,
            "length": length,
            "width": width,
            "height": height,
            "destination": dict(destination),
            "billing_time": int(billing_time if billing_time is not None else time.time() * 1000),
            **extra,
        }
        if product_codes is not None:
            params["product_codes"] = list(product_codes)
        if cargo_units is not None:
            params["cargo_units"] = list(cargo_units)
        if sku_list is not None:
            params["skuList"] = list(sku_list)
        return self.call("com.css.price_calculator", params, version="1.0.0")

    def get_billing(
        self,
        *,
        business_type: str,
        order_no: Optional[str] = None,
        ref_no: Optional[str] = None,
        **extra,
    ) -> JsonDict:
        """费用查询 ``com.basis.billing.getbilling``。

        文档：https://open.4px.com/v2/doc/detail?ids=54,74,159

        ``business_type``：``I`` 入库委托 / ``O`` 出库委托 / ``T`` 调拨委托 / ``L`` 尾程管理运单。
        ``order_no``（业务单号）与 ``ref_no``（参考号）二选一；同时传入时以业务单号为准。
        """
        if not (order_no or "").strip() and not (ref_no or "").strip():
            raise ValueError("get_billing 需要 order_no 或 ref_no 之一")
        params: Dict[str, Any] = {"business_type": business_type, **extra}
        if order_no is not None and str(order_no).strip():
            params["order_no"] = str(order_no).strip()
        if ref_no is not None and str(ref_no).strip():
            params["ref_no"] = str(ref_no).strip()
        return self.call("com.basis.billing.getbilling", params, version="1.0.0")

    def get_warehouse_list(self, params: Params = None) -> JsonDict:
        """查询仓库信息 ``com.basis.warehouse.getlist``。"""
        return self.call("com.basis.warehouse.getlist", params or {}, version="1.0.0")

    def get_logistics_product_list(self, params: Params = None) -> JsonDict:
        """查询物流产品 ``com.basis.logistics_product.getlist``。"""
        return self.call("com.basis.logistics_product.getlist", params or {}, version="1.0.0")
