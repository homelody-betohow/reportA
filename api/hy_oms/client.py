"""鸿羽 OMS SOAP 客户端。"""

from __future__ import annotations

import json
import xml.etree.ElementTree as ET
from typing import Any, Dict, Iterable, Iterator, List, Mapping, MutableMapping, Optional, Sequence, Union
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
from xml.sax.saxutils import escape

from .config import HyOmsConfig
from .exceptions import HyOmsAuthError, HyOmsError, HyOmsResponseError

Params = Optional[Mapping[str, Any]]
JsonDict = Dict[str, Any]


def _xml_text(value: str) -> str:
    return escape(value, {"'": "&apos;", '"': "&quot;"})


def _build_soap_envelope(
    *,
    app_token: str,
    app_key: str,
    service: str,
    params_json: str,
    language: str,
) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<SOAP-ENV:Envelope xmlns:SOAP-ENV="http://schemas.xmlsoap.org/soap/envelope/" '
        'xmlns:ns1="http://www.example.org/Ec/">'
        "<SOAP-ENV:Body>"
        "<ns1:callService>"
        f"<paramsJson>{_xml_text(params_json)}</paramsJson>"
        f"<appToken>{_xml_text(app_token)}</appToken>"
        f"<appKey>{_xml_text(app_key)}</appKey>"
        f"<service>{_xml_text(service)}</service>"
        f"<language>{_xml_text(language)}</language>"
        "</ns1:callService>"
        "</SOAP-ENV:Body>"
        "</SOAP-ENV:Envelope>"
    )


def _extract_response_text(soap_xml: str) -> str:
    root = ET.fromstring(soap_xml)
    # 忽略命名空间差异，找名为 response 的节点
    for node in root.iter():
        tag = node.tag.rsplit("}", 1)[-1]
        if tag == "response" and node.text is not None:
            return node.text.strip()
    raise HyOmsResponseError("SOAP 响应中未找到 <response> 节点", raw=soap_xml)


def _as_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    return str(value).strip().lower() in {"1", "true", "yes", "y"}


class HyOmsClient:
    """鸿羽海外仓 OMS 通用客户端。

    任意接口均可 ``call("serviceName", {...})``；
    报表常用接口另有语义化封装（仓租、库存、产品、订单等）。
    """

    def __init__(self, config: HyOmsConfig):
        self.config = config

    @classmethod
    def from_config(cls) -> "HyOmsClient":
        """使用 ``config.py`` 中的 APP_TOKEN / APP_KEY。"""
        return cls(HyOmsConfig.default())

    # 兼容旧调用名
    @classmethod
    def from_env(cls, **_kwargs) -> "HyOmsClient":
        return cls.from_config()

    # ------------------------------------------------------------------
    # 底层调用
    # ------------------------------------------------------------------

    def call(
        self,
        service: str,
        params: Params = None,
        *,
        raise_on_failure: bool = True,
    ) -> JsonDict:
        """调用任意 OMS service，返回解析后的 JSON dict。"""
        payload = "" if params is None else json.dumps(params, ensure_ascii=False, separators=(",", ":"))
        envelope = _build_soap_envelope(
            app_token=self.config.app_token,
            app_key=self.config.app_key,
            service=service,
            params_json=payload,
            language=self.config.language,
        )
        req = Request(
            self.config.service_url,
            data=envelope.encode("utf-8"),
            headers={
                "Content-Type": "text/xml; charset=utf-8",
                "SOAPAction": "http://www.example.org/Ec/callService",
            },
            method="POST",
        )
        try:
            with urlopen(req, timeout=self.config.timeout) as resp:
                body = resp.read().decode("utf-8", errors="replace")
        except HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace") if exc.fp else ""
            raise HyOmsError(f"HTTP {exc.code}: {detail or exc.reason}") from exc
        except URLError as exc:
            raise HyOmsError(f"无法连接鸿羽 OMS: {exc.reason}") from exc

        text = _extract_response_text(body)
        try:
            data = json.loads(text)
        except json.JSONDecodeError as exc:
            raise HyOmsResponseError(f"响应不是合法 JSON: {text[:200]}", raw=body) from exc

        if not isinstance(data, dict):
            raise HyOmsResponseError("响应 JSON 不是对象", raw=data)

        ask = str(data.get("ask", "")).strip()
        if raise_on_failure and ask.lower() != "success":
            err = data.get("Error") or {}
            err_code = err.get("errCode") if isinstance(err, dict) else None
            message = (
                (err.get("errMessage") if isinstance(err, dict) else None)
                or data.get("message")
                or ask
                or "未知错误"
            )
            if str(err_code) == "50001" or "appToken" in str(message) or "appKey" in str(message):
                raise HyOmsAuthError(str(message))
            raise HyOmsResponseError(str(message), err_code=err_code, raw=data)
        return data

    def iter_pages(
        self,
        service: str,
        params: Params = None,
        *,
        page_size: int = 100,
        start_page: int = 1,
        max_pages: Optional[int] = None,
        data_key: str = "data",
    ) -> Iterator[Any]:
        """按页拉取列表接口，逐条 yield ``data`` 中的元素。

        终止条件（任一满足即停止）：
        - ``nextPage`` 为 false / 缺失且本页无数据
        - 本页条数 < pageSize
        - 已累计条数达到 ``count`` / ``total``
        """
        base: MutableMapping[str, Any] = dict(params or {})
        page = start_page
        fetched = 0
        pages = 0

        while True:
            if max_pages is not None and pages >= max_pages:
                break
            query = dict(base)
            query["page"] = page
            query["pageSize"] = page_size
            result = self.call(service, query)
            rows = result.get(data_key) or []
            if isinstance(rows, dict):
                rows = [rows]
            if not isinstance(rows, list):
                raise HyOmsResponseError(f"{service} 的 {data_key} 不是列表", raw=result)

            for row in rows:
                yield row
            fetched += len(rows)
            pages += 1

            total = result.get("count", result.get("total"))
            try:
                total_n = int(total) if total is not None and str(total).strip() != "" else None
            except (TypeError, ValueError):
                total_n = None

            next_page = result.get("nextPage")
            if next_page is not None and not _as_bool(next_page):
                break
            if total_n is not None and fetched >= total_n:
                break
            if len(rows) == 0 or len(rows) < page_size:
                break
            page += 1

    def fetch_all(
        self,
        service: str,
        params: Params = None,
        *,
        page_size: int = 100,
        **kwargs,
    ) -> List[Any]:
        """``iter_pages`` 的列表包装。"""
        return list(self.iter_pages(service, params, page_size=page_size, **kwargs))

    # ------------------------------------------------------------------
    # 基础数据
    # ------------------------------------------------------------------

    def get_warehouse(self, params: Params = None) -> JsonDict:
        """系统仓库 ``getWarehouse``。"""
        return self.call("getWarehouse", params or {})

    def get_shipping_method(self, params: Params = None) -> JsonDict:
        return self.call("getShippingMethod", params or {})

    def get_fee_type(self, params: Params = None) -> JsonDict:
        return self.call("getFeeType", params or {})

    def get_currency(self, params: Params = None) -> JsonDict:
        return self.call("getCurrency", params or {})

    # ------------------------------------------------------------------
    # 产品 / 库存
    # ------------------------------------------------------------------

    def get_product_list(self, params: Params = None, *, all_pages: bool = False, page_size: int = 100):
        """产品列表 ``getProductList``。"""
        if all_pages:
            return self.fetch_all("getProductList", params, page_size=page_size)
        return self.call("getProductList", params or {"page": 1, "pageSize": page_size})

    def get_product_inventory(
        self,
        params: Params = None,
        *,
        all_pages: bool = False,
        page_size: int = 100,
    ):
        """产品库存 ``getProductInventory``。"""
        if all_pages:
            return self.fetch_all("getProductInventory", params, page_size=page_size)
        p = dict(params or {})
        p.setdefault("page", 1)
        p.setdefault("pageSize", page_size)
        return self.call("getProductInventory", p)

    def get_inventory_storage(
        self,
        charge_date: str,
        *,
        page: int = 1,
        page_size: int = 50,
        all_pages: bool = False,
    ):
        """库存快照 ``getWhInventoryStorage``（按计费日）。

        返回 SKU/体积/库龄等，不含金额；金额类仓租见 ``get_storage_costs`` / 费用流水。
        """
        params = {"chargeDate": charge_date, "page": page, "pageSize": page_size}
        if all_pages:
            return self.fetch_all("getWhInventoryStorage", {"chargeDate": charge_date}, page_size=page_size)
        return self.call("getWhInventoryStorage", params)

    # ------------------------------------------------------------------
    # 费用 / 仓租
    # ------------------------------------------------------------------

    def get_storage_costs(
        self,
        *,
        date_for: Optional[str] = None,
        date_to: Optional[str] = None,
        warehouse_id: Optional[int] = None,
        page: int = 1,
        page_size: int = 100,
        all_pages: bool = False,
        **extra,
    ):
        """仓租列表 ``getStorageCosts``（账单级汇总，非 SKU 明细）。"""
        params: Dict[str, Any] = {"page": page, "pageSize": page_size, **extra}
        if date_for is not None:
            params["dateFor"] = date_for
        if date_to is not None:
            params["dateTo"] = date_to
        if warehouse_id is not None:
            params["warehouseId"] = warehouse_id
        if all_pages:
            base = {k: v for k, v in params.items() if k not in {"page", "pageSize"}}
            return self.fetch_all("getStorageCosts", base, page_size=page_size)
        return self.call("getStorageCosts", params)

    def get_cost_water(
        self,
        params: Params = None,
        *,
        all_pages: bool = False,
        page_size: int = 100,
    ):
        """费用流水 ``getCostWater``。"""
        if all_pages:
            return self.fetch_all("getCostWater", params, page_size=page_size)
        p = dict(params or {})
        p.setdefault("page", 1)
        p.setdefault("pageSize", page_size)
        return self.call("getCostWater", p)

    def get_billing_detail(
        self,
        params: Params = None,
        *,
        all_pages: bool = False,
        page_size: int = 100,
    ):
        """费用账单 ``getBillingDetail``。"""
        if all_pages:
            return self.fetch_all("getBillingDetail", params, page_size=page_size)
        p = dict(params or {})
        p.setdefault("page", 1)
        p.setdefault("pageSize", page_size)
        return self.call("getBillingDetail", p)

    def get_balance(self, params: Params = None) -> JsonDict:
        return self.call("getBalance", params or {})

    def get_calculate_fee(
        self,
        *,
        warehouse_code: str,
        country_code: str,
        shipping_method: str,
        weight: float,
        postcode: Optional[str] = None,
        length: Optional[float] = None,
        width: Optional[float] = None,
        height: Optional[float] = None,
        city: Optional[str] = None,
        state: Optional[str] = None,
        address1: Optional[str] = None,
        name: Optional[str] = None,
        phone: Optional[str] = None,
        self_lifting: Optional[int] = None,
        items: Optional[List[Mapping[str, Any]]] = None,
        **extra,
    ) -> JsonDict:
        """运费试算 ``getCalculateFee``。

        必填：仓库、目的国、配送方式、包裹重量。
        返回 ``data`` 含 SHIPPING / OPF / totalFee / currency_code 等费用明细。
        """
        params: Dict[str, Any] = {
            "warehouse_code": warehouse_code,
            "country_code": country_code,
            "shipping_method": shipping_method,
            "weight": weight,
            **extra,
        }
        optional = {
            "postcode": postcode,
            "length": length,
            "width": width,
            "height": height,
            "city": city,
            "state": state,
            "address1": address1,
            "name": name,
            "phone": phone,
            "self_lifting": self_lifting,
            "items": items,
        }
        for key, value in optional.items():
            if value is not None:
                params[key] = value
        return self.call("getCalculateFee", params)

    def get_calculate_fee_batch(
        self,
        *,
        warehouse_code: str,
        country_code: str,
        shipping_method: Union[str, Sequence[str]],
        weight: float,
        postcode: Optional[str] = None,
        length: Optional[float] = None,
        width: Optional[float] = None,
        height: Optional[float] = None,
        state: Optional[str] = None,
        address1: Optional[str] = None,
        self_lifting: Optional[int] = None,
        items: Optional[List[Mapping[str, Any]]] = None,
        **extra,
    ) -> JsonDict:
        """批量运费试算 ``getCalculateFeeBatch``。

        ``shipping_method`` 为配送方式代码列表（也可传单个字符串，会自动包成列表）。
        """
        methods: List[str]
        if isinstance(shipping_method, str):
            methods = [shipping_method]
        else:
            methods = list(shipping_method)
        params: Dict[str, Any] = {
            "warehouse_code": warehouse_code,
            "country_code": country_code,
            "shipping_method": methods,
            "weight": weight,
            **extra,
        }
        optional = {
            "postcode": postcode,
            "length": length,
            "width": width,
            "height": height,
            "state": state,
            "address1": address1,
            "self_lifting": self_lifting,
            "items": items,
        }
        for key, value in optional.items():
            if value is not None:
                params[key] = value
        return self.call("getCalculateFeeBatch", params)

    # ------------------------------------------------------------------
    # 订单
    # ------------------------------------------------------------------

    def get_order_list(
        self,
        params: Params = None,
        *,
        all_pages: bool = False,
        page_size: int = 100,
    ):
        """订单列表 ``getOrderList``。"""
        if all_pages:
            return self.fetch_all("getOrderList", params, page_size=page_size)
        p = dict(params or {})
        p.setdefault("page", 1)
        p.setdefault("pageSize", page_size)
        return self.call("getOrderList", p)

    def get_order_by_code(self, order_code: str, **extra) -> JsonDict:
        return self.call("getOrderByCode", {"order_code": order_code, **extra})

    def get_order_by_ref_code(self, reference_no: Union[str, Iterable[str]], **extra) -> JsonDict:
        if not isinstance(reference_no, str):
            reference_no = list(reference_no)
        return self.call("getOrderByRefCode", {"reference_no": reference_no, **extra})

    # ------------------------------------------------------------------
    # 退件
    # ------------------------------------------------------------------

    def get_special_orders_list(
        self,
        params: Params = None,
        *,
        all_pages: bool = False,
        page_size: int = 100,
    ):
        """退件列表 ``getSpecialOrdersList``。"""
        if all_pages:
            return self.fetch_all("getSpecialOrdersList", params, page_size=page_size)
        p = dict(params or {})
        p.setdefault("page", 1)
        p.setdefault("pageSize", page_size)
        return self.call("getSpecialOrdersList", p)

    def create_return_bill(
        self,
        *,
        warehouse_code: str,
        items: Sequence[Mapping[str, Any]],
        tracking_no: Optional[Union[str, Sequence[str]]] = None,
        return_type: Optional[str] = None,
        verify: Optional[Union[int, str]] = None,
        reference_no: Optional[str] = None,
        order_code: Optional[str] = None,
        claim_code: Optional[str] = None,
        expected_date: Optional[str] = None,
        return_desc: Optional[str] = None,
        operation_desc: Optional[str] = None,
        buyer_name: Optional[str] = None,
        buyers_ein: Optional[str] = None,
        seller_store: Optional[str] = None,
        images: Optional[Sequence[Mapping[str, Any]]] = None,
        return_identification: Optional[int] = None,
        sm_code: Optional[str] = None,
        sender_info: Optional[Mapping[str, Any]] = None,
        **extra,
    ) -> JsonDict:
        """创建退件 ``createReturnBill``。

        标准退件：``tracking_no`` + ``return_type``(S/L/C) + ``items``；
        ``S`` 需 ``order_code``，``C`` 需 ``claim_code``。

        回邮退件：``return_identification=1``，另需 ``reference_no`` / ``sm_code`` / ``sender_info``。

        成功响应含 ``return_code``（退件单号）。
        """
        if not items:
            raise ValueError("createReturnBill 的 items 不能为空")

        params: Dict[str, Any] = {
            "warehouse_code": warehouse_code,
            "items": [dict(x) for x in items],
            **extra,
        }
        if isinstance(tracking_no, (list, tuple)):
            tracking_no = list(tracking_no)
        optional = {
            "tracking_no": tracking_no,
            "return_type": return_type,
            "verify": verify,
            "reference_no": reference_no,
            "order_code": order_code,
            "claim_code": claim_code,
            "expected_date": expected_date,
            "return_desc": return_desc,
            "operation_desc": operation_desc,
            "buyer_name": buyer_name,
            "buyers_ein": buyers_ein,
            "seller_store": seller_store,
            "images": [dict(x) for x in images] if images is not None else None,
            "return_identification": return_identification,
            "sm_code": sm_code,
            "sender_info": dict(sender_info) if sender_info is not None else None,
        }
        for key, value in optional.items():
            if value is not None:
                params[key] = value
        return self.call("createReturnBill", params)
