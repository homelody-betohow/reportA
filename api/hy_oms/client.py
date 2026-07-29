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

    def create_product(
        self,
        *,
        product_sku: str,
        product_title: str,
        product_weight: float,
        product_length: float,
        product_width: float,
        product_height: float,
        product_declared_value: float,
        product_declared_name: str,
        reference_no: Optional[str] = None,
        product_title_en: Optional[str] = None,
        product_net_weight: Optional[float] = None,
        contain_battery: Optional[int] = None,
        battery_type: Optional[str] = None,
        product_declared_name_zh: Optional[str] = None,
        hs_code: Optional[str] = None,
        cat_lang: Optional[str] = None,
        warning_qty: Optional[int] = None,
        warning_days: Optional[int] = None,
        product_brand: Optional[str] = None,
        product_model: Optional[str] = None,
        product_origin: Optional[str] = None,
        product_material: Optional[str] = None,
        product_use_en: Optional[str] = None,
        product_material_en: Optional[str] = None,
        product_desc_url: Optional[str] = None,
        cat_id_level0: Optional[int] = None,
        cat_id_level1: Optional[int] = None,
        cat_id_level2: Optional[int] = None,
        verify: Optional[int] = None,
        customer_img: Optional[Mapping[str, Any]] = None,
        product_color: Optional[str] = None,
        shared_product: Optional[int] = None,
        shared_unit_price: Optional[Union[str, float]] = None,
        product_description: Optional[str] = None,
        is_box_more_sku: Optional[int] = None,
        fragile_property: Optional[int] = None,
        product_size_type: Optional[str] = None,
        is_batch_tag: Optional[int] = None,
        ean: Optional[str] = None,
        ncm: Optional[str] = None,
        cest: Optional[str] = None,
        sku_sort_code: Optional[str] = None,
        is_serialized: Optional[int] = None,
        **extra,
    ) -> JsonDict:
        """创建产品 ``createProduct``。

        必填：SKU、标题、重量、长宽高、申报价值、申报英文名。
        ``verify``：0/缺省=草稿，1=正式产品。成功响应含 ``product_sku``。
        """
        from api.hy_oms.request.create_product import build_params

        params = build_params(
            product_sku=product_sku,
            product_title=product_title,
            product_weight=product_weight,
            product_length=product_length,
            product_width=product_width,
            product_height=product_height,
            product_declared_value=product_declared_value,
            product_declared_name=product_declared_name,
            reference_no=reference_no,
            product_title_en=product_title_en,
            product_net_weight=product_net_weight,
            contain_battery=contain_battery,
            battery_type=battery_type,
            product_declared_name_zh=product_declared_name_zh,
            hs_code=hs_code,
            cat_lang=cat_lang,
            warning_qty=warning_qty,
            warning_days=warning_days,
            product_brand=product_brand,
            product_model=product_model,
            product_origin=product_origin,
            product_material=product_material,
            product_use_en=product_use_en,
            product_material_en=product_material_en,
            product_desc_url=product_desc_url,
            cat_id_level0=cat_id_level0,
            cat_id_level1=cat_id_level1,
            cat_id_level2=cat_id_level2,
            verify=verify,
            customer_img=customer_img,
            product_color=product_color,
            shared_product=shared_product,
            shared_unit_price=shared_unit_price,
            product_description=product_description,
            is_box_more_sku=is_box_more_sku,
            fragile_property=fragile_property,
            product_size_type=product_size_type,
            is_batch_tag=is_batch_tag,
            ean=ean,
            ncm=ncm,
            cest=cest,
            sku_sort_code=sku_sort_code,
            is_serialized=is_serialized,
            **extra,
        )
        return self.call("createProduct", params)

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

    def get_return_bill(
        self,
        *,
        return_code: str,
        **extra,
    ) -> JsonDict:
        """获取退件详情 ``getReturnBill``。

        必填 ``return_code``（退件单号）。成功响应 ``data`` 含状态、跟踪号、明细等。
        """
        code = str(return_code or "").strip()
        if not code:
            raise ValueError("getReturnBill 的 return_code 不能为空")
        params: Dict[str, Any] = {"return_code": code, **extra}
        return self.call("getReturnBill", params)

    # ------------------------------------------------------------------
    # 用户 / 模拟登录
    # ------------------------------------------------------------------

    def log_on(
        self,
        *,
        user_account: Optional[str] = None,
        user_password: Optional[str] = None,
        **extra,
    ) -> JsonDict:
        """登录 OMS 账户 ``logOn``（模拟登录）。

        成功时 ``data`` 为 URL 编码的快捷登录地址，需 ``urllib.parse.unquote`` 解码后使用。
        账号/密码缺省时用 ``config.user_account`` / ``config.user_password``。
        """
        account = (user_account if user_account is not None else self.config.user_account) or ""
        password = (user_password if user_password is not None else self.config.user_password) or ""
        account = str(account).strip()
        password = str(password).strip()
        if not account or not password:
            raise ValueError(
                "logOn 需要 user_account / user_password："
                "调用时传入，或在 config/secrets.json 的 hy_oms 中填写 user_account / user_password。"
            )
        params: Dict[str, Any] = {
            "user_account": account,
            "user_password": password,
            **extra,
        }
        return self.call("logOn", params)

    def get_sso_token(
        self,
        *,
        company_code: Optional[str] = None,
        **extra,
    ) -> JsonDict:
        """获取登陆 token ``getSsoToken``。

        必填 ``company_code``（客户代码）；缺省用 ``config.company_code``。
        成功时 ``data`` 含 ``userCode`` / ``token``，可拼快捷登录 URL。
        """
        code = company_code if company_code is not None else self.config.company_code
        code = str(code or "").strip()
        if not code:
            raise ValueError(
                "getSsoToken 需要 company_code："
                "调用时传入，或在 config/secrets.json 的 hy_oms 中填写 company_code。"
            )
        params: Dict[str, Any] = {"company_code": code, **extra}
        return self.call("getSsoToken", params)

    def build_quick_login_url(self, *, user_code: str, token: str) -> str:
        """拼装 OMS 快捷登录地址 ``/default/index/quick-login``。"""
        from urllib.parse import urlencode

        user_code = str(user_code or "").strip()
        token = str(token or "").strip()
        if not user_code or not token:
            raise ValueError("build_quick_login_url 需要 user_code 与 token")
        return f"{self.config.quick_login_path}?{urlencode({'userCode': user_code, 'token': token})}"

    def simulate_login(
        self,
        *,
        user_account: Optional[str] = None,
        user_password: Optional[str] = None,
        company_code: Optional[str] = None,
        mode: str = "logOn",
        **extra,
    ) -> JsonDict:
        """模拟登录，返回含解码后快捷登录 URL 的结果。

        ``mode``::
            - ``logOn``（默认）：账号密码登录，``data`` 为 URL 编码地址
            - ``getSsoToken``：按客户代码取 token，再拼快捷登录 URL

        返回 dict 在原始响应基础上增加 ``login_url``（已 urldecode / 拼装）。
        """
        from urllib.parse import unquote

        mode_key = str(mode or "logOn").strip().lower().replace("-", "_")
        if mode_key in {"logon", "log_on", "login"}:
            result = self.log_on(
                user_account=user_account,
                user_password=user_password,
                **extra,
            )
            raw = result.get("data")
            login_url = unquote(str(raw)) if raw is not None else ""
            out = dict(result)
            out["login_url"] = login_url
            return out

        if mode_key in {"getssotoken", "get_sso_token", "sso", "sso_token"}:
            result = self.get_sso_token(company_code=company_code, **extra)
            data = result.get("data") or {}
            # 部分环境 data 再包一层 {ask, data:{userCode,token}}
            if isinstance(data, dict) and "userCode" not in data and isinstance(
                data.get("data"), dict
            ):
                data = data["data"]
            if not isinstance(data, dict):
                raise HyOmsResponseError("getSsoToken 的 data 不是对象", raw=result)
            user_code = data.get("userCode") or data.get("user_code") or ""
            token = data.get("token") or ""
            login_url = self.build_quick_login_url(
                user_code=str(user_code), token=str(token)
            )
            out = dict(result)
            out["login_url"] = login_url
            out["user_code"] = str(user_code)
            out["token"] = str(token)
            return out

        raise ValueError(
            f"simulate_login 不支持的 mode: {mode!r}（可用 logOn / getSsoToken）"
        )
