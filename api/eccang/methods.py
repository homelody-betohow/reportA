from __future__ import annotations

from typing import Any

from .client import EccangClient


class EccangMethods:
    """常用易仓 ERP 接口 method 常量（参考官方文档）。

    文档地址：https://open.eccang.com/#/documentCenter
    
    根据实际业务需求，可以添加更多接口方法。
    """

    # 基础数据类接口
    GET_SHIP_ADDRESS_BOOKS = "getShipAddressBooks"  # 获取发货地址簿
    GET_WAREHOUSE_LIST = "getWarehouseList"  # 获取仓库列表（简要）
    GET_WAREHOUSE = "getWarehouse"  # WMS-获取仓库信息（docId=393）
    GET_CARRIER_LIST = "getCarrierList"  # 获取物流商列表
    
    # 订单类接口
    GET_ORDER_LIST = "getOrderList"  # 获取订单列表（ERP）
    GET_ORDER_DETAIL = "getOrderDetail"  # 获取订单详情
    CREATE_ORDER = "createOrder"  # 创建订单
    CANCEL_ORDER = "cancelOrder"  # 取消订单
    # WMS-查询仓储订单信息（docId=735，version=V1.0.0）
    # https://open.eccang.com/#/documentCenter?docId=735&catId=0-181-181,0-177
    GET_ORDERS = "getOrders"
    
    # 产品类接口
    GET_PRODUCT_LIST = "getProductList"  # 获取产品列表
    # WMS-获取产品列表（docId=737，version=V1.0.0）
    # https://open.eccang.com/#/documentCenter?docId=737&catId=0-187-187,0-177
    GET_WMS_PRODUCT_LIST = "getWmsProductList"
    # 获取产品报关属性字典（docId=111799，version=V1.0.0）
    # https://open.eccang.com/#/documentCenter?docId=111799&catId=0-187-187,0-177
    GET_PRODUCT_CUSTOMS_ATTRIBUTE = "getProductCustomsAttribute"
    GET_PRODUCT_DETAIL = "getProductDetail"  # 获取产品详情
    CREATE_PRODUCT = "createProduct"  # 创建产品
    UPDATE_PRODUCT = "updateProduct"  # 更新产品
    
    # 库存类接口
    GET_INVENTORY_LIST = "getInventoryList"  # 获取库存列表
    GET_INVENTORY_DETAIL = "getInventoryDetail"  # 获取库存详情
    GET_PRODUCT_INVENTORY_NEW = "getProductInventoryNew"  # 产品库存（新，docId=112171）
    
    # 入库单类接口
    GET_INBOUND_LIST = "getInboundList"  # 获取入库单列表
    GET_INBOUND_DETAIL = "getInboundDetail"  # 获取入库单详情
    CREATE_INBOUND = "createInbound"  # 创建入库单
    
    # 出库单类接口
    GET_OUTBOUND_LIST = "getOutboundList"  # 获取出库单列表
    GET_OUTBOUND_DETAIL = "getOutboundDetail"  # 获取出库单详情
    CREATE_OUTBOUND = "createOutbound"  # 创建出库单
    
    # 费用类接口
    GET_BILLING_LIST = "getBillingList"  # 获取账单列表
    GET_BILLING_DETAIL = "getBillingDetail"  # 获取账单详情
    
    # 仓租类接口（具体 method 名称需参考文档）
    GET_WAREHOUSE_RENT = "getWarehouseRent"  # 获取仓租明细

    # Amazon transaction 交易明细（新）
    # 文档：https://open.eccang.com/#/documentCenter?docId=112265&catId=0-508-508,0-177
    GET_TRANSACTION_REPORT_DETAIL_LIST = "getTransactionReportDetailList"

    # 财务 SellerSKU 维度利润列表（新）
    # 文档：https://open.eccang.com/#/documentCenter?docId=112282&catId=0-508-508,0-177
    # 请求方法：getFinancialSellerSKUReportListNew；biz_content 必填 companyCode/startTime/endTime
    GET_FINANCIAL_SELLER_SKU_REPORT_LIST_NEW = "getFinancialSellerSKUReportListNew"

    # 平台账号（店铺）列表
    # 文档：https://open.eccang.com/#/documentCenter?docId=469&catId=0-226-226,0-177
    # 请求方法：getUserAccountList；biz_content.platform 必填
    GET_USER_ACCOUNT_LIST = "getUserAccountList"


class EccangService(EccangClient):
    """在通用客户端上封装常用业务调用。"""

    def get_ship_address_books(
        self,
        *,
        page: int = 1,
        page_size: int = 50,
        extra: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """获取发货地址簿列表。"""
        body: dict[str, Any] = {
            "page": page,
            "page_size": page_size,
        }
        if extra:
            body.update(extra)
        return self.call(EccangMethods.GET_SHIP_ADDRESS_BOOKS, body)

    def get_warehouse_list(
        self,
        *,
        page: int = 1,
        page_size: int = 50,
        extra: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """获取仓库列表。"""
        body: dict[str, Any] = {
            "page": page,
            "page_size": page_size,
        }
        if extra:
            body.update(extra)
        return self.call(EccangMethods.GET_WAREHOUSE_LIST, body)

    def get_order_list(
        self,
        *,
        page: int = 1,
        page_size: int = 50,
        start_time: str | None = None,
        end_time: str | None = None,
        extra: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """获取订单列表。

        Args:
            page: 页码
            page_size: 每页数量
            start_time: 更新时间起（格式：YYYY-MM-DD HH:MM:SS）
            end_time: 更新时间止（格式：YYYY-MM-DD HH:MM:SS）
            extra: 其他额外参数
        """
        body: dict[str, Any] = {
            "page": page,
            "page_size": page_size,
            "pageSize": page_size,
        }
        if start_time:
            body["update_date_start"] = start_time
            body["start_time"] = start_time
        if end_time:
            body["update_date_end"] = end_time
            body["end_time"] = end_time
        if extra:
            body.update(extra)
        return self.call(EccangMethods.GET_ORDER_LIST, body, version="V1.0.0")

    def get_orders(
        self,
        *,
        page: int = 1,
        page_size: int = 10,
        platform_arr: list[str] | None = None,
        seller_id_arr: list[str] | None = None,
        warehouse_id_arr: list[int] | None = None,
        category: list[int] | None = None,
        product_barcode_arr: list[str] | None = None,
        sm_code_arr: list[str] | None = None,
        country_code_in: list[str] | None = None,
        code: list[str] | None = None,
        order_status: str | int | None = None,
        addressee: str | None = None,
        buyer_id: list[str] | None = None,
        buyer_name: str | None = None,
        buyer_mail: str | None = None,
        buyer_responsible_id: list[int] | None = None,
        develop_responsible_id: list[int] | None = None,
        seller_responsible_id: list[int] | None = None,
        pay_date_for: str | None = None,
        pay_date_to: str | None = None,
        add_date_for: str | None = None,
        add_date_to: str | None = None,
        ship_date_for: str | None = None,
        ship_date_to: str | None = None,
        print_date_for: str | None = None,
        print_date_to: str | None = None,
        pack_date_for: str | None = None,
        pack_date_to: str | None = None,
        update_date_for: str | None = None,
        update_date_to: str | None = None,
        ec_update_time_for: str | None = None,
        ec_update_time_to: str | None = None,
        order_year: int | None = None,
        order_by: list[str] | None = None,
        extra: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """WMS-查询仓储订单信息（docId=735，interface_method=getOrders）。

        Args:
            page: 当前页，默认 1
            page_size: 每页条数，默认 10
            platform_arr: 平台代码列表，如 ["ebay","amazon"]
            seller_id_arr: 账号列表
            warehouse_id_arr: 仓库 Id 列表
            category: 品类 Id 列表
            product_barcode_arr: SKU 列表
            sm_code_arr: 运输方式代码列表
            country_code_in: 国家二字码列表
            code: 仓库单号 / 参考号 / 跟踪号列表
            order_status: 订单状态（0删除/1草稿/2确认/3缺货/4已提交/5已打印/7已打包/8已出库）
            addressee / buyer_name / buyer_mail: 模糊匹配字段
            *_date_for / *_date_to: 时间区间（YYYY-MM-DD HH:MM:SS）
            order_year: 历史订单年份（需配合拆表）
            order_by: 排序，如 ["order_id desc"]
            extra: 其他 biz_content 参数
        """
        size = min(max(1, page_size), 1000)
        body: dict[str, Any] = {
            "page": max(1, page),
            "page_size": size,
        }
        optional: dict[str, Any] = {
            "platform_arr": platform_arr,
            "seller_id_arr": seller_id_arr,
            "warehouse_id_arr": warehouse_id_arr,
            "category": category,
            "product_barcode_arr": product_barcode_arr,
            "sm_code_arr": sm_code_arr,
            "country_code_in": country_code_in,
            "code": code,
            "order_status": (
                str(order_status).strip()
                if order_status is not None and str(order_status).strip() != ""
                else None
            ),
            "addressee": addressee,
            "buyer_id": buyer_id,
            "buyer_name": buyer_name,
            "buyer_mail": buyer_mail,
            "buyer_responsible_id": buyer_responsible_id,
            "develop_responsible_id": develop_responsible_id,
            "seller_responsible_id": seller_responsible_id,
            "pay_date_for": pay_date_for,
            "pay_date_to": pay_date_to,
            "add_date_for": add_date_for,
            "add_date_to": add_date_to,
            "ship_date_for": ship_date_for,
            "ship_date_to": ship_date_to,
            "print_date_for": print_date_for,
            "print_date_to": print_date_to,
            "pack_date_for": pack_date_for,
            "pack_date_to": pack_date_to,
            "update_date_for": update_date_for,
            "update_date_to": update_date_to,
            "ec_update_time_for": ec_update_time_for,
            "ec_update_time_to": ec_update_time_to,
            "order_year": order_year,
            "order_by": order_by,
        }
        for key, value in optional.items():
            if value is not None and value != "":
                body[key] = value
        if extra:
            body.update(extra)
            body["page"] = max(1, page)
            body["page_size"] = size
        return self.call(EccangMethods.GET_ORDERS, body, version="V1.0.0")

    def get_order_detail(self, order_no: str) -> dict[str, Any]:
        """获取订单详情。

        Args:
            order_no: 订单号
        """
        return self.call(EccangMethods.GET_ORDER_DETAIL, {"order_no": order_no})

    def get_product_list(
        self,
        *,
        page: int = 1,
        page_size: int = 50,
        extra: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """获取产品列表。"""
        body: dict[str, Any] = {
            "page": page,
            "page_size": page_size,
        }
        if extra:
            body.update(extra)
        return self.call(EccangMethods.GET_PRODUCT_LIST, body)

    def get_wms_product_list(
        self,
        *,
        page: int = 1,
        page_size: int = 20,
        product_sku: str | None = None,
        product_sku_like: str | None = None,
        product_spu: str | None = None,
        product_title_like: str | None = None,
        warehouse_barcode: str | None = None,
        product_update_time_from: str | None = None,
        product_update_time_to: str | None = None,
        extra: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """WMS-获取产品列表（docId=737，interface_method=getWmsProductList）。

        Args:
            page: 当前页
            page_size: 每页条数（最大 1000）
            product_sku: 产品 SKU（精确）
            product_sku_like: 产品 SKU 模糊查询
            product_spu: 产品款式代码
            product_title_like: 产品名称模糊查询
            warehouse_barcode: 仓库条码
            product_update_time_from: 产品更新时间起
            product_update_time_to: 产品更新时间止
            extra: 其他 biz_content 参数
        """
        # 该接口 biz_content 使用 snake_case；camelCase 会被忽略
        size = min(max(1, page_size), 1000)
        body: dict[str, Any] = {
            "page": page,
            "page_size": size,
        }
        if product_sku:
            body["product_sku"] = product_sku
            body["get_property"] = 1  # 精确查 SKU 时一并返回自定义属性
        if product_sku_like:
            body["product_sku_like"] = product_sku_like
        if product_spu:
            body["product_spu"] = product_spu
        if product_title_like:
            body["product_title_like"] = product_title_like
        if warehouse_barcode:
            body["warehouse_barcode"] = warehouse_barcode
        if product_update_time_from:
            body["product_update_time_from"] = product_update_time_from
        if product_update_time_to:
            body["product_update_time_to"] = product_update_time_to
        if extra:
            body.update(extra)
        return self.call(
            EccangMethods.GET_WMS_PRODUCT_LIST,
            body,
            version="V1.0.0",
        )

    def get_product_customs_attribute(
        self,
        *,
        extra: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """获取产品报关属性字典（docId=111799，interface_method=getProductCustomsAttribute）。

        返回报关属性枚举（如电池类型、磁性物质、木质类等），无需必填业务参数。
        """
        body: dict[str, Any] = {}
        if extra:
            body.update(extra)
        return self.call(
            EccangMethods.GET_PRODUCT_CUSTOMS_ATTRIBUTE,
            body,
            version="V1.0.0",
        )

    def get_inventory_list(
        self,
        *,
        warehouse_code: str | None = None,
        page: int = 1,
        page_size: int = 50,
        extra: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """获取库存列表。

        Args:
            warehouse_code: 仓库编码
            page: 页码
            page_size: 每页数量
            extra: 其他额外参数
        """
        body: dict[str, Any] = {
            "page": page,
            "page_size": page_size,
        }
        if warehouse_code:
            body["warehouse_code"] = warehouse_code
        if extra:
            body.update(extra)
        return self.call(EccangMethods.GET_INVENTORY_LIST, body)

    def get_billing_list(
        self,
        *,
        start_date: str | None = None,
        end_date: str | None = None,
        page: int = 1,
        page_size: int = 50,
        extra: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """获取账单列表。

        Args:
            start_date: 开始日期（格式：YYYY-MM-DD）
            end_date: 结束日期（格式：YYYY-MM-DD）
            page: 页码
            page_size: 每页数量
            extra: 其他额外参数
        """
        body: dict[str, Any] = {
            "page": page,
            "page_size": page_size,
        }
        if start_date:
            body["start_date"] = start_date
        if end_date:
            body["end_date"] = end_date
        if extra:
            body.update(extra)
        return self.call(EccangMethods.GET_BILLING_LIST, body)

    def get_transaction_report_detail_list(
        self,
        *,
        page: int = 1,
        page_size: int = 100,
        date_from: str | None = None,
        date_to: str | None = None,
        extra: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """获取 transaction 交易明细（新）。

        Args:
            page: 页码
            page_size: 每页条数（建议 50~200，最大以文档为准）
            date_from: 开始时间（YYYY-MM-DD 或 YYYY-MM-DD HH:MM:SS）
            date_to: 结束时间
            extra: 其他 biz_content 参数（如 source_type、user_account 等）
        """
        body: dict[str, Any] = {
            "page": page,
            "page_size": page_size,
        }
        if date_from:
            body["start_time"] = date_from
        if date_to:
            body["end_time"] = date_to
        if extra:
            body.update(extra)
        return self.call(EccangMethods.GET_TRANSACTION_REPORT_DETAIL_LIST, body)

    def get_financial_seller_sku_report_list_new(
        self,
        *,
        company_code: str,
        start_time: str,
        end_time: str,
        page: int = 1,
        page_size: int = 50,
        unit_currency: str | None = None,
        site_list: list[str] | None = None,
        user_account_list: list[str] | None = None,
        user_account: str | None = None,
        time_zone_type: int | None = None,
        time_type: int | None = None,
        seller_sku_item_status_list: list[int] | None = None,
        cost_type: int | None = None,
        profit_formula_type: int | None = None,
        search_type: int | None = None,
        keyword: str | None = None,
        seller_sku_list: list[str] | None = None,
        asin_list: list[str] | None = None,
        parent_asin_list: list[str] | None = None,
        transaction_status: str | None = None,
        charge_type: str | None = None,
        account_skus: list[dict[str, str]] | None = None,
        system_code: str | None = None,
        extra: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """获取财务 SellerSKU 维度利润列表（新）。

        文档：docId=112282，interface_method=getFinancialSellerSKUReportListNew。
        起止时间间隔不能超过 31 天；pageSize 最大 500。
        报表存分析库，每日约 9 点前 / 14 点前各更新一次，请控制请求频率。

        Args:
            company_code: 公司代码（必填）
            start_time: 报表开始时间（yyyy-MM-dd HH:mm:ss）
            end_time: 报表结束时间（yyyy-MM-dd HH:mm:ss）
            page: 当前页，默认 1
            page_size: 每页条数，默认 50，最大 500
            unit_currency: 币种 ORIGINAL/RMB/USD/EUR/JPY/GBP/CAD/MXN
            site_list: 站点列表
            user_account_list: 平台账号列表（别名 userAccounts 亦可经 extra 传入）
            user_account: 单个平台账号
            time_zone_type: 时区类型（北京时间 1；站点时间 2，接口默认多为 2）
            time_type: 时间类型（下单时间 1；结算时间 2；接口实测亦接受 3/4）
            seller_sku_item_status_list: 销售状态（在售 1；停售 2；下架 3；已删除 4）
            cost_type: 成本来源（商品成本配置 1；FBA进销存 2；ERP先进先出 4；月末加权 5）
            profit_formula_type: 利润公式（自定义 1；系统默认 2）
            search_type: 查询类型（SellerSku 1；子Asin 2；父Asin 3；品牌 6；品类 7；产品名称 10）
            keyword: 与 search_type 配合的搜索值
            seller_sku_list / asin_list / parent_asin_list: 多值精确匹配（各最大 100）
            transaction_status: 交易状态（已发放 / 已推迟）
            charge_type: 汇率方式（ord / settle）
            account_skus: [{"userAccount":"...","sellerSku":"..."}, ...]
            system_code: 系统编码（如 AMAZON_OPERATE）
            extra: 其他 biz_content 参数
        """
        size = min(max(1, page_size), 500)
        body: dict[str, Any] = {
            "companyCode": company_code,
            "startTime": start_time,
            "endTime": end_time,
            "page": page,
            "pageSize": size,
        }
        if unit_currency:
            body["unitCurrency"] = unit_currency
        if site_list:
            body["siteList"] = site_list
        if user_account_list:
            body["userAccountList"] = user_account_list
        if user_account:
            body["userAccount"] = user_account
        if time_zone_type is not None:
            body["timeZoneType"] = time_zone_type
        if time_type is not None:
            body["timeType"] = time_type
        if seller_sku_item_status_list:
            body["sellerSkuItemStatusList"] = seller_sku_item_status_list
        if cost_type is not None:
            body["costType"] = cost_type
        if profit_formula_type is not None:
            body["profitFormulaType"] = profit_formula_type
        if search_type is not None:
            body["searchType"] = search_type
        if keyword:
            body["keyword"] = keyword
        if seller_sku_list:
            body["sellerSkuList"] = seller_sku_list
        if asin_list:
            body["asinList"] = asin_list
        if parent_asin_list:
            body["parentAsinList"] = parent_asin_list
        if transaction_status:
            body["transactionStatus"] = transaction_status
        if charge_type:
            body["chargeType"] = charge_type
        if account_skus:
            body["accountSkus"] = account_skus
        if system_code:
            body["systemCode"] = system_code
        if extra:
            body.update(extra)
            # 必填字段不被 extra 覆盖丢失
            body["companyCode"] = company_code
            body["startTime"] = start_time
            body["endTime"] = end_time
        return self.call(
            EccangMethods.GET_FINANCIAL_SELLER_SKU_REPORT_LIST_NEW,
            body,
            version="1.0.0",
        )

    def get_user_account_list(
        self,
        *,
        platform: str,
        page: int = 1,
        page_size: int = 200,
        extra: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """获取平台账号列表（店铺/账号）。

        Args:
            platform: 平台代码（必填，如 amazon）
            page: 页码（部分环境会忽略分页，一次返回全量）
            page_size: 每页条数
            extra: 其他 biz_content 参数
        """
        body: dict[str, Any] = {
            "platform": platform,
            "page": page,
            "page_size": page_size,
            "pageSize": page_size,
        }
        if extra:
            body.update(extra)
            body["platform"] = platform
        return self.call(EccangMethods.GET_USER_ACCOUNT_LIST, body, version="V1.0.0")

    def get_warehouse_rent(
        self,
        *,
        warehouse_code: str | None = None,
        start_date: str | None = None,
        end_date: str | None = None,
        page: int = 1,
        page_size: int = 100,
        extra: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """获取仓租明细。

        Args:
            warehouse_code: 仓库编码
            start_date: 开始日期（格式：YYYY-MM-DD）
            end_date: 结束日期（格式：YYYY-MM-DD）
            page: 页码
            page_size: 每页数量
            extra: 其他额外参数
        """
        body: dict[str, Any] = {
            "page": page,
            "page_size": page_size,
        }
        if warehouse_code:
            body["warehouse_code"] = warehouse_code
        if start_date:
            body["start_date"] = start_date
        if end_date:
            body["end_date"] = end_date
        if extra:
            body.update(extra)
        return self.call(EccangMethods.GET_WAREHOUSE_RENT, body)
