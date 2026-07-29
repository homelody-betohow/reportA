"""同步易仓 SellerSKU 利润报表到 ``amz_seller_sku_profit_snapshot``。

按月拉取 ``getFinancialSellerSKUReportListNew``：
- 默认当前月：开始 = 本月 1 号 00:00:00，结束 = 当天 00:00:00
- 历史月：开始 = 该月 1 号 00:00:00，结束 = 该月最后一天 23:59:59

用法（项目根目录）::

    python app/eccang/syncSellerSKUReport.py
    python app/eccang/syncSellerSKUReport.py --month 2026-06
    python app/eccang/syncSellerSKUReport.py --month 2026-07 --company-code ERP2009186VG
    python app/eccang/syncSellerSKUReport.py --dry-run --limit-pages 1
    python app/eccang/syncSellerSKUReport.py --seller-sku E02022001#FBFBA
    python app/eccang/syncSellerSKUReport.py --seller-sku E02022001#FBFBA --dry-run
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import secrets
import sys
import time
from calendar import monthrange
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Sequence

_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
assert _spec.loader is not None
_spec.loader.exec_module(_epr_mod)
_PROJECT_ROOT = _epr_mod.bootstrap(__file__)

from api.eccang.exceptions import EccangApiError, EccangConfigError  # noqa: E402
from api.eccang.request.getSellerSKUReport import (  # noqa: E402
    DEFAULT_TIME_TYPE,
    DEFAULT_TIME_ZONE_TYPE,
    DEFAULT_UNIT_CURRENCY,
    get_financial_seller_sku_report_list,
)
from database.db_connection import get_db_manager  # noqa: E402

TABLE = "amz_seller_sku_profit_snapshot"
DEFAULT_COMPANY_CODE = "ERP2009186VG"
DEFAULT_PAGE_SIZE = 500
SNAPSHOT_TYPE = "eccang_api_monthly"
BATCH_SIZE = 200
UNKNOWN_SELLER_SKU_PREFIX = "_unknown"
UNKNOWN_SELLER_SKU_RAND_LEN = 12


def make_unknown_seller_sku() -> str:
    """空 sellerSku 占位：``_unknown`` + 12 位十六进制随机串。"""
    return f"{UNKNOWN_SELLER_SKU_PREFIX}{secrets.token_hex(UNKNOWN_SELLER_SKU_RAND_LEN // 2)}"

# DB 列 → API 字段（按 COMMENT / Excel 列语义对齐）
FIELD_MAP: dict[str, str] = {
    "seller_sku": "sellerSku",
    "asin": "asin",
    "parent_asin": "parentAsin",
    "product_info": "exportProductTitle",
    "warehouse_sku": "exportWarehouseSku",
    "shop_name": "userAccount",
    "marketplace_site": "site",
    "ops_owner": "sellerRealName",
    "brand": "sellerSkuBrand",
    "selling_status": "sellerSkuItemStatusText",
    "category": "sellerSkuPcName1",
    "tag_label": "sellerSkuTagsText",
    "currency": "currencyCode",
    "sales_qty": "quantityTotal",
    "fba_sales_qty": "fbaQuantity",
    "fbm_sales_qty": "fbmQuantity",
    "multi_channel_sales_qty": "multiChannelQuantity",
    "ad_sales_qty": "adSalesQty",
    "sp_ad_sales_qty": "spAdSalesQty",
    "sd_ad_sales_qty": "sdAdSalesQty",
    "refund_qty": "refundQuantityTotal",
    "fba_refund_qty": "fbaRefundQuantity",
    "fbm_refund_qty": "fbmRefundQuantity",
    "refund_rate_text": "refundRateText",
    "return_qty": "returnQuantityTotal",
    "return_qty_sellable": "returnSellableQuantity",
    "return_rate_text": "returnRateText",
    "sales_amount": "productSalesTotal",
    "fba_sales_amount": "fbaProductSales",
    "fbm_sales_amount": "fbmProductSales",
    "ad_sales_amount": "adSales",
    "sp_ad_sales_amount": "spAdSales",
    "sd_ad_sales_amount": "sdAdSales",
    "sb_ad_sales_amount": "sbAdSales",
    "sbv_sales_amount": "sbvAdSales",
    "sales_refund_amount": "refundFbaProductSalesTotal",
    "fba_sales_refund_amount": "fbaRefundFbaProductSales",
    "fbm_sales_refund_amount": "fbmRefundFbmProductSales",
    "buyer_shipping_total": "shippingCreditsTotal",
    "fba_buyer_shipping": "fbaShippingCredits",
    "fbm_buyer_shipping": "fbmShippingCredits",
    "buyer_shipping_refund": "refundShippingCredits",
    "gift_wrap_fee_total": "giftWrapCreditsTotal",
    "gift_wrap_fee": "giftWrapCredits",
    "gift_wrap_fee_refund": "refundGiftWrapCredits",
    "promo_discount_total": "promotionalRebatesTotal",
    "promo_discount": "promotionalRebates",
    "promo_discount_refund": "refundPromotionalRebates",
    "cod_cash_on_delivery": "codFee",
    "fba_inventory_reimbursement_total": "fbaReimbursementTotal",
    "fba_inv_reimb_customer_related": "fbaReimbursementCustomer",
    "fba_inv_reimb_warehouse": "fbaReimbursementWarehouse",
    "fba_inv_reimb_warehouse_other": "fbaReimbursementOther",
    "other_adjustment_income": "otherAdjustmentFee",
    "platform_other_income_total": "paltformOtherIncomeTotal",
    "liquidation_adjustment": "liquidationFeeAdjustment",
    "credit_card_chargeback": "creditCardFee",
    "amazon_shipping_reimbursement": "amazonShipmentReimbursement",
    "selling_commission_total": "sellingFeesTotal",
    "fba_selling_commission": "fbaSellingFees",
    "fbm_selling_commission": "fbmSellingFees",
    "selling_commission_refund": "refundSellingFees",
    "fulfillment_fee_total": "fbaFeesTotal",
    "fba_fulfillment_fee": "fbaFees",
    "fba_fulfillment_fee_refund": "refundFbaFees",
    "fbm_fulfillment_fee": "fbmFees",
    "erp_fba_fee": "fmsFbaFee",
    "multi_channel_fulfillment_fee": "multiFbaFees",
    "other_txn_fee_total": "otherTransactionFeeTotal",
    "other_txn_fee": "otherTransactionFee",
    "other_txn_fee_refund": "refundOtherTransactionFee",
    "fba_inventory_inbound_service_fee_total": "fbaInventoryFeeTotal",
    "fba_monthly_storage_fee": "fbaMonthStorageFee",
    "fba_long_term_storage_fee": "fbaLongtermStorageFee",
    "storage_fee_allocated": "storageFee",
    "long_term_storage_fee_allocated": "longTermStorageFee",
    "removal_fee": "fbaRemovalFee",
    "inbound_placement_fee": "fbaInboundConvenienceFee",
    "fba_intl_freight": "fbaInternationalInboundFreightFee",
    "other_fba_inventory_inbound_fee": "otherFbaInventoryFeeFbaFee",
    "fba_returns_processing_fee": "fbaReturnDisposalFeeTotal",
    "sd_ad_fee": "sdcost",
    "sp_ad_fee": "spcost",
    "sb_ad_fee": "sbCost",
    "sbv_ad_fee": "sbvCost",
    "ad_invoice_allocated": "adInvoiceCost",
    "promotion_service_fee_total": "couponAndLightningDealFeeTotal",
    "coupon": "couponOtherTransactionFee",
    "lightning_deal_fee": "lightningDealFeeOtherTransactionFees",
    "service_fee_total": "serviceFeeTotal",
    "fba_subscription_fee": "serviceFeeSubscriptionOtherFee",
    "shipping_label_fee": "serviceFeeFbaFrepFeeLabelingOtherFee",
    "other_service_fee": "otherServiceFee",
    "amazon_global_logistics_freight_total": "fbaInternationalFreightTotal",
    "amazon_global_logistics_freight": "fbaInternationalFreightShippingCharge",
    "amazon_global_logistics_duties_taxes": "fbaInternationalFreightDutiesAndTaxesCharge",
    "points_purchase": "amazonPointCosts",
    "platform_other_expense_total": "platformOtherExpendTotal",
    "liquidation_fee": "liquidationFeeExpend",
    "platform_other_expense": "platformOtherExpend",
    "other_uncategorized_fee": "otherUnclassifiedExpenses",
    "tax_total": "taxTotal",
    "product_tax": "productSalesTax",
    "shipping_tax": "shippingCreditsTax",
    "gift_wrap_tax": "giftWrapCreditsTax",
    "regulatory_fee": "regulatoryFee",
    "promo_discount_tax": "promotionalRebatesTax",
    "product_tax_refund": "refundProductSalesTax",
    "shipping_tax_refund": "refundShippingCreditsTax",
    "gift_wrap_tax_refund": "refundGiftWrapCreditsTax",
    "promo_discount_tax_refund": "refundPromotionalRebatesTax",
    "amazon_withheld_tax": "marketplaceWithheldTaxAndVat",
    "low_value_tax": "lowValueGoods",
    "hidden_tax_income": "baseTaxIncome",
    "hidden_tax_expense": "baseTaxExpense",
    "mixed_vat": "comminglingVat",
    "tcs_cgst": "tcsCgst",
    "tcs_sgst": "tcsSgst",
    "tcs_igst": "tcsIgst",
    "tds": "tds",
    "product_cost": "productCostTotal",
    "fba_order_purchase_cost": "fbaPurchaseCostFee",
    "fba_order_purchase_shipping": "fbaPurchaseShippingFee",
    "fba_order_purchase_tax": "fbaPurchaseTariffFee",
    "fbm_order_purchase_cost": "fbmPurchaseCostFee",
    "fbm_order_purchase_shipping": "fbmPurchaseShippingFee",
    "fbm_order_purchase_tax": "fbmPurchaseTariffFee",
    "review_order_purchase_cost": "evaluatePurchaseCostFee",
    "review_order_purchase_shipping": "evaluatePurchaseShippingFee",
    "review_order_purchase_tax": "evaluatePurchaseTariffFee",
    "multi_channel_order_purchase_cost": "multiPurchaseCostFee",
    "multi_channel_order_purchase_shipping": "multiPurchaseShippingFee",
    "multi_channel_order_purchase_tax": "multiPurchaseTariffFee",
    "fba_order_first_leg_shipping": "fbaFirstShippingFee",
    "fbm_order_first_leg_shipping": "fbmFirstShippingFee",
    "review_order_first_leg_shipping": "evaluateFirstShippingFee",
    "multi_channel_order_first_leg_shipping": "multiFirstShippingFee",
    "fba_order_first_leg_tax": "fbaFirstShippingTariffFee",
    "fbm_order_first_leg_tax": "fbmFirstShippingTariffFee",
    "review_order_first_leg_tax": "evaluateFirstShippingTariffFee",
    "multi_channel_order_first_leg_tax": "multiFirstShippingTariffFee",
    "packaging_fee": "packageFee",
    "return_restore_purchase_cost_sellable": "fbaReturnProductSellablePurchaseCost",
    "return_restore_purchase_cost_unsellable": "fbaReturnProductCustomerUnsellablePurchaseCost",
    "return_restore_purchase_cost_unsellable_amz_comp": "fbaReturnProductAmazonUnsellablePurchaseCost",
    "return_restore_first_leg_shipping_sellable": "fbaReturnProductSellableFirstShippingFee",
    "return_restore_first_leg_shipping_unsellable": "fbaReturnProductCustomerUnsellableFirstShippingFee",
    "return_restore_first_leg_shipping_unsellable_amz_comp": "fbaReturnProductAmazonUnsellableFirstShippingFee",
    "return_restore_first_leg_tax_sellable": "fbaReturnProductSellableFirstShippingTariffFee",
    "return_restore_first_leg_tax_unsellable": "fbaReturnProductCustomerUnsellableFirstShippingTariffFee",
    "return_restore_first_leg_tax_unsellable_amz_comp": "fbaReturnProductAmazonUnsellableFirstShippingTariffFee",
    "warehouse_other_purchase_cost": "fbaOtherProductPurchaseCostTotal",
    "warehouse_other_purchase_cost_adj_in": "fbaOtherProductInAdjustmentsPurchaseCostTotal",
    "warehouse_other_purchase_cost_adj_out": "fbaOtherProductOutAdjustmentsPurchaseCostTotal",
    "warehouse_other_purchase_cost_removal_sellable": "fbaOtherProductOutRemovalSellablePurchaseCostTotal",
    "warehouse_other_purchase_cost_removal_unsellable": "fbaOtherProductOutRemovalUnsellablePurchaseCostTotal",
    "warehouse_other_purchase_cost_surplus": "fbaOtherProductInInventoryPurchaseCostTotal",
    "warehouse_other_purchase_cost_shortage": "fbaOtherProductOutInventoryPurchaseCostTotal",
    "warehouse_other_first_leg_cost": "fbaOtherProductShippingCostTotal",
    "warehouse_other_first_leg_cost_adj_in": "fbaOtherProductInAdjustmentsShippingCostTotal",
    "warehouse_other_first_leg_cost_adj_out": "fbaOtherProductOutAdjustmentsShippingCostTotal",
    "warehouse_other_first_leg_cost_removal_sellable": "fbaOtherProductOutRemovalSellableShippingCostTotal",
    "warehouse_other_first_leg_cost_removal_unsellable": "fbaOtherProductOutRemovalUnsellableShippingCostTotal",
    "warehouse_other_first_leg_cost_surplus": "fbaOtherProductInInventoryShippingCostTotal",
    "warehouse_other_first_leg_cost_shortage": "fbaOtherProductOutInventoryShippingCostTotal",
    "offline_fee_total": "customFeeTotal",
    "review_fee_total": "amountEvaluateFee",
    "platform_gross_profit": "platformGrossProfit",
    "default_gross_profit": "grossProfitTotal",
    "gross_margin_rate_text": "grossProfitRateText",
    "roi_text": "roiText",
    "first_order_time": "firstOrderTime",
    "listing_time": "openDate",
    "dev_owner": "personDevelopNames",
    "low_price_shop_item_flag": "lowPriceText",
}

TEXT_COLS = {
    "seller_sku",
    "asin",
    "parent_asin",
    "historical_asin",
    "product_info",
    "warehouse_sku",
    "shop_name",
    "marketplace_site",
    "ops_owner",
    "brand",
    "selling_status",
    "category",
    "tag_label",
    "currency",
    "refund_rate_text",
    "return_rate_text",
    "gross_margin_rate_text",
    "roi_text",
    "dev_owner",
    "low_price_shop_item_flag",
    "snapshot_name",
    "snapshot_type",
    "snapshot_id",
}

DATETIME_COLS = {"first_order_time", "listing_time"}
META_COLS = ("snapshot_id", "snapshot_name", "snapshot_date", "snapshot_type")
DATA_COLS = tuple(FIELD_MAP.keys())
ALL_COLS = META_COLS + DATA_COLS


def resolve_month_range(
    month: str | None = None,
    *,
    today: date | None = None,
) -> tuple[date, datetime, datetime]:
    """返回 (snapshot_date, start_dt, end_dt)。

    snapshot_date 取该月最后一天。
    当前月：end = 当天 00:00:00；历史月：end = 月末 23:59:59。
    """
    now = today or date.today()
    if month:
        try:
            year_s, month_s = month.strip().split("-", 1)
            year, mon = int(year_s), int(month_s)
            if not (1 <= mon <= 12):
                raise ValueError
        except ValueError as exc:
            raise ValueError(f"无效 --month：{month!r}，期望 YYYY-MM") from exc
        year_i, mon_i = year, mon
    else:
        year_i, mon_i = now.year, now.month

    last_day = monthrange(year_i, mon_i)[1]
    snap = date(year_i, mon_i, last_day)
    start_dt = datetime(year_i, mon_i, 1, 0, 0, 0)
    is_current = year_i == now.year and mon_i == now.month
    if is_current:
        end_dt = datetime(now.year, now.month, now.day, 0, 0, 0)
        if end_dt <= start_dt:
            # 月初当天：至少给 1 秒窗口，避免 start==end
            end_dt = start_dt.replace(second=1)
    else:
        end_dt = datetime(year_i, mon_i, last_day, 23, 59, 59)
    return snap, start_dt, end_dt


def build_snapshot_id(
    *,
    company_code: str,
    snapshot_date: date,
    start_dt: datetime,
    end_dt: datetime,
) -> str:
    raw = (
        f"eccang|getFinancialSellerSKUReportListNew|{company_code}|"
        f"{snapshot_date:%Y-%m}|{start_dt:%Y-%m-%d %H:%M:%S}|{end_dt:%Y-%m-%d %H:%M:%S}"
    )
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def _as_text(value: Any, *, max_len: int | None = None) -> str | None:
    if value is None:
        return None
    if isinstance(value, bool):
        s = "1" if value else "0"
    else:
        s = str(value).strip()
    if not s:
        return None
    if max_len is not None and len(s) > max_len:
        return s[:max_len]
    return s


def _as_decimal(value: Any) -> Decimal | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, Decimal):
        return value
    try:
        return Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError):
        return None


def _as_datetime(value: Any) -> datetime | None:
    text = _as_text(value)
    if not text:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _pick(record: dict[str, Any], *keys: str) -> Any:
    for key in keys:
        if key not in record:
            continue
        val = record.get(key)
        if val is None:
            continue
        if isinstance(val, str) and not val.strip():
            continue
        return val
    return None


def extract_records(response: dict[str, Any]) -> tuple[list[dict[str, Any]], int | None, int | None]:
    data = response.get("data")
    if not isinstance(data, dict):
        return [], None, None
    records = data.get("records")
    total = data.get("total")
    pages = data.get("pages")
    if records is None and isinstance(data.get("data"), dict):
        inner = data["data"]
        records = inner.get("records")
        total = inner.get("total")
        pages = inner.get("pages")
    if not isinstance(records, list):
        records = []
    return records, (int(total) if total is not None else None), (
        int(pages) if pages is not None else None
    )


def map_record(
    record: dict[str, Any],
    *,
    snapshot_id: str,
    snapshot_name: str,
    snapshot_date: date,
    snapshot_type: str,
) -> dict[str, Any]:
    seller_sku = _as_text(_pick(record, "sellerSku"), max_len=128)
    filled_unknown = False
    if not seller_sku:
        # seller_sku NOT NULL，且同店铺多条空值会撞唯一键 → 自动占位
        seller_sku = make_unknown_seller_sku()
        filled_unknown = True

    shop = _as_text(
        _pick(record, "userAccount", "platformUserName"),
        max_len=128,
    ) or ""

    warehouse = _as_text(
        _pick(record, "exportWarehouseSku", "warehouseSku"),
        max_len=128,
    )
    product_info = _as_text(
        _pick(record, "exportProductTitle", "sellerSkuTitle"),
        max_len=512,
    )

    row: dict[str, Any] = {
        "snapshot_id": snapshot_id,
        "snapshot_name": snapshot_name,
        "snapshot_date": snapshot_date,
        "snapshot_type": snapshot_type,
        "seller_sku": seller_sku,
        "shop_name": shop,
        "warehouse_sku": warehouse,
        "product_info": product_info,
        "historical_asin": None,
        "_filled_unknown_seller_sku": filled_unknown,
    }

    for db_col, api_key in FIELD_MAP.items():
        if db_col in row:
            continue
        raw = record.get(api_key)
        if db_col in DATETIME_COLS:
            row[db_col] = _as_datetime(raw)
        elif db_col in TEXT_COLS:
            maxlen = 512 if db_col == "product_info" else 128
            if db_col in {"asin", "parent_asin", "historical_asin", "marketplace_site"}:
                maxlen = 32
            if db_col in {"refund_rate_text", "return_rate_text", "currency"}:
                maxlen = 32
            if db_col in {"gross_margin_rate_text", "roi_text"}:
                maxlen = 32
            row[db_col] = _as_text(raw, max_len=maxlen)
        else:
            row[db_col] = _as_decimal(raw)

    # 广告费用汇总
    ad_parts = [
        row.get("sd_ad_fee"),
        row.get("sp_ad_fee"),
        row.get("sb_ad_fee"),
        row.get("sbv_ad_fee"),
    ]
    if any(v is not None for v in ad_parts):
        row["ad_fee_total"] = sum((v or Decimal("0") for v in ad_parts), Decimal("0"))
    else:
        row["ad_fee_total"] = None

    return row


def fetch_all_records(
    *,
    company_code: str,
    start_time: str,
    end_time: str,
    page_size: int = DEFAULT_PAGE_SIZE,
    unit_currency: str = DEFAULT_UNIT_CURRENCY,
    time_zone_type: int = DEFAULT_TIME_ZONE_TYPE,
    time_type: int = DEFAULT_TIME_TYPE,
    seller_sku: str | None = None,
    limit_pages: int | None = None,
    sleep_sec: float = 0.2,
) -> list[dict[str, Any]]:
    page = 1
    all_records: list[dict[str, Any]] = []
    total: int | None = None
    pages: int | None = None
    sku = (seller_sku or "").strip() or None
    seller_sku_list = [sku] if sku else None

    while True:
        if limit_pages is not None and page > limit_pages:
            break
        resp = get_financial_seller_sku_report_list(
            company_code=company_code,
            start_time=start_time,
            end_time=end_time,
            page=page,
            page_size=page_size,
            unit_currency=unit_currency,
            time_zone_type=time_zone_type,
            time_type=time_type,
            seller_sku_list=seller_sku_list,
        )
        if sku:
            printable = {k: v for k, v in resp.items() if k != "biz_content"}
            print(json.dumps(printable, ensure_ascii=False, indent=2))
        batch, total, pages = extract_records(resp)
        all_records.extend(batch)
        print(
            f"[API] page={page}/{pages or '?'} got={len(batch)} "
            f"accum={len(all_records)} total={total}",
            file=sys.stderr,
        )
        if not batch:
            break
        if pages is not None and page >= pages:
            break
        if total is not None and len(all_records) >= total:
            break
        if len(batch) < page_size and pages is None:
            break
        page += 1
        if sleep_sec > 0:
            time.sleep(sleep_sec)
    return all_records


def _row_values(row: dict[str, Any]) -> tuple[Any, ...]:
    cols = list(ALL_COLS) + ["ad_fee_total"]
    # ad_fee_total 已在 FIELD_MAP 外，单独补列
    return tuple(row.get(c) for c in cols)


UPSERT_COLS = list(ALL_COLS) + ["ad_fee_total"]

UPSERT_SQL = f"""
INSERT INTO `{TABLE}` (
    {", ".join(f"`{c}`" for c in UPSERT_COLS)}
) VALUES (
    {", ".join(["%s"] * len(UPSERT_COLS))}
)
ON DUPLICATE KEY UPDATE
    {", ".join(f"`{c}` = VALUES(`{c}`)" for c in UPSERT_COLS if c not in ("snapshot_id", "snapshot_date", "seller_sku", "shop_name"))}
"""


def upsert_rows(rows: Sequence[dict[str, Any]], *, dry_run: bool = False) -> int:
    if not rows:
        return 0
    if dry_run:
        return len(rows)

    db = get_db_manager()
    conn = db.get_connection()
    written = 0
    try:
        with conn.cursor() as cur:
            for i in range(0, len(rows), BATCH_SIZE):
                chunk = rows[i : i + BATCH_SIZE]
                params = [_row_values(r) for r in chunk]
                cur.executemany(UPSERT_SQL, params)
                written += len(chunk)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return written


def delete_snapshot(
    *,
    snapshot_id: str,
    snapshot_date: date,
    seller_sku: str | None = None,
    dry_run: bool = False,
) -> int:
    """删除快照行；传 ``seller_sku`` 时仅删该 SKU 相关行。"""
    if dry_run:
        return 0
    db = get_db_manager()
    conn = db.get_connection()
    try:
        with conn.cursor() as cur:
            if seller_sku:
                cur.execute(
                    f"DELETE FROM `{TABLE}` "
                    f"WHERE `snapshot_id` = %s AND `snapshot_date` = %s "
                    f"AND `seller_sku` = %s",
                    (snapshot_id, snapshot_date, seller_sku),
                )
            else:
                cur.execute(
                    f"DELETE FROM `{TABLE}` "
                    f"WHERE `snapshot_id` = %s AND `snapshot_date` = %s",
                    (snapshot_id, snapshot_date),
                )
            deleted = int(cur.rowcount or 0)
        conn.commit()
        return deleted
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def sync_seller_sku_report(
    *,
    company_code: str = DEFAULT_COMPANY_CODE,
    month: str | None = None,
    page_size: int = DEFAULT_PAGE_SIZE,
    unit_currency: str = DEFAULT_UNIT_CURRENCY,
    time_zone_type: int = DEFAULT_TIME_ZONE_TYPE,
    time_type: int = DEFAULT_TIME_TYPE,
    seller_sku: str | None = None,
    limit_pages: int | None = None,
    replace: bool = True,
    dry_run: bool = False,
    sleep_sec: float = 0.2,
) -> dict[str, Any]:
    snap_date, start_dt, end_dt = resolve_month_range(month)
    start_s = start_dt.strftime("%Y-%m-%d %H:%M:%S")
    end_s = end_dt.strftime("%Y-%m-%d %H:%M:%S")
    snapshot_id = build_snapshot_id(
        company_code=company_code,
        snapshot_date=snap_date,
        start_dt=start_dt,
        end_dt=end_dt,
    )
    snapshot_name = (
        f"SellerSKU利润报表-{snap_date:%Y-%m}"
    )
    sku = (seller_sku or "").strip() or None

    print(
        f"[SYNC] company={company_code} month={snap_date:%Y-%m} "
        f"currency={unit_currency} timeZoneType={time_zone_type} "
        f"timeType={time_type} seller_sku={sku or '-'} "
        f"range=[{start_s} ~ {end_s}] snapshot_id={snapshot_id[:12]}…",
        file=sys.stderr,
    )

    raw_records = fetch_all_records(
        company_code=company_code,
        start_time=start_s,
        end_time=end_s,
        page_size=page_size,
        unit_currency=unit_currency,
        time_zone_type=time_zone_type,
        time_type=time_type,
        seller_sku=sku,
        limit_pages=limit_pages,
        sleep_sec=sleep_sec,
    )

    rows: list[dict[str, Any]] = []
    skipped = 0
    filled_unknown = 0
    for rec in raw_records:
        if not isinstance(rec, dict):
            skipped += 1
            continue
        mapped = map_record(
            rec,
            snapshot_id=snapshot_id,
            snapshot_name=snapshot_name,
            snapshot_date=snap_date,
            snapshot_type=SNAPSHOT_TYPE,
        )
        if mapped.pop("_filled_unknown_seller_sku", False):
            filled_unknown += 1
        rows.append(mapped)

    deleted = 0
    if not dry_run:
        if sku:
            # 单 SKU：仅刷新该 seller_sku 相关行，不动其他 SKU
            deleted = delete_snapshot(
                snapshot_id=snapshot_id,
                snapshot_date=snap_date,
                seller_sku=sku,
                dry_run=False,
            )
        elif replace:
            deleted = delete_snapshot(
                snapshot_id=snapshot_id,
                snapshot_date=snap_date,
                dry_run=False,
            )

    written = upsert_rows(rows, dry_run=dry_run)
    return {
        "company_code": company_code,
        "month": f"{snap_date:%Y-%m}",
        "start_time": start_s,
        "end_time": end_s,
        "snapshot_id": snapshot_id,
        "snapshot_date": snap_date.isoformat(),
        "snapshot_name": snapshot_name,
        "fetched": len(raw_records),
        "mapped": len(rows),
        "skipped": skipped,
        "filled_unknown_seller_sku": filled_unknown,
        "deleted": deleted,
        "written": written,
        "seller_sku": sku,
        "update_mode": "seller_sku" if sku else ("replace" if replace else "upsert"),
        "dry_run": dry_run,
    }


def main(argv: list[str] | None = None) -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

    parser = argparse.ArgumentParser(
        description="同步易仓 SellerSKU 利润报表到 amz_seller_sku_profit_snapshot",
    )
    parser.add_argument(
        "--month",
        help="目标月份 YYYY-MM；默认当前月（结束时间为当天 0 点）",
    )
    parser.add_argument(
        "--company-code",
        default=DEFAULT_COMPANY_CODE,
        dest="company_code",
        help=f"易仓 companyCode，默认 {DEFAULT_COMPANY_CODE}",
    )
    parser.add_argument(
        "--page-size",
        type=int,
        default=DEFAULT_PAGE_SIZE,
        dest="page_size",
        help=f"分页大小，默认 {DEFAULT_PAGE_SIZE}，最大 500",
    )
    parser.add_argument(
        "--unit-currency",
        default=DEFAULT_UNIT_CURRENCY,
        dest="unit_currency",
        help=f"币种 unitCurrency，默认 {DEFAULT_UNIT_CURRENCY}",
    )
    parser.add_argument(
        "--time-zone-type",
        type=int,
        default=DEFAULT_TIME_ZONE_TYPE,
        dest="time_zone_type",
        help=f"时区类型 timeZoneType：1=北京时间，2=站点时间，默认 {DEFAULT_TIME_ZONE_TYPE}",
    )
    parser.add_argument(
        "--time-type",
        type=int,
        default=DEFAULT_TIME_TYPE,
        dest="time_type",
        help=f"时间类型 timeType：1=下单时间，2=结算时间，默认 {DEFAULT_TIME_TYPE}",
    )
    parser.add_argument(
        "--seller-sku",
        dest="seller_sku",
        help="单个 SellerSku：打印 API 返回，并仅更新该 SKU 相关数据行（不清空整月）",
    )
    parser.add_argument(
        "--limit-pages",
        type=int,
        default=None,
        dest="limit_pages",
        help="最多拉取页数（调试用）",
    )
    parser.add_argument(
        "--no-replace",
        action="store_true",
        help="不清空同 snapshot 旧数据，仅 UPSERT",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只拉取并映射，不写库",
    )
    parser.add_argument(
        "--sleep",
        type=float,
        default=0.2,
        help="分页请求间隔秒数，默认 0.2",
    )
    args = parser.parse_args(argv)

    try:
        result = sync_seller_sku_report(
            company_code=args.company_code,
            month=args.month,
            page_size=min(max(1, args.page_size), 500),
            unit_currency=args.unit_currency,
            time_zone_type=args.time_zone_type,
            time_type=args.time_type,
            seller_sku=args.seller_sku,
            limit_pages=args.limit_pages,
            replace=not args.no_replace,
            dry_run=args.dry_run,
            sleep_sec=max(0.0, args.sleep),
        )
    except ValueError as exc:
        print(f"[FAIL] 参数错误：{exc}", file=sys.stderr)
        return 2
    except EccangConfigError as exc:
        print(f"[FAIL] 配置错误：{exc}", file=sys.stderr)
        return 2
    except EccangApiError as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 1
    except FileNotFoundError as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 2

    print(json.dumps(result, ensure_ascii=False, indent=2))
    mode = "dry-run" if result["dry_run"] else result.get("update_mode", "saved")
    print(
        f"[OK] {mode} month={result['month']} "
        f"seller_sku={result.get('seller_sku') or '-'} "
        f"fetched={result['fetched']} mapped={result['mapped']} "
        f"unknown_sku={result['filled_unknown_seller_sku']} "
        f"written={result['written']} deleted={result['deleted']}",
        file=sys.stderr,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
