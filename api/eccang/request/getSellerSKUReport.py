"""易仓 获取财务 SellerSKU 维度利润列表-新（getFinancialSellerSKUReportListNew）。

文档：https://open.eccang.com/#/documentCenter?docId=112282&catId=0-508-508,0-177

必填：companyCode、startTime、endTime（起止间隔 ≤ 31 天）。
pageSize 最大 500。报表存分析库，约每日 9 点前 / 14 点前更新，请控制请求频率。

运行（在项目根目录）::

    python api/eccang/request/getSellerSKUReport.py \\
        --start-time "2026-07-01 00:00:00" \\
        --end-time "2026-07-31 23:59:59"

    python api/eccang/request/getSellerSKUReport.py \\
        --start-time "2026-07-01 00:00:00" \\
        --end-time "2026-07-07 23:59:59" \\
        --page 1 --page-size 50 \\
        --user-account YOUR_ACCOUNT
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path
from typing import Any

DEFAULT_COMPANY_CODE = "ERP2009186VG"
DEFAULT_UNIT_CURRENCY = "EUR"
DEFAULT_TIME_ZONE_TYPE = 2  # 1=北京时间；2=站点时间
DEFAULT_TIME_TYPE = 3  # 1=下单时间；2=结算时间；3=Datetime；4=发货时间


def _bootstrap() -> None:
    root = Path(__file__).resolve().parents[3]
    epr = root / "ensure_project_root.py"
    spec = importlib.util.spec_from_file_location("ensure_project_root", epr)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    mod.bootstrap(__file__)


def _parse_csv(value: str | None) -> list[str] | None:
    if not value:
        return None
    items = [x.strip() for x in value.split(",") if x.strip()]
    return items or None


def _parse_int_csv(value: str | None) -> list[int] | None:
    items = _parse_csv(value)
    if not items:
        return None
    return [int(x) for x in items]


def get_financial_seller_sku_report_list(
    *,
    company_code: str = DEFAULT_COMPANY_CODE,
    start_time: str,
    end_time: str,
    page: int = 1,
    page_size: int = 50,
    unit_currency: str | None = DEFAULT_UNIT_CURRENCY,
    site_list: list[str] | None = None,
    user_account_list: list[str] | None = None,
    user_account: str | None = None,
    time_zone_type: int | None = DEFAULT_TIME_ZONE_TYPE,
    time_type: int | None = DEFAULT_TIME_TYPE,
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
    """调用 ``getFinancialSellerSKUReportListNew``，返回解析后的完整响应。"""
    from api.eccang import EccangService

    client = EccangService()
    return client.get_financial_seller_sku_report_list_new(
        company_code=company_code,
        start_time=start_time,
        end_time=end_time,
        page=page,
        page_size=page_size,
        unit_currency=unit_currency,
        site_list=site_list,
        user_account_list=user_account_list,
        user_account=user_account,
        time_zone_type=time_zone_type,
        time_type=time_type,
        seller_sku_item_status_list=seller_sku_item_status_list,
        cost_type=cost_type,
        profit_formula_type=profit_formula_type,
        search_type=search_type,
        keyword=keyword,
        seller_sku_list=seller_sku_list,
        asin_list=asin_list,
        parent_asin_list=parent_asin_list,
        transaction_status=transaction_status,
        charge_type=charge_type,
        account_skus=account_skus,
        system_code=system_code,
        extra=extra,
    )


def main(argv: list[str] | None = None) -> int:
    _bootstrap()
    from api.eccang.exceptions import EccangApiError, EccangConfigError

    parser = argparse.ArgumentParser(
        description=(
            "易仓 获取财务 SellerSKU 维度利润列表 "
            "getFinancialSellerSKUReportListNew（docId=112282）"
        ),
    )
    parser.add_argument(
        "--company-code",
        default=DEFAULT_COMPANY_CODE,
        dest="company_code",
        help=f"公司代码（biz_content.companyCode），默认 {DEFAULT_COMPANY_CODE}",
    )
    parser.add_argument(
        "--start-time",
        required=True,
        dest="start_time",
        help="报表开始时间，yyyy-MM-dd HH:mm:ss",
    )
    parser.add_argument(
        "--end-time",
        required=True,
        dest="end_time",
        help="报表结束时间，yyyy-MM-dd HH:mm:ss（与开始间隔 ≤ 31 天）",
    )
    parser.add_argument("--page", type=int, default=1, help="页码，默认 1")
    parser.add_argument(
        "--page-size",
        type=int,
        default=50,
        dest="page_size",
        help="每页条数，默认 50，最大 500",
    )
    parser.add_argument(
        "--unit-currency",
        default=DEFAULT_UNIT_CURRENCY,
        dest="unit_currency",
        help=f"币种：ORIGINAL/RMB/USD/EUR/JPY/GBP/CAD/MXN，默认 {DEFAULT_UNIT_CURRENCY}",
    )
    parser.add_argument(
        "--site-list",
        dest="site_list",
        help="站点列表，逗号分隔",
    )
    parser.add_argument(
        "--user-account-list",
        dest="user_account_list",
        help="平台账号列表，逗号分隔",
    )
    parser.add_argument(
        "--user-account",
        dest="user_account",
        help="单个平台账号",
    )
    parser.add_argument(
        "--time-zone-type",
        type=int,
        default=DEFAULT_TIME_ZONE_TYPE,
        dest="time_zone_type",
        help=f"时区类型：1=北京时间，2=站点时间，默认 {DEFAULT_TIME_ZONE_TYPE}",
    )
    parser.add_argument(
        "--time-type",
        type=int,
        default=DEFAULT_TIME_TYPE,
        dest="time_type",
        help=f"时间类型：1=下单时间，2=结算时间（接口实测亦接受 3/4），默认 {DEFAULT_TIME_TYPE}",
    )
    parser.add_argument(
        "--seller-sku-status-list",
        dest="seller_sku_item_status_list",
        help="销售状态列表，逗号分隔：1在售 2停售 3下架 4已删除",
    )
    parser.add_argument(
        "--cost-type",
        type=int,
        dest="cost_type",
        help="成本来源：1商品成本配置 2FBA进销存 4ERP先进先出 5月末加权",
    )
    parser.add_argument(
        "--profit-formula-type",
        type=int,
        dest="profit_formula_type",
        help="利润公式：1自定义 2系统默认",
    )
    parser.add_argument(
        "--search-type",
        type=int,
        dest="search_type",
        help="查询类型：1SellerSku 2子Asin 3父Asin 6品牌 7品类 10产品名称",
    )
    parser.add_argument("--keyword", help="与 --search-type 配合的搜索值")
    parser.add_argument(
        "--seller-sku-list",
        dest="seller_sku_list",
        help="SellerSku 列表，逗号分隔（最大 100）",
    )
    parser.add_argument(
        "--asin-list",
        dest="asin_list",
        help="子 ASIN 列表，逗号分隔（最大 100）",
    )
    parser.add_argument(
        "--parent-asin-list",
        dest="parent_asin_list",
        help="父 ASIN 列表，逗号分隔（最大 100）",
    )
    parser.add_argument(
        "--transaction-status",
        dest="transaction_status",
        help="交易状态：已发放 / 已推迟",
    )
    parser.add_argument(
        "--charge-type",
        dest="charge_type",
        help="汇率方式：ord=下单时间；settle=结算时间",
    )
    parser.add_argument(
        "--system-code",
        dest="system_code",
        help="系统编码，如 AMAZON_OPERATE",
    )
    parser.add_argument(
        "--account-skus",
        dest="account_skus",
        help='账户SKU组合 JSON 数组，如 [{"userAccount":"a","sellerSku":"s"}]',
    )
    parser.add_argument(
        "--body",
        help="额外 biz_content JSON 对象，会合并进请求（不覆盖必填字段）",
    )
    args = parser.parse_args(argv)

    account_skus: list[dict[str, str]] | None = None
    if args.account_skus:
        try:
            parsed = json.loads(args.account_skus)
        except json.JSONDecodeError as exc:
            print(f"[FAIL] --account-skus JSON 格式错误：{exc}", file=sys.stderr)
            return 2
        if not isinstance(parsed, list):
            print("[FAIL] --account-skus 必须是 JSON 数组", file=sys.stderr)
            return 2
        account_skus = parsed

    extra: dict[str, Any] | None = None
    if args.body:
        try:
            parsed_body = json.loads(args.body)
        except json.JSONDecodeError as exc:
            print(f"[FAIL] --body JSON 格式错误：{exc}", file=sys.stderr)
            return 2
        if not isinstance(parsed_body, dict):
            print("[FAIL] --body 必须是 JSON 对象", file=sys.stderr)
            return 2
        extra = parsed_body

    try:
        response = get_financial_seller_sku_report_list(
            company_code=args.company_code,
            start_time=args.start_time,
            end_time=args.end_time,
            page=args.page,
            page_size=args.page_size,
            unit_currency=args.unit_currency,
            site_list=_parse_csv(args.site_list),
            user_account_list=_parse_csv(args.user_account_list),
            user_account=args.user_account,
            time_zone_type=args.time_zone_type,
            time_type=args.time_type,
            seller_sku_item_status_list=_parse_int_csv(args.seller_sku_item_status_list),
            cost_type=args.cost_type,
            profit_formula_type=args.profit_formula_type,
            search_type=args.search_type,
            keyword=args.keyword,
            seller_sku_list=_parse_csv(args.seller_sku_list),
            asin_list=_parse_csv(args.asin_list),
            parent_asin_list=_parse_csv(args.parent_asin_list),
            transaction_status=args.transaction_status,
            charge_type=args.charge_type,
            account_skus=account_skus,
            system_code=args.system_code,
            extra=extra,
        )
    except EccangConfigError as exc:
        print(f"[FAIL] 配置错误：{exc}", file=sys.stderr)
        return 2
    except EccangApiError as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        if exc.code:
            print(f"  code={exc.code}", file=sys.stderr)
        if exc.raw_payload is not None:
            print(json.dumps(exc.raw_payload, ensure_ascii=False, indent=2))
        return 1
    except ValueError as exc:
        print(f"[FAIL] 参数错误：{exc}", file=sys.stderr)
        return 2

    printable = {k: v for k, v in response.items() if k != "biz_content"}
    data = printable.get("data") or {}
    records = None
    total = None
    if isinstance(data, dict):
        records = data.get("records")
        if records is None and isinstance(data.get("data"), dict):
            inner = data["data"]
            records = inner.get("records")
            total = inner.get("total")
        else:
            total = data.get("total")
    count = len(records) if isinstance(records, list) else 0
    print(
        f"[OK] method=getFinancialSellerSKUReportListNew version=1.0.0 "
        f"page={args.page} page_size={args.page_size} "
        f"returned={count} total={total}"
    )
    print(json.dumps(printable, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
