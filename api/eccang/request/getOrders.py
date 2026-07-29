"""易仓 WMS-查询仓储订单信息（getOrders）。

文档：https://open.eccang.com/#/documentCenter?docId=735&catId=0-181-181,0-177

interface_method=getOrders，version=V1.0.0，systemCode=WMS_MANAGER。

biz_content 常用筛选项均为可选；分页默认 page=1、page_size=10。

订单状态 order_status：
  0 删除 / 1 草稿 / 2 确认 / 3 缺货 / 4 已提交 / 5 已打印 / 7 已打包 / 8 已出库

运行（在项目根目录）::

    python -m api.eccang.request.getOrders
    python -m api.eccang.request.getOrders --page 1 --page-size 10
    python -m api.eccang.request.getOrders --order-status 8 --ship-from "2026-07-01 00:00:00"
    python -m api.eccang.request.getOrders --code SO311507140050,188444436012
    python -m api.eccang.request.getOrders --body path/to/body.json --dry-run
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path
from typing import Any, Mapping, Sequence


# 订单状态（data[].order_status）
ORDER_STATUS: dict[str, str] = {
    "0": "删除",
    "1": "草稿",
    "2": "确认",
    "3": "缺货",
    "4": "已提交",
    "5": "已打印",
    "7": "已打包",
    "8": "已出库",
}

METHOD = "getOrders"
VERSION = "V1.0.0"
DEFAULT_PAGE_SIZE = 10
MAX_PAGE_SIZE = 1000


def _bootstrap() -> None:
    root = Path(__file__).resolve().parents[3]
    epr = root / "ensure_project_root.py"
    spec = importlib.util.spec_from_file_location("ensure_project_root", epr)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    mod.bootstrap(__file__)


def _load_body(path_or_json: str) -> dict[str, Any]:
    p = Path(path_or_json)
    text = p.read_text(encoding="utf-8") if p.is_file() else path_or_json
    data = json.loads(text)
    if not isinstance(data, dict):
        raise ValueError("--body 必须是 JSON 对象")
    return data


def _parse_csv(value: str | None) -> list[str] | None:
    if value is None or not str(value).strip():
        return None
    text = str(value).strip()
    sep = ";" if ";" in text and "," not in text else ","
    items = [x.strip() for x in text.split(sep) if x.strip()]
    return items or None


def _parse_int_csv(value: str | None) -> list[int] | None:
    items = _parse_csv(value)
    if not items:
        return None
    return [int(x) for x in items]


def build_body(
    *,
    page: int = 1,
    page_size: int = DEFAULT_PAGE_SIZE,
    platform_arr: Sequence[str] | None = None,
    seller_id_arr: Sequence[str] | None = None,
    warehouse_id_arr: Sequence[int] | None = None,
    category: Sequence[int] | None = None,
    product_barcode_arr: Sequence[str] | None = None,
    sm_code_arr: Sequence[str] | None = None,
    country_code_in: Sequence[str] | None = None,
    code: Sequence[str] | None = None,
    order_status: str | int | None = None,
    addressee: str | None = None,
    buyer_id: Sequence[str] | None = None,
    buyer_name: str | None = None,
    buyer_mail: str | None = None,
    buyer_responsible_id: Sequence[int] | None = None,
    develop_responsible_id: Sequence[int] | None = None,
    seller_responsible_id: Sequence[int] | None = None,
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
    order_by: Sequence[str] | None = None,
    validate: bool = True,
    extra: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """组装 ``getOrders`` 的 biz_content（不发请求）。"""
    page_n = int(page)
    size_n = int(page_size)
    if validate:
        if page_n < 1:
            raise ValueError("page 必须 >= 1")
        if size_n < 1:
            raise ValueError("page_size 必须 >= 1")
        size_n = min(size_n, MAX_PAGE_SIZE)

    body: dict[str, Any] = {
        "page": page_n,
        "page_size": size_n,
    }
    optional: dict[str, Any] = {
        "platform_arr": list(platform_arr) if platform_arr else None,
        "seller_id_arr": list(seller_id_arr) if seller_id_arr else None,
        "warehouse_id_arr": list(warehouse_id_arr) if warehouse_id_arr else None,
        "category": list(category) if category else None,
        "product_barcode_arr": list(product_barcode_arr) if product_barcode_arr else None,
        "sm_code_arr": list(sm_code_arr) if sm_code_arr else None,
        "country_code_in": list(country_code_in) if country_code_in else None,
        "code": list(code) if code else None,
        "order_status": (
            str(order_status).strip() if order_status is not None and str(order_status).strip() != "" else None
        ),
        "addressee": (str(addressee).strip() if addressee is not None else None) or None,
        "buyer_id": list(buyer_id) if buyer_id else None,
        "buyer_name": (str(buyer_name).strip() if buyer_name is not None else None) or None,
        "buyer_mail": (str(buyer_mail).strip() if buyer_mail is not None else None) or None,
        "buyer_responsible_id": list(buyer_responsible_id) if buyer_responsible_id else None,
        "develop_responsible_id": list(develop_responsible_id) if develop_responsible_id else None,
        "seller_responsible_id": list(seller_responsible_id) if seller_responsible_id else None,
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
        "order_by": list(order_by) if order_by else None,
    }
    for key, value in optional.items():
        if value is not None and value != "":
            body[key] = value
    if extra:
        body.update(dict(extra))
        body["page"] = page_n
        body["page_size"] = size_n
    return body


def get_orders(
    *,
    page: int = 1,
    page_size: int = DEFAULT_PAGE_SIZE,
    platform_arr: Sequence[str] | None = None,
    seller_id_arr: Sequence[str] | None = None,
    warehouse_id_arr: Sequence[int] | None = None,
    category: Sequence[int] | None = None,
    product_barcode_arr: Sequence[str] | None = None,
    sm_code_arr: Sequence[str] | None = None,
    country_code_in: Sequence[str] | None = None,
    code: Sequence[str] | None = None,
    order_status: str | int | None = None,
    addressee: str | None = None,
    buyer_id: Sequence[str] | None = None,
    buyer_name: str | None = None,
    buyer_mail: str | None = None,
    buyer_responsible_id: Sequence[int] | None = None,
    develop_responsible_id: Sequence[int] | None = None,
    seller_responsible_id: Sequence[int] | None = None,
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
    order_by: Sequence[str] | None = None,
    extra: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """调用 WMS ``getOrders``，返回解析后的完整响应。"""
    from api.eccang import EccangService

    client = EccangService()
    return client.get_orders(
        page=page,
        page_size=page_size,
        platform_arr=list(platform_arr) if platform_arr else None,
        seller_id_arr=list(seller_id_arr) if seller_id_arr else None,
        warehouse_id_arr=list(warehouse_id_arr) if warehouse_id_arr else None,
        category=list(category) if category else None,
        product_barcode_arr=list(product_barcode_arr) if product_barcode_arr else None,
        sm_code_arr=list(sm_code_arr) if sm_code_arr else None,
        country_code_in=list(country_code_in) if country_code_in else None,
        code=list(code) if code else None,
        order_status=order_status,
        addressee=addressee,
        buyer_id=list(buyer_id) if buyer_id else None,
        buyer_name=buyer_name,
        buyer_mail=buyer_mail,
        buyer_responsible_id=list(buyer_responsible_id) if buyer_responsible_id else None,
        develop_responsible_id=(
            list(develop_responsible_id) if develop_responsible_id else None
        ),
        seller_responsible_id=list(seller_responsible_id) if seller_responsible_id else None,
        pay_date_for=pay_date_for,
        pay_date_to=pay_date_to,
        add_date_for=add_date_for,
        add_date_to=add_date_to,
        ship_date_for=ship_date_for,
        ship_date_to=ship_date_to,
        print_date_for=print_date_for,
        print_date_to=print_date_to,
        pack_date_for=pack_date_for,
        pack_date_to=pack_date_to,
        update_date_for=update_date_for,
        update_date_to=update_date_to,
        ec_update_time_for=ec_update_time_for,
        ec_update_time_to=ec_update_time_to,
        order_year=order_year,
        order_by=list(order_by) if order_by else None,
        extra=dict(extra) if extra else None,
    )


def _extract_page(payload: Mapping[str, Any]) -> tuple[list[Any], Any, Any, Any]:
    """从响应中取出订单列表与分页信息。

    客户端会把 biz_content 解析到 data；内层常见结构::

        data = {"data": [...], "page": 1, "page_size": 10, "total": "...", "next_page": "true"}
    """
    outer = payload.get("data")
    if not isinstance(outer, Mapping):
        return [], None, None, None
    inner = outer.get("data") if isinstance(outer.get("data"), Mapping) else outer
    if isinstance(inner, list):
        return inner, outer.get("total"), outer.get("page"), outer.get("next_page")
    if not isinstance(inner, Mapping):
        return [], None, None, None
    items = inner.get("data")
    if not isinstance(items, list):
        items = []
    return items, inner.get("total"), inner.get("page"), inner.get("next_page")


def main(argv: list[str] | None = None) -> int:
    _bootstrap()
    from api.eccang.exceptions import EccangApiError, EccangConfigError

    parser = argparse.ArgumentParser(
        description="易仓 WMS-查询仓储订单信息 getOrders（docId=735）",
    )
    parser.add_argument("--page", type=int, default=1, help="页码，默认 1")
    parser.add_argument(
        "--page-size",
        type=int,
        default=DEFAULT_PAGE_SIZE,
        dest="page_size",
        help=f"每页条数，默认 {DEFAULT_PAGE_SIZE}",
    )
    parser.add_argument(
        "--order-status",
        dest="order_status",
        help="订单状态：0删除/1草稿/2确认/3缺货/4已提交/5已打印/7已打包/8已出库",
    )
    parser.add_argument(
        "--code",
        help="单号（仓库单号/参考号/跟踪号，逗号分隔）",
    )
    parser.add_argument(
        "--platform",
        dest="platform_arr",
        help="平台代码，逗号分隔，如 ebay,amazon",
    )
    parser.add_argument(
        "--seller-id",
        dest="seller_id_arr",
        help="账号，逗号分隔",
    )
    parser.add_argument(
        "--warehouse-id",
        dest="warehouse_id_arr",
        help="仓库 Id，逗号分隔，如 10,11",
    )
    parser.add_argument(
        "--product-barcode",
        dest="product_barcode_arr",
        help="SKU，逗号分隔",
    )
    parser.add_argument(
        "--sm-code",
        dest="sm_code_arr",
        help="运输方式代码，逗号分隔",
    )
    parser.add_argument(
        "--country",
        dest="country_code_in",
        help="国家二字码，逗号分隔，如 DE,US",
    )
    parser.add_argument("--addressee", help="收件人（模糊）")
    parser.add_argument("--buyer-name", dest="buyer_name", help="买家姓名（模糊）")
    parser.add_argument("--buyer-mail", dest="buyer_mail", help="买家邮箱（模糊）")
    parser.add_argument(
        "--pay-from",
        dest="pay_date_for",
        help="付款时间起 YYYY-MM-DD HH:MM:SS",
    )
    parser.add_argument(
        "--pay-to",
        dest="pay_date_to",
        help="付款时间止 YYYY-MM-DD HH:MM:SS",
    )
    parser.add_argument(
        "--ship-from",
        dest="ship_date_for",
        help="发货时间起 YYYY-MM-DD HH:MM:SS",
    )
    parser.add_argument(
        "--ship-to",
        dest="ship_date_to",
        help="发货时间止 YYYY-MM-DD HH:MM:SS",
    )
    parser.add_argument(
        "--update-from",
        dest="update_date_for",
        help="更新时间起 YYYY-MM-DD HH:MM:SS",
    )
    parser.add_argument(
        "--update-to",
        dest="update_date_to",
        help="更新时间止 YYYY-MM-DD HH:MM:SS",
    )
    parser.add_argument(
        "--ec-update-from",
        dest="ec_update_time_for",
        help="EC 更新时间起 YYYY-MM-DD HH:MM:SS",
    )
    parser.add_argument(
        "--ec-update-to",
        dest="ec_update_time_to",
        help="EC 更新时间止 YYYY-MM-DD HH:MM:SS",
    )
    parser.add_argument(
        "--order-year",
        type=int,
        dest="order_year",
        help="历史订单年份（需配合拆表功能）",
    )
    parser.add_argument(
        "--order-by",
        dest="order_by",
        help='排序，逗号分隔，如 "order_id desc"',
    )
    parser.add_argument(
        "--body",
        help="完整 biz_content（JSON 字符串或文件路径），优先于其它字段参数",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只打印将要发送的 biz_content，不实际调用",
    )
    args = parser.parse_args(argv)

    try:
        if args.body:
            body = _load_body(args.body)
        else:
            body = build_body(
                page=args.page,
                page_size=args.page_size,
                platform_arr=_parse_csv(args.platform_arr),
                seller_id_arr=_parse_csv(args.seller_id_arr),
                warehouse_id_arr=_parse_int_csv(args.warehouse_id_arr),
                product_barcode_arr=_parse_csv(args.product_barcode_arr),
                sm_code_arr=_parse_csv(args.sm_code_arr),
                country_code_in=_parse_csv(args.country_code_in),
                code=_parse_csv(args.code),
                order_status=args.order_status,
                addressee=args.addressee,
                buyer_name=args.buyer_name,
                buyer_mail=args.buyer_mail,
                pay_date_for=args.pay_date_for,
                pay_date_to=args.pay_date_to,
                ship_date_for=args.ship_date_for,
                ship_date_to=args.ship_date_to,
                update_date_for=args.update_date_for,
                update_date_to=args.update_date_to,
                ec_update_time_for=args.ec_update_time_for,
                ec_update_time_to=args.ec_update_time_to,
                order_year=args.order_year,
                order_by=_parse_csv(args.order_by),
            )

        if args.dry_run:
            print(f"[DRY-RUN] method={METHOD} version={VERSION} biz_content:")
            print(json.dumps(body, ensure_ascii=False, indent=2))
            return 0

        from api.eccang import EccangService

        response = EccangService().call(METHOD, body, version=VERSION)
    except (ValueError, json.JSONDecodeError) as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 2
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

    printable = {k: v for k, v in response.items() if k != "biz_content"}
    items, total, page, next_page = _extract_page(printable)
    print(
        f"[OK] method={METHOD} version={VERSION} "
        f"page={page} returned={len(items)} total={total} next_page={next_page}"
    )
    print(json.dumps(printable, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
