#!/usr/bin/env python
"""易仓 ERP API 命令行工具。

用法（在项目根目录）::

    python -m api.eccang.cli --help
    python -m api.eccang.cli test
    python -m api.eccang.cli get-warehouses
    python -m api.eccang.cli get-orders --start-time "2024-01-01 00:00:00"
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path


def _bootstrap() -> None:
    root = Path(__file__).resolve().parents[2]
    epr = root / "ensure_project_root.py"
    spec = importlib.util.spec_from_file_location("ensure_project_root", epr)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    mod.bootstrap(__file__)


def test_connection(args: argparse.Namespace) -> None:
    """测试易仓 ERP API 连接（获取仓库列表）。"""
    from api.eccang.config import EccangConfig
    from api.eccang.exceptions import EccangApiError, EccangConfigError
    from api.eccang.methods import EccangService

    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    try:
        config = EccangConfig.default()
        print(f"✓ 配置加载成功：app_key={config.app_key[:8]}***, service_id={config.service_id}")
        print(f"✓ 请求地址：{config.base_url}\n")

        client = EccangService(config)
        print("正在测试连接（获取仓库列表）...")
        resp = client.get_warehouse_list(page=1, page_size=10)

        print("✓ 连接成功！")
        print(f"响应：{json.dumps(resp, ensure_ascii=False, indent=2)}")
    except EccangConfigError as exc:
        print(f"✗ 配置错误：{exc}", file=sys.stderr)
        sys.exit(1)
    except EccangApiError as exc:
        print(f"✗ 接口调用失败：{exc}", file=sys.stderr)
        if exc.code:
            print(f"  错误代码：{exc.code}")
        sys.exit(1)
    except Exception as exc:
        print(f"✗ 未知错误：{exc}", file=sys.stderr)
        sys.exit(1)


def get_warehouses(args: argparse.Namespace) -> None:
    """获取仓库列表。"""
    from api.eccang.exceptions import EccangApiError
    from api.eccang.methods import EccangService

    try:
        client = EccangService()
        resp = client.get_warehouse_list(
            page=args.page,
            page_size=args.page_size,
        )
        print(json.dumps(resp, ensure_ascii=False, indent=2))
    except EccangApiError as exc:
        print(f"✗ 接口调用失败：{exc}", file=sys.stderr)
        sys.exit(1)


def get_orders(args: argparse.Namespace) -> None:
    """获取订单列表。"""
    from api.eccang.exceptions import EccangApiError
    from api.eccang.methods import EccangService

    try:
        client = EccangService()
        resp = client.get_order_list(
            page=args.page,
            page_size=args.page_size,
            start_time=args.start_time,
            end_time=args.end_time,
        )
        print(json.dumps(resp, ensure_ascii=False, indent=2))
    except EccangApiError as exc:
        print(f"✗ 接口调用失败：{exc}", file=sys.stderr)
        sys.exit(1)


def get_products(args: argparse.Namespace) -> None:
    """获取产品列表。"""
    from api.eccang.exceptions import EccangApiError
    from api.eccang.methods import EccangService

    try:
        client = EccangService()
        resp = client.get_product_list(
            page=args.page,
            page_size=args.page_size,
        )
        print(json.dumps(resp, ensure_ascii=False, indent=2))
    except EccangApiError as exc:
        print(f"✗ 接口调用失败：{exc}", file=sys.stderr)
        sys.exit(1)


def get_inventory(args: argparse.Namespace) -> None:
    """获取库存列表。"""
    from api.eccang.exceptions import EccangApiError
    from api.eccang.methods import EccangService

    try:
        client = EccangService()
        resp = client.get_inventory_list(
            warehouse_code=args.warehouse_code,
            page=args.page,
            page_size=args.page_size,
        )
        print(json.dumps(resp, ensure_ascii=False, indent=2))
    except EccangApiError as exc:
        print(f"✗ 接口调用失败：{exc}", file=sys.stderr)
        sys.exit(1)


def get_billing(args: argparse.Namespace) -> None:
    """获取账单列表。"""
    from api.eccang.exceptions import EccangApiError
    from api.eccang.methods import EccangService

    try:
        client = EccangService()
        resp = client.get_billing_list(
            start_date=args.start_date,
            end_date=args.end_date,
            page=args.page,
            page_size=args.page_size,
        )
        print(json.dumps(resp, ensure_ascii=False, indent=2))
    except EccangApiError as exc:
        print(f"✗ 接口调用失败：{exc}", file=sys.stderr)
        sys.exit(1)


def custom_call(args: argparse.Namespace) -> None:
    """调用任意接口方法。"""
    from api.eccang.exceptions import EccangApiError
    from api.eccang.methods import EccangService

    try:
        client = EccangService()

        body = None
        if args.body:
            try:
                body = json.loads(args.body)
            except json.JSONDecodeError as exc:
                print(f"✗ 请求体格式错误：{exc}", file=sys.stderr)
                sys.exit(1)

        resp = client.call(args.method, body)
        print(json.dumps(resp, ensure_ascii=False, indent=2))
    except EccangApiError as exc:
        print(f"✗ 接口调用失败：{exc}", file=sys.stderr)
        sys.exit(1)


def main() -> None:
    _bootstrap()

    parser = argparse.ArgumentParser(
        description="易仓 ERP API 命令行工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    subparsers = parser.add_subparsers(dest="command", help="子命令")

    parser_test = subparsers.add_parser("test", help="测试 API 连接")
    parser_test.set_defaults(func=test_connection)

    parser_warehouses = subparsers.add_parser("get-warehouses", help="获取仓库列表")
    parser_warehouses.add_argument("--page", type=int, default=1, help="页码")
    parser_warehouses.add_argument("--page-size", type=int, default=50, help="每页数量")
    parser_warehouses.set_defaults(func=get_warehouses)

    parser_orders = subparsers.add_parser("get-orders", help="获取订单列表")
    parser_orders.add_argument("--page", type=int, default=1, help="页码")
    parser_orders.add_argument("--page-size", type=int, default=50, help="每页数量")
    parser_orders.add_argument("--start-time", help="开始时间（YYYY-MM-DD HH:MM:SS）")
    parser_orders.add_argument("--end-time", help="结束时间（YYYY-MM-DD HH:MM:SS）")
    parser_orders.set_defaults(func=get_orders)

    parser_products = subparsers.add_parser("get-products", help="获取产品列表")
    parser_products.add_argument("--page", type=int, default=1, help="页码")
    parser_products.add_argument("--page-size", type=int, default=50, help="每页数量")
    parser_products.set_defaults(func=get_products)

    parser_inventory = subparsers.add_parser("get-inventory", help="获取库存列表")
    parser_inventory.add_argument("--warehouse-code", help="仓库编码")
    parser_inventory.add_argument("--page", type=int, default=1, help="页码")
    parser_inventory.add_argument("--page-size", type=int, default=50, help="每页数量")
    parser_inventory.set_defaults(func=get_inventory)

    parser_billing = subparsers.add_parser("get-billing", help="获取账单列表")
    parser_billing.add_argument("--start-date", help="开始日期（YYYY-MM-DD）")
    parser_billing.add_argument("--end-date", help="结束日期（YYYY-MM-DD）")
    parser_billing.add_argument("--page", type=int, default=1, help="页码")
    parser_billing.add_argument("--page-size", type=int, default=50, help="每页数量")
    parser_billing.set_defaults(func=get_billing)

    parser_call = subparsers.add_parser("call", help="调用任意接口方法")
    parser_call.add_argument("method", help="接口方法名")
    parser_call.add_argument("--body", help="请求体（JSON 格式）")
    parser_call.set_defaults(func=custom_call)

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(0)

    args.func(args)


if __name__ == "__main__":
    main()
