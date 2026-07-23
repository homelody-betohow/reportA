#!/usr/bin/env python
"""易仓 ERP API 使用示例。

演示如何使用易仓 ERP API 客户端获取各类数据。

使用前请先在 ``api/eccang/config.py`` 中配置：
- APP_KEY
- APP_SECRET
- SERVICE_ID

运行（在项目根目录）::

    python -m api.eccang.example
"""

from __future__ import annotations

import importlib.util
import json
import sys
from datetime import datetime, timedelta
from pathlib import Path


def _bootstrap() -> None:
    root = Path(__file__).resolve().parents[2]
    epr = root / "ensure_project_root.py"
    spec = importlib.util.spec_from_file_location("ensure_project_root", epr)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    mod.bootstrap(__file__)


def print_json(data: dict, title: str = "") -> None:
    """美化打印 JSON 数据。"""
    if title:
        print(f"\n{'='*60}")
        print(f"  {title}")
        print(f"{'='*60}")
    print(json.dumps(data, ensure_ascii=False, indent=2))


def example_basic_usage() -> None:
    """示例 1：基本使用。"""
    from api.eccang import EccangService
    from api.eccang.exceptions import EccangApiError, EccangConfigError

    print("\n【示例 1】基本使用：获取仓库列表")

    try:
        client = EccangService()
        response = client.get_warehouse_list(page=1, page_size=10)
        print_json(response, "仓库列表")

        if response.get("code") == "0":
            data = response.get("data", {})
            warehouse_list = data.get("list", [])
            total = data.get("total", 0)
            print(f"\n✓ 成功获取 {len(warehouse_list)} 个仓库（共 {total} 个）")
        else:
            print(f"\n✗ 获取失败：{response.get('message')}")

    except EccangConfigError as e:
        print(f"✗ 配置错误：{e}")
        print("\n请在 api/eccang/config.py 中填写 API 凭证")
    except EccangApiError as e:
        print(f"✗ 接口调用失败：{e}")
        if e.code:
            print(f"  错误代码：{e.code}")


def example_get_orders() -> None:
    """示例 2：获取订单列表。"""
    from api.eccang import EccangService
    from api.eccang.exceptions import EccangApiError

    print("\n【示例 2】获取订单列表（最近 7 天）")

    try:
        client = EccangService()

        end_time = datetime.now()
        start_time = end_time - timedelta(days=7)
        start_time_str = start_time.strftime("%Y-%m-%d 00:00:00")
        end_time_str = end_time.strftime("%Y-%m-%d 23:59:59")

        print(f"时间范围：{start_time_str} ~ {end_time_str}")

        response = client.get_order_list(
            page=1,
            page_size=10,
            start_time=start_time_str,
            end_time=end_time_str,
        )

        print_json(response, "订单列表")

        if response.get("code") == "0":
            data = response.get("data", {})
            order_list = data.get("list", [])
            total = data.get("total", 0)
            print(f"\n✓ 成功获取 {len(order_list)} 个订单（共 {total} 个）")
        else:
            print(f"\n✗ 获取失败：{response.get('message')}")

    except EccangApiError as e:
        print(f"✗ 接口调用失败：{e}")


def example_get_inventory() -> None:
    """示例 3：获取库存列表。"""
    from api.eccang import EccangService
    from api.eccang.exceptions import EccangApiError

    print("\n【示例 3】获取库存列表")

    try:
        client = EccangService()
        response = client.get_inventory_list(
            page=1,
            page_size=10,
            # warehouse_code="WH001",
        )

        print_json(response, "库存列表")

        if response.get("code") == "0":
            data = response.get("data", {})
            inventory_list = data.get("list", [])
            total = data.get("total", 0)
            print(f"\n✓ 成功获取 {len(inventory_list)} 个库存记录（共 {total} 个）")
        else:
            print(f"\n✗ 获取失败：{response.get('message')}")

    except EccangApiError as e:
        print(f"✗ 接口调用失败：{e}")


def example_get_products() -> None:
    """示例 4：获取产品列表。"""
    from api.eccang import EccangService
    from api.eccang.exceptions import EccangApiError

    print("\n【示例 4】获取产品列表")

    try:
        client = EccangService()
        response = client.get_product_list(page=1, page_size=10)

        print_json(response, "产品列表")

        if response.get("code") == "0":
            data = response.get("data", {})
            product_list = data.get("list", [])
            total = data.get("total", 0)
            print(f"\n✓ 成功获取 {len(product_list)} 个产品（共 {total} 个）")
        else:
            print(f"\n✗ 获取失败：{response.get('message')}")

    except EccangApiError as e:
        print(f"✗ 接口调用失败：{e}")


def example_get_billing() -> None:
    """示例 5：获取账单列表。"""
    from api.eccang import EccangService
    from api.eccang.exceptions import EccangApiError

    print("\n【示例 5】获取账单列表（本月）")

    try:
        client = EccangService()

        today = datetime.now()
        start_date = today.replace(day=1).strftime("%Y-%m-%d")
        end_date = today.strftime("%Y-%m-%d")

        print(f"时间范围：{start_date} ~ {end_date}")

        response = client.get_billing_list(
            start_date=start_date,
            end_date=end_date,
            page=1,
            page_size=10,
        )

        print_json(response, "账单列表")

        if response.get("code") == "0":
            data = response.get("data", {})
            billing_list = data.get("list", [])
            total = data.get("total", 0)
            print(f"\n✓ 成功获取 {len(billing_list)} 条账单（共 {total} 条）")
        else:
            print(f"\n✗ 获取失败：{response.get('message')}")

    except EccangApiError as e:
        print(f"✗ 接口调用失败：{e}")


def example_custom_call() -> None:
    """示例 6：调用自定义接口。"""
    from api.eccang import EccangService
    from api.eccang.exceptions import EccangApiError

    print("\n【示例 6】调用自定义接口（获取发货地址簿）")

    try:
        client = EccangService()
        response = client.call(
            method="getShipAddressBooks",
            body={
                "page": 1,
                "page_size": 10,
            },
        )

        print_json(response, "发货地址簿")

        if response.get("code") == "0":
            data = response.get("data", {})
            address_list = data.get("list", [])
            total = data.get("total", 0)
            print(f"\n✓ 成功获取 {len(address_list)} 个地址（共 {total} 个）")
        else:
            print(f"\n✗ 获取失败：{response.get('message')}")

    except EccangApiError as e:
        print(f"✗ 接口调用失败：{e}")


def example_pagination() -> None:
    """示例 7：分页获取所有数据。"""
    from api.eccang import EccangService
    from api.eccang.exceptions import EccangApiError

    print("\n【示例 7】分页获取所有仓库数据")

    try:
        client = EccangService()

        all_warehouses = []
        page = 1
        page_size = 50

        while True:
            print(f"正在获取第 {page} 页...")

            response = client.get_warehouse_list(page=page, page_size=page_size)

            if response.get("code") != "0":
                print(f"✗ 获取失败：{response.get('message')}")
                break

            data = response.get("data", {})
            warehouse_list = data.get("list", [])
            total = data.get("total", 0)

            if not warehouse_list:
                break

            all_warehouses.extend(warehouse_list)
            print(f"  已获取 {len(all_warehouses)}/{total} 个仓库")

            if len(all_warehouses) >= total:
                break

            page += 1

        print(f"\n✓ 成功获取所有 {len(all_warehouses)} 个仓库")

    except EccangApiError as e:
        print(f"✗ 接口调用失败：{e}")


def main() -> None:
    """运行所有示例。"""
    _bootstrap()

    print("=" * 60)
    print("  易仓 ERP API 使用示例")
    print("=" * 60)

    # 运行示例 1：基本使用
    example_basic_usage()

    # 运行示例 2：获取订单
    # example_get_orders()

    # 运行示例 3：获取库存
    # example_get_inventory()

    # 运行示例 4：获取产品
    # example_get_products()

    # 运行示例 5：获取账单
    # example_get_billing()

    # 运行示例 6：自定义接口调用
    # example_custom_call()

    # 运行示例 7：分页获取数据
    # example_pagination()

    print("\n" + "=" * 60)
    print("  示例运行完成")
    print("=" * 60)
    print("\n提示：取消注释其他示例函数来运行更多示例")


if __name__ == "__main__":
    main()
