"""易仓 WMS-获取产品列表（getWmsProductList）。

文档：https://open.eccang.com/#/documentCenter?docId=737&catId=0-187-187,0-177

运行（在项目根目录）::

    python -m api.eccang.getWmsProductList
    python -m api.eccang.getWmsProductList --page 1 --page-size 20
    python -m api.eccang.getWmsProductList --product-sku YOUR_SKU
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path
from typing import Any


def _bootstrap() -> None:
    root = Path(__file__).resolve().parents[2]
    epr = root / "ensure_project_root.py"
    spec = importlib.util.spec_from_file_location("ensure_project_root", epr)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    mod.bootstrap(__file__)


def get_wms_product_list(
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
    """调用 WMS ``getProductList``，返回解析后的完整响应。"""
    from api.eccang import EccangService

    client = EccangService()
    return client.get_wms_product_list(
        page=page,
        page_size=page_size,
        product_sku=product_sku,
        product_sku_like=product_sku_like,
        product_spu=product_spu,
        product_title_like=product_title_like,
        warehouse_barcode=warehouse_barcode,
        product_update_time_from=product_update_time_from,
        product_update_time_to=product_update_time_to,
        extra=extra,
    )


def main(argv: list[str] | None = None) -> int:
    _bootstrap()
    from api.eccang.exceptions import EccangApiError, EccangConfigError

    parser = argparse.ArgumentParser(
        description="易仓 WMS-获取产品列表 getWmsProductList（docId=737）",
    )
    parser.add_argument("--page", type=int, default=1, help="页码，默认 1")
    parser.add_argument(
        "--page-size",
        type=int,
        default=None,
        dest="page_size",
        help="每页条数，最大 1000；指定 --product-sku 时默认 1，否则默认 20",
    )
    parser.add_argument(
        "--product-sku",
        dest="product_sku",
        help="产品 SKU（精确匹配，对应 biz_content.product_sku）",
    )
    parser.add_argument(
        "--product-sku-like",
        dest="product_sku_like",
        help="产品 SKU 模糊查询",
    )
    parser.add_argument("--product-spu", dest="product_spu", help="产品款式代码")
    parser.add_argument(
        "--product-title-like",
        dest="product_title_like",
        help="产品名称模糊查询",
    )
    parser.add_argument(
        "--warehouse-barcode",
        dest="warehouse_barcode",
        help="仓库条码",
    )
    parser.add_argument(
        "--update-from",
        dest="product_update_time_from",
        help="产品更新时间起（YYYY-MM-DD HH:MM:SS）",
    )
    parser.add_argument(
        "--update-to",
        dest="product_update_time_to",
        help="产品更新时间止（YYYY-MM-DD HH:MM:SS）",
    )
    args = parser.parse_args(argv)

    # 精确查单个 SKU 时默认只取 1 条；列表查询默认 20
    if args.page_size is None:
        args.page_size = 1 if args.product_sku else 20

    try:
        response = get_wms_product_list(
            page=args.page,
            page_size=args.page_size,
            product_sku=args.product_sku,
            product_sku_like=args.product_sku_like,
            product_spu=args.product_spu,
            product_title_like=args.product_title_like,
            warehouse_barcode=args.warehouse_barcode,
            product_update_time_from=args.product_update_time_from,
            product_update_time_to=args.product_update_time_to,
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

    # 客户端已把 biz_content 解析到 data，打印时去掉原始字符串避免重复
    printable = {k: v for k, v in response.items() if k != "biz_content"}
    data = printable.get("data") or {}
    items = data.get("data") if isinstance(data, dict) else None
    total = data.get("total") if isinstance(data, dict) else None
    count = len(items) if isinstance(items, list) else 0
    print(
        f"[OK] method=getWmsProductList version=V1.0.0 "
        f"page={args.page} page_size={args.page_size} "
        f"returned={count} total={total}"
    )
    print(json.dumps(printable, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
