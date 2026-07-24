"""鸿羽 OMS 运费试算（getCalculateFee / getCalculateFeeBatch）。

文档：http://oms.gindalogistik.com/api-doc/index.php（费用模块 → 运费试算）

先在 ``api/hy_oms/config.py`` 填写凭证，再执行::

    python -m api.hy_oms.fee_calculate ^
      --warehouse-code DEHY --country-code DE --shipping-method D4 --weight 1.7

    python -m api.hy_oms.fee_calculate ^
      --warehouse-code DEHY --country-code DE --shipping-method D4,D1 ^
      --weight 1.7 --postcode 10115 --length 30 --width 20 --height 10 --batch

    python -m api.hy_oms.fee_calculate --body path/to/body.json --raw
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


def _load_body(path_or_json: str) -> dict:
    p = Path(path_or_json)
    text = p.read_text(encoding="utf-8") if p.is_file() else path_or_json
    data = json.loads(text)
    if not isinstance(data, dict):
        raise ValueError("--body 必须是 JSON 对象")
    return data


def _parse_items(raw: str | None) -> list[dict[str, Any]] | None:
    """``SKU:qty`` 或 ``SKU:qty:declared``，逗号分隔 → items 数组。"""
    if not raw or not raw.strip():
        return None
    items: list[dict[str, Any]] = []
    for part in raw.split(","):
        part = part.strip()
        if not part:
            continue
        bits = [x.strip() for x in part.split(":")]
        if len(bits) < 2:
            raise ValueError(f"items 格式应为 SKU:数量[:申报价值]，收到: {part}")
        item: dict[str, Any] = {
            "product_sku": bits[0],
            "quantity": int(float(bits[1])),
        }
        if len(bits) >= 3 and bits[2] != "":
            item["product_declared_value"] = float(bits[2])
        items.append(item)
    return items or None


def _summarize_fee(data: Any) -> str:
    """从 data 对象或批量列表提炼 totalFee / currency。"""
    if isinstance(data, list):
        parts = []
        for row in data:
            if isinstance(row, dict):
                fee = row.get("totalFee")
                cur = row.get("currency_code") or row.get("currency") or ""
                method = row.get("shipping_method") or row.get("sm_code") or ""
                label = f"{method}=" if method else ""
                parts.append(f"{label}{fee} {cur}".strip())
            else:
                parts.append(str(row))
        return "; ".join(parts) if parts else "(empty)"
    if isinstance(data, dict):
        fee = data.get("totalFee")
        cur = data.get("currency_code") or data.get("currency") or ""
        return f"totalFee={fee} {cur}".strip()
    return str(data)


def calculate_fee(
    *,
    warehouse_code: str,
    country_code: str,
    shipping_method: str,
    weight: float,
    postcode: str | None = None,
    length: float | None = None,
    width: float | None = None,
    height: float | None = None,
    city: str | None = None,
    state: str | None = None,
    address1: str | None = None,
    name: str | None = None,
    phone: str | None = None,
    self_lifting: int | None = None,
    items: list[dict[str, Any]] | None = None,
    **extra: Any,
) -> dict[str, Any]:
    """调用 ``getCalculateFee``，返回完整响应 dict（含 ask / data）。"""
    from api.hy_oms import HyOmsClient

    client = HyOmsClient.from_config()
    return client.get_calculate_fee(
        warehouse_code=warehouse_code,
        country_code=country_code,
        shipping_method=shipping_method,
        weight=weight,
        postcode=postcode,
        length=length,
        width=width,
        height=height,
        city=city,
        state=state,
        address1=address1,
        name=name,
        phone=phone,
        self_lifting=self_lifting,
        items=items,
        **extra,
    )


def calculate_fee_batch(
    *,
    warehouse_code: str,
    country_code: str,
    shipping_method: list[str] | str,
    weight: float,
    postcode: str | None = None,
    length: float | None = None,
    width: float | None = None,
    height: float | None = None,
    state: str | None = None,
    address1: str | None = None,
    self_lifting: int | None = None,
    items: list[dict[str, Any]] | None = None,
    **extra: Any,
) -> dict[str, Any]:
    """调用 ``getCalculateFeeBatch``，返回完整响应 dict。"""
    from api.hy_oms import HyOmsClient

    client = HyOmsClient.from_config()
    return client.get_calculate_fee_batch(
        warehouse_code=warehouse_code,
        country_code=country_code,
        shipping_method=shipping_method,
        weight=weight,
        postcode=postcode,
        length=length,
        width=width,
        height=height,
        state=state,
        address1=address1,
        self_lifting=self_lifting,
        items=items,
        **extra,
    )


def main(argv=None) -> int:
    _bootstrap()
    from api.hy_oms.exceptions import HyOmsError

    parser = argparse.ArgumentParser(
        description="鸿羽 OMS 运费试算 getCalculateFee / getCalculateFeeBatch"
    )
    parser.add_argument("--warehouse-code", dest="warehouse_code", default=None, help="配送仓库代码")
    parser.add_argument("--country-code", dest="country_code", default=None, help="目的国家代码，如 DE")
    parser.add_argument(
        "--shipping-method",
        dest="shipping_method",
        default=None,
        help="配送方式代码；批量时逗号分隔多个",
    )
    parser.add_argument("--weight", type=float, default=None, help="包裹重量（kg）")
    parser.add_argument("--postcode", default=None, help="收件人邮编")
    parser.add_argument("--length", type=float, default=None, help="长 cm")
    parser.add_argument("--width", type=float, default=None, help="宽 cm")
    parser.add_argument("--height", type=float, default=None, help="高 cm")
    parser.add_argument("--city", default=None, help="城市")
    parser.add_argument("--state", default=None, help="州/省")
    parser.add_argument("--address1", default=None, help="地址1")
    parser.add_argument("--name", default=None, help="收件人姓名")
    parser.add_argument("--phone", default=None, help="电话")
    parser.add_argument(
        "--self-lifting",
        dest="self_lifting",
        type=int,
        choices=(0, 1),
        default=None,
        help="是否自提：0 否 / 1 是",
    )
    parser.add_argument(
        "--items",
        default=None,
        help="产品明细：SKU:数量 或 SKU:数量:申报价值，逗号分隔",
    )
    parser.add_argument(
        "--body",
        default=None,
        help="完整 paramsJson（JSON 字符串或文件路径），优先于其它字段参数",
    )
    parser.add_argument(
        "--batch",
        action="store_true",
        help="调用 getCalculateFeeBatch（shipping_method 可为多个）",
    )
    parser.add_argument("--raw", action="store_true", help="打印完整响应 JSON")
    args = parser.parse_args(argv)

    try:
        if args.body:
            params = _load_body(args.body)
            from api.hy_oms import HyOmsClient

            client = HyOmsClient.from_config()
            service = "getCalculateFeeBatch" if args.batch else "getCalculateFee"
            # body 里若已是列表 shipping_method，自动走 batch
            sm = params.get("shipping_method")
            if isinstance(sm, list) and not args.batch:
                service = "getCalculateFeeBatch"
            result = client.call(service, params)
        else:
            missing = [
                name
                for name, val in [
                    ("--warehouse-code", args.warehouse_code),
                    ("--country-code", args.country_code),
                    ("--shipping-method", args.shipping_method),
                    ("--weight", args.weight),
                ]
                if val is None
            ]
            if missing:
                print(
                    "缺少必填参数: " + ", ".join(missing) + "（或改用 --body）",
                    file=sys.stderr,
                )
                return 2

            methods = [m.strip() for m in str(args.shipping_method).split(",") if m.strip()]
            items = _parse_items(args.items)
            common = dict(
                warehouse_code=args.warehouse_code,
                country_code=str(args.country_code).strip().upper(),
                weight=float(args.weight),
                postcode=args.postcode,
                length=args.length,
                width=args.width,
                height=args.height,
                self_lifting=args.self_lifting,
                items=items,
            )
            if args.batch or len(methods) > 1:
                result = calculate_fee_batch(
                    shipping_method=methods,
                    state=args.state,
                    address1=args.address1,
                    **common,
                )
                service = "getCalculateFeeBatch"
            else:
                result = calculate_fee(
                    shipping_method=methods[0],
                    city=args.city,
                    state=args.state,
                    address1=args.address1,
                    name=args.name,
                    phone=args.phone,
                    **common,
                )
                service = "getCalculateFee"
    except (ValueError, json.JSONDecodeError) as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 2
    except HyOmsError as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        raw = getattr(exc, "raw", None)
        if raw is not None:
            print(json.dumps(raw, ensure_ascii=False, indent=2)[:4000], file=sys.stderr)
        return 1

    data = result.get("data")
    print(f"[OK] service={service} ask={result.get('ask')} {_summarize_fee(data)}")
    payload = result if args.raw else data
    if payload is not None:
        print(json.dumps(payload, ensure_ascii=False, indent=2)[:8000])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
