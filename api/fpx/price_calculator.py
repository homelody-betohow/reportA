"""4PX 费用试算（com.css.price_calculator）。

文档：https://open.4px.com/v2/doc/detail?ids=54,73,144

先在 ``api/fpx/config.py`` 填写凭证，再执行::

    python -m api.fpx.price_calculator ^
      --warehouse-code CNDGMA --country DE --weight 520 --length 20 --width 15 --height 10

    python -m api.fpx.price_calculator --body path/to/body.json --raw
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
import time
from pathlib import Path


def _bootstrap() -> None:
    root = Path(__file__).resolve().parents[2]
    epr = root / "ensure_project_root.py"
    spec = importlib.util.spec_from_file_location("ensure_project_root", epr)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    mod.bootstrap(__file__)


def _load_body(path_or_json: str) -> dict:
    p = Path(path_or_json)
    text = p.read_text(encoding="utf-8") if p.is_file() else path_or_json
    data = json.loads(text)
    if not isinstance(data, dict):
        raise ValueError("--body 必须是 JSON 对象")
    return data


def _parse_sku_list(raw: str | None) -> list | None:
    """``SKU:qty,SKU2:qty`` → ``[{"sku_code":..., "sku_qty":...}, ...]``。"""
    if not raw or not raw.strip():
        return None
    items = []
    for part in raw.split(","):
        part = part.strip()
        if not part:
            continue
        if ":" not in part:
            raise ValueError(f"sku 格式应为 SKU:数量，收到: {part}")
        code, qty = part.rsplit(":", 1)
        items.append({"sku_code": code.strip(), "sku_qty": float(qty.strip())})
    return items or None


def _parse_product_codes(raw: str | None) -> list | None:
    if not raw or not raw.strip():
        return None
    return [x.strip() for x in raw.split(",") if x.strip()]


def main(argv=None) -> int:
    _bootstrap()
    from api.fpx import FpxClient
    from api.fpx.config import FpxConfig
    from api.fpx.exceptions import FpxError

    parser = argparse.ArgumentParser(
        description="4PX 费用试算 com.css.price_calculator（ids=54,73,144）"
    )
    parser.add_argument(
        "--service-code",
        default="FB4",
        help="服务类别：目前仅支持 FB4（订单履约），默认 FB4",
    )
    parser.add_argument("--warehouse-code", dest="warehouse_code", default=None, help="仓库编码")
    parser.add_argument("--weight", type=float, default=None, help="包裹实重，单位 g")
    parser.add_argument("--length", type=float, default=None, help="长 cm")
    parser.add_argument("--width", type=float, default=None, help="宽 cm")
    parser.add_argument("--height", type=float, default=None, help="高 cm")
    parser.add_argument("--country", default=None, help="目的国二字码，如 DE/FR/CN")
    parser.add_argument("--post-code", dest="post_code", default=None, help="邮编")
    parser.add_argument("--state", default=None, help="州/省")
    parser.add_argument("--city", default=None, help="城市")
    parser.add_argument("--street", default=None, help="街道/详细地址")
    parser.add_argument(
        "--address-type",
        dest="address_type",
        default=None,
        help="RESIDENTIAL / BUSINESS",
    )
    parser.add_argument(
        "--billing-time",
        dest="billing_time",
        type=int,
        default=None,
        help="计费时间毫秒时间戳，默认当前时间",
    )
    parser.add_argument(
        "--product-codes",
        dest="product_codes",
        default=None,
        help="业务产品代码，逗号分隔",
    )
    parser.add_argument(
        "--sku-list",
        dest="sku_list",
        default=None,
        help="SKU 试算：SKU1:数量,SKU2:数量",
    )
    parser.add_argument(
        "--body",
        default=None,
        help="完整请求 JSON 字符串或文件路径（优先于其它字段参数）",
    )
    parser.add_argument("--raw", action="store_true", help="打印完整响应 JSON")
    parser.add_argument("--sandbox", action="store_true", help="使用沙箱域名")
    args = parser.parse_args(argv)

    try:
        cfg = FpxConfig.default()
        if args.sandbox:
            cfg = FpxConfig(
                app_key=cfg.app_key,
                app_secret=cfg.app_secret,
                access_token=cfg.access_token,
                base_url=cfg.base_url,
                api_version=cfg.api_version,
                language=cfg.language,
                format=cfg.format,
                timeout=cfg.timeout,
                sandbox=True,
            )
        client = FpxClient(cfg)
    except ValueError as exc:
        print(exc, file=sys.stderr)
        return 2

    try:
        if args.body:
            params = _load_body(args.body)
            result = client.call("com.css.price_calculator", params, version="1.0.0")
        else:
            missing = [
                name
                for name, val in [
                    ("--warehouse-code", args.warehouse_code),
                    ("--weight", args.weight),
                    ("--length", args.length),
                    ("--width", args.width),
                    ("--height", args.height),
                    ("--country", args.country),
                ]
                if val is None
            ]
            if missing:
                print(
                    "缺少必填参数: " + ", ".join(missing) + "（或改用 --body）",
                    file=sys.stderr,
                )
                return 2

            destination = {"country": str(args.country).strip().upper()}
            if args.post_code is not None:
                destination["post_code"] = args.post_code
            if args.state is not None:
                destination["state"] = args.state
            if args.city is not None:
                destination["city"] = args.city
            if args.street is not None:
                destination["street"] = args.street
            if args.address_type is not None:
                destination["address_type"] = args.address_type

            result = client.price_calculator(
                service_code=args.service_code,
                warehouse_code=args.warehouse_code,
                weight=args.weight,
                length=args.length,
                width=args.width,
                height=args.height,
                destination=destination,
                billing_time=args.billing_time
                if args.billing_time is not None
                else int(time.time() * 1000),
                product_codes=_parse_product_codes(args.product_codes),
                sku_list=_parse_sku_list(args.sku_list),
            )
    except (ValueError, json.JSONDecodeError) as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 2
    except FpxError as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 1

    print(
        f"[OK] method=com.css.price_calculator "
        f"result={result.get('result')} msg={result.get('msg')}"
    )
    payload = result if args.raw else result.get("data")
    if payload is not None:
        print(json.dumps(payload, ensure_ascii=False, indent=2)[:8000])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
