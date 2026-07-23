"""命令行冒烟：验证凭证与直发费用查询等接口。

先在 ``api/fpx/config.py`` 填写 ``APP_KEY`` / ``APP_SECRET``，再执行::

    python -m api.fpx.smoke_test
    python -m api.fpx.smoke_test --method ds.xms.order.getFreight --request-no YOUR_NO
    python -m api.fpx.smoke_test --method com.basis.warehouse.getlist
    python -m api.fpx.smoke_test --check-sign
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path


def _bootstrap():
    root = Path(__file__).resolve().parents[2]
    epr = root / "ensure_project_root.py"
    spec = importlib.util.spec_from_file_location("ensure_project_root", epr)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    mod.bootstrap(__file__)


def _check_sign() -> int:
    """对照官方接入文档示例校验签名算法。"""
    from api.fpx.client import build_sign

    sign = build_sign(
        app_key="16081f05-e8fc-4250-b9c4-0660d1ecbb28",
        app_secret="7eebf328-8e5a-4030-904d-ec6e89174fbc",
        method="ds.xms.order.create",
        timestamp="1532592413187",
        body='{"aa":"bb"}',
        version="1.0",
    )
    expected = "ff4af77c062a9b97d98aa29777621c4a"
    if sign != expected:
        print(f"[FAIL] sign mismatch: got={sign} expected={expected}", file=sys.stderr)
        return 1
    print(f"[OK] sign algorithm matches official sample: {sign}")
    return 0


def main(argv=None) -> int:
    _bootstrap()
    from api.fpx import FpxClient
    from api.fpx.exceptions import FpxError

    parser = argparse.ArgumentParser(description="4PX 开放平台 API 冒烟测试")
    parser.add_argument(
        "--method",
        default="ds.xms.order.getFreight",
        help="API method（默认直发费用查询）",
    )
    parser.add_argument("--request-no", dest="request_no", default=None, help="请求单号")
    parser.add_argument("--delivery-order-no", dest="delivery_order_no", default=None)
    parser.add_argument("--version", default=None, help="覆盖接口版本号，如 1.0.0")
    parser.add_argument("--body", default=None, help="原始 JSON body，优先于其它参数")
    parser.add_argument("--check-sign", action="store_true", help="仅校验签名算法示例")
    parser.add_argument("--raw", action="store_true", help="打印完整 JSON")
    parser.add_argument("--sandbox", action="store_true", help="使用沙箱域名")
    args = parser.parse_args(argv)

    if args.check_sign:
        return _check_sign()

    try:
        from api.fpx.config import FpxConfig

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

    if args.body:
        try:
            params = json.loads(args.body)
        except json.JSONDecodeError as exc:
            print(f"无效 --body JSON: {exc}", file=sys.stderr)
            return 2
    else:
        params = {}
        if args.request_no:
            params["request_no"] = args.request_no
        if args.delivery_order_no:
            params["deliveryOrderNo"] = args.delivery_order_no

    need_no = {
        "ds.xms.order.getFreight",
        "ds.xms.order.cancel",
        "ds.xms.label.get",
    }
    if args.method in need_no and "request_no" not in params and not args.body:
        print(f"{args.method} 需要 --request-no", file=sys.stderr)
        return 2
    if args.method == "tr.order.tracking.get" and "deliveryOrderNo" not in params and not args.body:
        print("tr.order.tracking.get 需要 --delivery-order-no", file=sys.stderr)
        return 2

    try:
        if args.method == "ds.xms.order.getFreight" and args.request_no and not args.body:
            result = client.get_freight(args.request_no)
        elif args.method == "tr.order.tracking.get" and args.delivery_order_no and not args.body:
            result = client.get_tracking(args.delivery_order_no)
        else:
            result = client.call(args.method, params or None, version=args.version)
    except FpxError as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 1

    print(f"[OK] method={args.method} result={result.get('result')} msg={result.get('msg')}")
    data = result.get("data")
    if args.raw:
        print(json.dumps(result, ensure_ascii=False, indent=2)[:8000])
    elif data is not None:
        print(json.dumps(data, ensure_ascii=False, indent=2)[:4000])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
