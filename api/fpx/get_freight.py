"""查询 4PX 直发订单费用（ds.xms.order.getFreight）。

    python -m api.fpx.get_freight --request-no YOUR_NO
    python -m api.fpx.get_freight --request-no YOUR_NO --raw
    python -m api.fpx.get_freight --request-no YOUR_NO --sandbox
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
    spec.loader.exec_module(mod)
    mod.bootstrap(__file__)


def main(argv=None) -> int:
    _bootstrap()
    from api.fpx import FpxClient
    from api.fpx.config import FpxConfig
    from api.fpx.exceptions import FpxError

    parser = argparse.ArgumentParser(description="4PX 查询订单费用信息 ds.xms.order.getFreight")
    parser.add_argument(
        "--request-no",
        dest="request_no",
        required=True,
        help="请求单号（4PX 单号 / 服务商单号 / 客户单号）",
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
        result = client.get_freight(args.request_no)
    except FpxError as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 1

    print(
        f"[OK] method=ds.xms.order.getFreight "
        f"request_no={args.request_no} "
        f"result={result.get('result')} msg={result.get('msg')}"
    )
    payload = result if args.raw else result.get("data")
    if payload is not None:
        print(json.dumps(payload, ensure_ascii=False, indent=2)[:8000])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
