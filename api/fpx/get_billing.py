"""4PX 费用查询（com.basis.billing.getbilling）。

文档：https://open.4px.com/v2/doc/detail?ids=54,74,159

先在 ``api/fpx/config.py`` 填写凭证，再执行::

    python -m api.fpx.get_billing --business-type O --order-no YOUR_ORDER_NO
    python -m api.fpx.get_billing --business-type O --ref-no YOUR_REF_NO --raw
    python -m api.fpx.get_billing --business-type L --order-no YOUR_NO --sandbox
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path

_BUSINESS_TYPES = {
    "I": "入库委托",
    "O": "出库委托",
    "T": "调拨委托",
    "L": "尾程管理运单",
}


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

    parser = argparse.ArgumentParser(
        description="4PX 费用查询 com.basis.billing.getbilling（ids=54,74,159）"
    )
    parser.add_argument(
        "--business-type",
        dest="business_type",
        required=True,
        choices=sorted(_BUSINESS_TYPES),
        help="业务类型：I 入库 / O 出库 / T 调拨 / L 尾程管理运单",
    )
    parser.add_argument("--order-no", dest="order_no", default=None, help="业务单号")
    parser.add_argument("--ref-no", dest="ref_no", default=None, help="参考号")
    parser.add_argument("--raw", action="store_true", help="打印完整响应 JSON")
    parser.add_argument("--sandbox", action="store_true", help="使用沙箱域名")
    args = parser.parse_args(argv)

    if not (args.order_no or "").strip() and not (args.ref_no or "").strip():
        print("需要 --order-no 或 --ref-no 之一", file=sys.stderr)
        return 2

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
        result = client.get_billing(
            business_type=args.business_type,
            order_no=args.order_no,
            ref_no=args.ref_no,
        )
    except (ValueError, FpxError) as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 1

    print(
        f"[OK] method=com.basis.billing.getbilling "
        f"business_type={args.business_type}({_BUSINESS_TYPES[args.business_type]}) "
        f"order_no={args.order_no or '-'} ref_no={args.ref_no or '-'} "
        f"result={result.get('result')} msg={result.get('msg')}"
    )
    if args.raw:
        payload = result
    else:
        payload = result.get("data")
        if payload is None and "billinglist" in result:
            payload = result.get("billinglist")
    if payload is not None:
        print(json.dumps(payload, ensure_ascii=False, indent=2)[:8000])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
