"""命令行冒烟：验证凭证与基础接口。

先在 ``api/hy_oms/config.py`` 填写 ``APP_TOKEN`` / ``APP_KEY``，再执行::

    python -m api.hy_oms.smoke_test
    python -m api.hy_oms.smoke_test --service getWarehouse
    python -m api.hy_oms.smoke_test --service getStorageCosts --date-for 2026-07-01 --date-to 2026-07-31
    python -m api.hy_oms.smoke_test --service getCalculateFee --warehouse-code HRBW --country-code DE --shipping-method F4 --weight 0.5
"""

from __future__ import annotations

import argparse
import json
import importlib.util
import sys
from pathlib import Path


def _bootstrap():
    root = Path(__file__).resolve().parents[2]
    epr = root / "ensure_project_root.py"
    spec = importlib.util.spec_from_file_location("ensure_project_root", epr)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    mod.bootstrap(__file__)


def main(argv=None) -> int:
    _bootstrap()
    from api.hy_oms import HyOmsClient
    from api.hy_oms.exceptions import HyOmsError

    parser = argparse.ArgumentParser(description="鸿羽 OMS API 冒烟测试")
    parser.add_argument("--service", default="getWarehouse", help="OMS service 名")
    parser.add_argument("--date-for", dest="date_for", default=None)
    parser.add_argument("--date-to", dest="date_to", default=None)
    parser.add_argument("--charge-date", dest="charge_date", default=None)
    parser.add_argument("--warehouse-code", dest="warehouse_code", default=None)
    parser.add_argument("--country-code", dest="country_code", default=None)
    parser.add_argument("--shipping-method", dest="shipping_method", default=None)
    parser.add_argument("--weight", type=float, default=None)
    parser.add_argument("--postcode", default=None)
    parser.add_argument("--page-size", type=int, default=5)
    parser.add_argument("--raw", action="store_true", help="打印完整 JSON")
    args = parser.parse_args(argv)

    try:
        client = HyOmsClient.from_env()
    except ValueError as exc:
        print(exc, file=sys.stderr)
        return 2

    params = {"page": 1, "pageSize": args.page_size}
    if args.date_for:
        params["dateFor"] = args.date_for
    if args.date_to:
        params["dateTo"] = args.date_to
    if args.charge_date:
        params["chargeDate"] = args.charge_date

    # 无业务参数的接口不传分页
    bare = {"getWarehouse", "getShippingMethod", "getFeeType", "getCurrency", "getBalance"}
    call_params = None if args.service in bare else params

    try:
        if args.service in {"getCalculateFee", "getCalculateFeeBatch"}:
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
                print(f"{args.service} 需要参数: {', '.join(missing)}", file=sys.stderr)
                return 2
            methods = [m.strip() for m in str(args.shipping_method).split(",") if m.strip()]
            if args.service == "getCalculateFee":
                result = client.get_calculate_fee(
                    warehouse_code=args.warehouse_code,
                    country_code=args.country_code,
                    shipping_method=methods[0],
                    weight=args.weight,
                    postcode=args.postcode,
                )
            else:
                result = client.get_calculate_fee_batch(
                    warehouse_code=args.warehouse_code,
                    country_code=args.country_code,
                    shipping_method=methods,
                    weight=args.weight,
                    postcode=args.postcode,
                )
        else:
            result = client.call(args.service, call_params)
    except HyOmsError as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 1

    print(f"[OK] service={args.service} ask={result.get('ask')}")
    data = result.get("data")
    if isinstance(data, list):
        print(f"rows={len(data)} total={result.get('total') or result.get('count')}")
        preview = data[: min(3, len(data))]
        print(json.dumps(preview, ensure_ascii=False, indent=2))
    elif args.raw or data is not None:
        print(json.dumps(result if args.raw else data, ensure_ascii=False, indent=2)[:4000])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
