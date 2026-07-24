"""鸿羽 OMS 创建退件 ``createReturnBill``。

文档：http://oms.gindalogistik.com/api-doc/index.php（退件模块 → 创建退件）

支持两种模式：
- **标准退件**：需 ``tracking_no`` / ``warehouse_code`` / ``return_type`` / ``items``；
  ``return_type=S`` 时必填 ``order_code``，``C`` 时必填 ``claim_code``。
- **回邮退件**：传 ``return_identification=1``，需 ``reference_no`` / ``sm_code`` / ``sender_info``。

先在 ``api/hy_oms/config.py`` 填写凭证，再执行::

    python -m api.hy_oms.request.create_return ^
      --warehouse-code DEHY --return-type S --tracking-no T001 ^
      --order-code ORD001 --items SKU1:1:1

    python -m api.hy_oms.request.create_return --body path/to/body.json --raw
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path
from typing import Any, Mapping, Sequence, Union


RETURN_TYPES = {"S", "L", "C"}  # 买家 / 物流 / 认领
PROCESS_CODES = {1, 2, 3, 4, 5, 6, 8, 9}


def _bootstrap() -> None:
    root = Path(__file__).resolve().parents[3]
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


def _parse_tracking(raw: str | None) -> Union[str, list[str], None]:
    """单个跟踪号，或 ``T1;T2`` / ``T1,T2`` 多个。"""
    if raw is None or not str(raw).strip():
        return None
    text = str(raw).strip()
    if ";" in text:
        parts = [x.strip() for x in text.split(";") if x.strip()]
        return parts if len(parts) > 1 else (parts[0] if parts else None)
    if "," in text:
        parts = [x.strip() for x in text.split(",") if x.strip()]
        return parts if len(parts) > 1 else (parts[0] if parts else None)
    return text


def _parse_items(raw: str | None) -> list[dict[str, Any]] | None:
    """``SKU:数量:处理方式[:备注]``，逗号分隔。

    处理方式：1重新上架 2退回国内 3不良品 4销毁 5待检查 6换标 8产品升级 9直接销毁。
    """
    if not raw or not str(raw).strip():
        return None
    items: list[dict[str, Any]] = []
    for part in str(raw).split(","):
        part = part.strip()
        if not part:
            continue
        bits = [x.strip() for x in part.split(":")]
        if len(bits) < 3:
            raise ValueError(
                f"items 格式应为 SKU:数量:处理方式[:备注]，收到: {part}"
            )
        process = int(float(bits[2]))
        if process not in PROCESS_CODES:
            raise ValueError(
                f"process 非法: {process}，允许值 {sorted(PROCESS_CODES)}"
            )
        item: dict[str, Any] = {
            "product_sku": bits[0],
            "quantity": int(float(bits[1])),
            "process": str(process),
        }
        if len(bits) >= 4 and bits[3] != "":
            item["note"] = bits[3]
        items.append(item)
    return items or None


def _parse_sender_info(raw: str | None) -> dict[str, Any] | None:
    """从 JSON 字符串/文件解析 sender_info。"""
    if not raw or not str(raw).strip():
        return None
    data = _load_body(raw)
    return data


def _set_optional(params: dict[str, Any], mapping: Mapping[str, Any]) -> None:
    for key, value in mapping.items():
        if value is not None:
            params[key] = value


def create_return_bill(
    *,
    warehouse_code: str,
    items: Sequence[Mapping[str, Any]],
    tracking_no: Union[str, Sequence[str], None] = None,
    return_type: str | None = None,
    verify: int | str | None = None,
    reference_no: str | None = None,
    order_code: str | None = None,
    claim_code: str | None = None,
    expected_date: str | None = None,
    return_desc: str | None = None,
    operation_desc: str | None = None,
    buyer_name: str | None = None,
    buyers_ein: str | None = None,
    seller_store: str | None = None,
    images: Sequence[Mapping[str, Any]] | None = None,
    # 回邮退件
    return_identification: int | None = None,
    sm_code: str | None = None,
    sender_info: Mapping[str, Any] | None = None,
    **extra: Any,
) -> dict[str, Any]:
    """调用 ``createReturnBill``，返回完整响应 dict（含 ask / return_code）。

    标准退件与回邮退件由 ``return_identification`` 区分：
    - 未传或非 1 → 标准退件（需 tracking_no、return_type）
    - ``return_identification=1`` → 回邮退件（需 reference_no、sm_code、sender_info）
    """
    from api.hy_oms import HyOmsClient

    client = HyOmsClient.from_config()
    return client.create_return_bill(
        warehouse_code=warehouse_code,
        items=items,
        tracking_no=tracking_no,
        return_type=return_type,
        verify=verify,
        reference_no=reference_no,
        order_code=order_code,
        claim_code=claim_code,
        expected_date=expected_date,
        return_desc=return_desc,
        operation_desc=operation_desc,
        buyer_name=buyer_name,
        buyers_ein=buyers_ein,
        seller_store=seller_store,
        images=images,
        return_identification=return_identification,
        sm_code=sm_code,
        sender_info=sender_info,
        **extra,
    )


def _validate_standard(params: dict[str, Any]) -> None:
    missing = [
        k
        for k in ("tracking_no", "warehouse_code", "return_type", "items")
        if not params.get(k)
    ]
    if missing:
        raise ValueError(f"标准退件缺少必填: {', '.join(missing)}")
    rt = str(params["return_type"]).strip().upper()
    if rt not in RETURN_TYPES:
        raise ValueError(f"return_type 须为 S/L/C，收到: {params['return_type']}")
    params["return_type"] = rt
    if rt == "S" and not params.get("order_code"):
        raise ValueError("return_type=S（买家退件）时 order_code 必填")
    if rt == "C" and not params.get("claim_code"):
        raise ValueError("return_type=C（认领退件）时 claim_code 必填")


def _validate_mail(params: dict[str, Any]) -> None:
    missing = [
        k
        for k in (
            "reference_no",
            "warehouse_code",
            "sm_code",
            "items",
            "sender_info",
        )
        if not params.get(k)
    ]
    if missing:
        raise ValueError(f"回邮退件缺少必填: {', '.join(missing)}")
    sender = params["sender_info"]
    if not isinstance(sender, Mapping):
        raise ValueError("sender_info 必须是对象")
    for key in (
        "sender_name",
        "sender_country",
        "sender_email",
        "sender_phone",
        "sender_city",
        "sender_zipcode",
        "sender_address1",
        "sender_address2",
    ):
        if not sender.get(key):
            raise ValueError(f"回邮退件 sender_info.{key} 必填")


def build_params(
    *,
    warehouse_code: str,
    items: Sequence[Mapping[str, Any]],
    tracking_no: Union[str, Sequence[str], None] = None,
    return_type: str | None = None,
    verify: int | str | None = None,
    reference_no: str | None = None,
    order_code: str | None = None,
    claim_code: str | None = None,
    expected_date: str | None = None,
    return_desc: str | None = None,
    operation_desc: str | None = None,
    buyer_name: str | None = None,
    buyers_ein: str | None = None,
    seller_store: str | None = None,
    images: Sequence[Mapping[str, Any]] | None = None,
    return_identification: int | None = None,
    sm_code: str | None = None,
    sender_info: Mapping[str, Any] | None = None,
    validate: bool = True,
    **extra: Any,
) -> dict[str, Any]:
    """组装 ``createReturnBill`` 的 paramsJson（不发请求）。"""
    if not items:
        raise ValueError("items 不能为空")

    params: dict[str, Any] = {
        "warehouse_code": warehouse_code,
        "items": [dict(x) for x in items],
        **extra,
    }
    _set_optional(
        params,
        {
            "tracking_no": tracking_no
            if not isinstance(tracking_no, (list, tuple))
            else list(tracking_no),
            "return_type": return_type,
            "verify": verify,
            "reference_no": reference_no,
            "order_code": order_code,
            "claim_code": claim_code,
            "expected_date": expected_date,
            "return_desc": return_desc,
            "operation_desc": operation_desc,
            "buyer_name": buyer_name,
            "buyers_ein": buyers_ein,
            "seller_store": seller_store,
            "images": [dict(x) for x in images] if images is not None else None,
            "return_identification": return_identification,
            "sm_code": sm_code,
            "sender_info": dict(sender_info) if sender_info is not None else None,
        },
    )

    if validate:
        if int(params.get("return_identification") or 0) == 1:
            _validate_mail(params)
        else:
            _validate_standard(params)
    return params


def main(argv=None) -> int:
    _bootstrap()
    from api.hy_oms.exceptions import HyOmsError

    parser = argparse.ArgumentParser(
        description="鸿羽 OMS 创建退件 createReturnBill（标准 / 回邮）"
    )
    parser.add_argument("--warehouse-code", dest="warehouse_code", default=None)
    parser.add_argument(
        "--return-type",
        dest="return_type",
        default=None,
        help="标准退件类型：S买家 / L物流 / C认领",
    )
    parser.add_argument(
        "--tracking-no",
        dest="tracking_no",
        default=None,
        help="退件跟踪号；多个用 ; 或 , 分隔",
    )
    parser.add_argument("--order-code", dest="order_code", default=None, help="原订单号（S 必填）")
    parser.add_argument("--claim-code", dest="claim_code", default=None, help="认领单号（C 必填）")
    parser.add_argument("--reference-no", dest="reference_no", default=None)
    parser.add_argument(
        "--verify",
        type=int,
        choices=(0, 1),
        default=None,
        help="1确认审核 / 0草稿",
    )
    parser.add_argument("--expected-date", dest="expected_date", default=None)
    parser.add_argument("--return-desc", dest="return_desc", default=None)
    parser.add_argument("--operation-desc", dest="operation_desc", default=None)
    parser.add_argument("--buyer-name", dest="buyer_name", default=None)
    parser.add_argument("--buyers-ein", dest="buyers_ein", default=None)
    parser.add_argument("--seller-store", dest="seller_store", default=None)
    parser.add_argument(
        "--items",
        default=None,
        help="产品明细：SKU:数量:处理方式[:备注]，逗号分隔",
    )
    parser.add_argument(
        "--mail",
        action="store_true",
        help="回邮退件（return_identification=1）",
    )
    parser.add_argument("--sm-code", dest="sm_code", default=None, help="回邮物流产品代码")
    parser.add_argument(
        "--sender-info",
        dest="sender_info",
        default=None,
        help="寄件人 JSON 字符串或文件路径（回邮必填）",
    )
    parser.add_argument(
        "--body",
        default=None,
        help="完整 paramsJson（JSON 字符串或文件路径），优先于其它字段参数",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只打印将要发送的 paramsJson，不实际调用",
    )
    parser.add_argument("--raw", action="store_true", help="打印完整响应 JSON")
    args = parser.parse_args(argv)

    try:
        if args.body:
            params = _load_body(args.body)
            if int(params.get("return_identification") or 0) == 1:
                _validate_mail(params)
            else:
                _validate_standard(params)
        else:
            items = _parse_items(args.items)
            if not args.warehouse_code or not items:
                print(
                    "缺少 --warehouse-code / --items（或改用 --body）",
                    file=sys.stderr,
                )
                return 2
            params = build_params(
                warehouse_code=args.warehouse_code,
                items=items,
                tracking_no=_parse_tracking(args.tracking_no),
                return_type=args.return_type,
                verify=args.verify,
                reference_no=args.reference_no,
                order_code=args.order_code,
                claim_code=args.claim_code,
                expected_date=args.expected_date,
                return_desc=args.return_desc,
                operation_desc=args.operation_desc,
                buyer_name=args.buyer_name,
                buyers_ein=args.buyers_ein,
                seller_store=args.seller_store,
                return_identification=1 if args.mail else None,
                sm_code=args.sm_code,
                sender_info=_parse_sender_info(args.sender_info),
            )

        if args.dry_run:
            print("[DRY-RUN] createReturnBill paramsJson:")
            print(json.dumps(params, ensure_ascii=False, indent=2))
            return 0

        from api.hy_oms import HyOmsClient

        result = HyOmsClient.from_config().call("createReturnBill", params)
    except (ValueError, json.JSONDecodeError) as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 2
    except HyOmsError as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        raw = getattr(exc, "raw", None)
        if raw is not None:
            print(json.dumps(raw, ensure_ascii=False, indent=2)[:4000], file=sys.stderr)
        return 1

    return_code = result.get("return_code")
    print(f"[OK] service=createReturnBill ask={result.get('ask')} return_code={return_code}")
    payload = result if args.raw else {"return_code": return_code, "message": result.get("message")}
    print(json.dumps(payload, ensure_ascii=False, indent=2)[:8000])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
