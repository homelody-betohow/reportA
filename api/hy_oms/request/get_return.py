"""鸿羽 OMS 获取退件详情 ``getReturnBill``。

文档：http://oms.gindalogistik.com/api-doc/index.php（退件模块 → 获取退件信息）

必填：``return_code``（退件单号）。

先在 ``api/hy_oms/config.py`` 填写凭证，再执行::

    python -m api.hy_oms.request.get_return --return-code RMA31-160930-0002
    python -m api.hy_oms.request.get_return --return-code RMA31-160930-0002 --raw
    python -m api.hy_oms.request.get_return --body path/to/body.json --dry-run
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path
from typing import Any, Mapping


# 退件状态（data.return_status）
RETURN_STATUS: dict[str, str] = {
    "C": "待确认",
    "W": "在途",
    "D": "到货",
    "E": "异常",
    "F": "已完成",
    "Q": "已作废",
    "A": "问题件",
    "B": "审核中",
    "G": "审核失败",
}

# 退件类型
RETURN_TYPES: dict[str, str] = {
    "S": "买家退件",
    "L": "物流退件",
    "C": "认领退件",
}


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


def build_params(
    *,
    return_code: str,
    validate: bool = True,
    **extra: Any,
) -> dict[str, Any]:
    """组装 ``getReturnBill`` 的 paramsJson（不发请求）。"""
    code = str(return_code or "").strip()
    params: dict[str, Any] = {"return_code": code, **extra}
    if validate and not code:
        raise ValueError("return_code（退件单号）必填")
    return params


def get_return_bill(
    *,
    return_code: str,
    **extra: Any,
) -> dict[str, Any]:
    """调用 ``getReturnBill``，返回完整响应 dict（含 ask / data）。"""
    from api.hy_oms import HyOmsClient

    client = HyOmsClient.from_config()
    return client.get_return_bill(return_code=return_code, **extra)


def summarize_return(data: Any) -> str:
    """从 data 提炼状态摘要，便于 CLI 一行展示。"""
    if not isinstance(data, Mapping):
        return str(data)
    code = data.get("return_code") or ""
    status = str(data.get("return_status") or "").strip().upper()
    status_label = RETURN_STATUS.get(status, status or "?")
    rtype = str(data.get("return_type") or "").strip().upper()
    type_label = RETURN_TYPES.get(rtype, rtype or "?")
    tracking = data.get("tracking_no") or ""
    mail = data.get("return_identification")
    mail_flag = "回邮" if str(mail) in {"1", "1.0", "True", "true"} else "标准"
    return (
        f"return_code={code} status={status}({status_label}) "
        f"type={rtype}({type_label}) mode={mail_flag} tracking_no={tracking}"
    )


def main(argv=None) -> int:
    _bootstrap()
    from api.hy_oms.exceptions import HyOmsError

    parser = argparse.ArgumentParser(
        description="鸿羽 OMS 获取退件详情 getReturnBill"
    )
    parser.add_argument(
        "--return-code",
        dest="return_code",
        default=None,
        help="退件单号（必填，除非用 --body）",
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
            if not str(params.get("return_code") or "").strip():
                raise ValueError("--body 中 return_code 必填")
        else:
            if not args.return_code:
                print("缺少 --return-code（或改用 --body）", file=sys.stderr)
                return 2
            params = build_params(return_code=args.return_code)

        if args.dry_run:
            print("[DRY-RUN] getReturnBill paramsJson:")
            print(json.dumps(params, ensure_ascii=False, indent=2))
            return 0

        from api.hy_oms import HyOmsClient

        result = HyOmsClient.from_config().call("getReturnBill", params)
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
    print(
        f"[OK] service=getReturnBill ask={result.get('ask')} "
        f"{summarize_return(data)}"
    )
    payload = result if args.raw else (data if data is not None else result)
    print(json.dumps(payload, ensure_ascii=False, indent=2)[:8000])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
