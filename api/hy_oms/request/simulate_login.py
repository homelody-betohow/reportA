"""鸿羽 OMS 模拟登录 ``logOn`` / 获取登陆 token ``getSsoToken``。

文档：http://oms.gindalogistik.com/api-doc/index.php（用户模块）

- **logOn**（默认）：``user_account`` + ``user_password`` → URL 编码的快捷登录地址
- **getSsoToken**：``company_code`` → ``userCode`` + ``token``，再拼快捷登录 URL

快捷登录路径::

    {base}/default/index/quick-login?userCode=...&token=...

先在 ``api/hy_oms/config.py`` 填写 ``APP_TOKEN`` / ``APP_KEY``，
以及可选的 ``USER_ACCOUNT`` / ``USER_PASSWORD`` / ``COMPANY_CODE``，再执行::

    python -m api.hy_oms.request.simulate_login --account A002 --password 123456
    python -m api.hy_oms.request.simulate_login --mode getSsoToken --company-code A001
    python -m api.hy_oms.request.simulate_login --dry-run
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path
from typing import Any, Optional
from urllib.parse import unquote


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


def build_log_on_params(
    *,
    user_account: str,
    user_password: str,
    validate: bool = True,
    **extra: Any,
) -> dict[str, Any]:
    """组装 ``logOn`` 的 paramsJson（不发请求）。"""
    account = str(user_account or "").strip()
    password = str(user_password or "").strip()
    params: dict[str, Any] = {
        "user_account": account,
        "user_password": password,
        **extra,
    }
    if validate and (not account or not password):
        raise ValueError("user_account / user_password 必填")
    return params


def build_sso_token_params(
    *,
    company_code: str,
    validate: bool = True,
    **extra: Any,
) -> dict[str, Any]:
    """组装 ``getSsoToken`` 的 paramsJson（不发请求）。"""
    code = str(company_code or "").strip()
    params: dict[str, Any] = {"company_code": code, **extra}
    if validate and not code:
        raise ValueError("company_code（客户代码）必填")
    return params


def _enrich_login_result(client: Any, mode: str, result: dict[str, Any]) -> dict[str, Any]:
    """在原始响应上附加 ``login_url``（及 SSO 的 user_code / token）。"""
    out = dict(result)
    if mode == "logOn":
        raw = result.get("data")
        out["login_url"] = unquote(str(raw)) if raw is not None else ""
        return out

    data = result.get("data") or {}
    if isinstance(data, dict) and "userCode" not in data and isinstance(data.get("data"), dict):
        data = data["data"]
    user_code = ""
    token = ""
    if isinstance(data, dict):
        user_code = str(data.get("userCode") or data.get("user_code") or "")
        token = str(data.get("token") or "")
    out["user_code"] = user_code
    out["token"] = token
    out["login_url"] = (
        client.build_quick_login_url(user_code=user_code, token=token)
        if user_code and token
        else ""
    )
    return out


def log_on(
    *,
    user_account: Optional[str] = None,
    user_password: Optional[str] = None,
    **extra: Any,
) -> dict[str, Any]:
    """调用 ``logOn``，返回完整响应（含解码后的 ``login_url``）。"""
    from api.hy_oms import HyOmsClient

    return HyOmsClient.from_config().simulate_login(
        mode="logOn",
        user_account=user_account,
        user_password=user_password,
        **extra,
    )


def get_sso_token(
    *,
    company_code: Optional[str] = None,
    **extra: Any,
) -> dict[str, Any]:
    """调用 ``getSsoToken``，返回完整响应（含 ``login_url`` / ``user_code`` / ``token``）。"""
    from api.hy_oms import HyOmsClient

    return HyOmsClient.from_config().simulate_login(
        mode="getSsoToken",
        company_code=company_code,
        **extra,
    )


def simulate_login(
    *,
    mode: str = "logOn",
    user_account: Optional[str] = None,
    user_password: Optional[str] = None,
    company_code: Optional[str] = None,
    **extra: Any,
) -> dict[str, Any]:
    """模拟登录统一入口，见 ``HyOmsClient.simulate_login``。"""
    from api.hy_oms import HyOmsClient

    return HyOmsClient.from_config().simulate_login(
        mode=mode,
        user_account=user_account,
        user_password=user_password,
        company_code=company_code,
        **extra,
    )


def main(argv=None) -> int:
    _bootstrap()
    from api.hy_oms import HyOmsClient
    from api.hy_oms.config import HyOmsConfig
    from api.hy_oms.exceptions import HyOmsError

    parser = argparse.ArgumentParser(
        description="鸿羽 OMS 模拟登录 logOn / getSsoToken"
    )
    parser.add_argument(
        "--mode",
        default="logOn",
        choices=["logOn", "getSsoToken"],
        help="logOn=账号密码登录；getSsoToken=按客户代码取 token（默认 logOn）",
    )
    parser.add_argument(
        "--account",
        "--user-account",
        dest="user_account",
        default=None,
        help="OMS 登录账号（logOn；缺省用 config.USER_ACCOUNT）",
    )
    parser.add_argument(
        "--password",
        "--user-password",
        dest="user_password",
        default=None,
        help="OMS 登录密码（logOn；缺省用 config.USER_PASSWORD）",
    )
    parser.add_argument(
        "--company-code",
        dest="company_code",
        default=None,
        help="客户代码（getSsoToken；缺省用 config.COMPANY_CODE）",
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
        cfg = HyOmsConfig.default()
        mode = args.mode
        service = "logOn" if mode == "logOn" else "getSsoToken"

        if args.body:
            params = _load_body(args.body)
            if mode == "logOn":
                if not str(params.get("user_account") or "").strip() or not str(
                    params.get("user_password") or ""
                ).strip():
                    raise ValueError("--body 中 user_account / user_password 必填")
            elif not str(params.get("company_code") or "").strip():
                raise ValueError("--body 中 company_code 必填")
        elif mode == "logOn":
            account = args.user_account if args.user_account is not None else cfg.user_account
            password = (
                args.user_password if args.user_password is not None else cfg.user_password
            )
            params = build_log_on_params(
                user_account=str(account or ""),
                user_password=str(password or ""),
            )
        else:
            code = args.company_code if args.company_code is not None else cfg.company_code
            params = build_sso_token_params(company_code=str(code or ""))

        if args.dry_run:
            preview = dict(params)
            if preview.get("user_password"):
                preview["user_password"] = "***"
            print(f"[DRY-RUN] {service} paramsJson:")
            print(json.dumps(preview, ensure_ascii=False, indent=2))
            return 0

        client = HyOmsClient.from_config()
        raw_result = client.call(service, params)
        result = _enrich_login_result(client, mode, raw_result)
    except (ValueError, json.JSONDecodeError) as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 2
    except HyOmsError as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        raw = getattr(exc, "raw", None)
        if raw is not None:
            print(json.dumps(raw, ensure_ascii=False, indent=2)[:4000], file=sys.stderr)
        return 1

    login_url = result.get("login_url") or ""
    print(f"[OK] service={service} ask={result.get('ask')} login_url={login_url}")
    if args.raw:
        print(json.dumps(result, ensure_ascii=False, indent=2)[:8000])
    elif login_url:
        print(login_url)
    else:
        print(json.dumps(result.get("data"), ensure_ascii=False, indent=2)[:8000])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
