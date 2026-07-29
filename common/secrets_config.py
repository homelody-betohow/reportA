"""从 ``config/secrets.json`` / 环境变量加载 API 凭证（不入库）。

优先级：环境变量 > secrets.json > 调用方传入的默认值。

查找路径（发布布局优先 ``dist/config``）::

    <dist>/config/secrets.json
    <exe 同级>/config/secrets.json   # 兼容旧布局
    <cwd>/config/secrets.json
"""
from __future__ import annotations

import json
import os
from functools import lru_cache
from pathlib import Path
from typing import Any, Mapping, Optional

from common.dist_paths import resolve_config_file

_SECRETS_NAME = "secrets.json"


def resolve_secrets_path() -> Optional[Path]:
    return resolve_config_file(_SECRETS_NAME)


@lru_cache(maxsize=1)
def load_secrets() -> Mapping[str, Any]:
    path = resolve_secrets_path()
    if path is None:
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError(f"secrets.json 须为 JSON 对象: {path}")
    return data


def secrets_section(name: str) -> Mapping[str, Any]:
    raw = load_secrets().get(name)
    return raw if isinstance(raw, dict) else {}


def _pick(
    *candidates: Any,
    default: str = "",
) -> str:
    for c in candidates:
        if c is None:
            continue
        text = str(c).strip()
        if text:
            return text
    return default


def ding_disk_secrets(*, defaults: Optional[Mapping[str, Any]] = None) -> dict[str, Any]:
    """合并环境变量、secrets.json、代码默认值 → 钉钉配置字段。"""
    d = defaults or {}
    s = secrets_section("ding_disk")
    return {
        "app_key": _pick(os.getenv("DING_DISK_APP_KEY"), s.get("app_key"), d.get("app_key")),
        "app_secret": _pick(
            os.getenv("DING_DISK_APP_SECRET"), s.get("app_secret"), d.get("app_secret")
        ),
        "operator_id": _pick(
            os.getenv("DING_DISK_OPERATOR_ID"), s.get("operator_id"), d.get("operator_id")
        ),
        "workbook_id": _pick(
            os.getenv("DING_DISK_WORKBOOK_ID"), s.get("workbook_id"), d.get("workbook_id")
        ),
        "api_host": _pick(
            os.getenv("DING_DISK_API_HOST"), s.get("api_host"), d.get("api_host"),
            default="https://api.dingtalk.com",
        ),
        "oapi_host": _pick(
            os.getenv("DING_DISK_OAPI_HOST"), s.get("oapi_host"), d.get("oapi_host"),
            default="https://oapi.dingtalk.com",
        ),
        "timeout": float(
            os.getenv("DING_DISK_TIMEOUT")
            or s.get("timeout")
            or d.get("timeout")
            or 60.0
        ),
        "token_refresh_skew": float(
            os.getenv("DING_DISK_TOKEN_REFRESH_SKEW")
            or s.get("token_refresh_skew")
            or d.get("token_refresh_skew")
            or 120.0
        ),
    }


def hy_oms_secrets(*, defaults: Optional[Mapping[str, Any]] = None) -> dict[str, Any]:
    """合并环境变量、secrets.json、代码默认值 → 鸿羽 OMS 配置字段。"""
    d = defaults or {}
    s = secrets_section("hy_oms")
    return {
        "app_token": _pick(
            os.getenv("HY_OMS_APP_TOKEN"), s.get("app_token"), d.get("app_token")
        ),
        "app_key": _pick(os.getenv("HY_OMS_APP_KEY"), s.get("app_key"), d.get("app_key")),
        "user_account": _pick(
            os.getenv("HY_OMS_USER_ACCOUNT"), s.get("user_account"), d.get("user_account")
        ),
        "user_password": _pick(
            os.getenv("HY_OMS_USER_PASSWORD"),
            s.get("user_password"),
            d.get("user_password"),
        ),
        "company_code": _pick(
            os.getenv("HY_OMS_COMPANY_CODE"), s.get("company_code"), d.get("company_code")
        ),
        "base_url": _pick(
            os.getenv("HY_OMS_BASE_URL"),
            s.get("base_url"),
            d.get("base_url"),
            default="http://oms.gindalogistik.com",
        ),
        "language": _pick(
            os.getenv("HY_OMS_LANGUAGE"),
            s.get("language"),
            d.get("language"),
            default="zh_CN",
        ),
        "timeout": float(
            os.getenv("HY_OMS_TIMEOUT") or s.get("timeout") or d.get("timeout") or 60.0
        ),
    }


def clear_secrets_cache() -> None:
    """测试用：清空 load_secrets 缓存。"""
    load_secrets.cache_clear()
