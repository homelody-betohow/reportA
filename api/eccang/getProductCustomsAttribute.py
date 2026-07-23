"""易仓 获取产品报关属性（getProductCustomsAttribute）。

文档：https://open.eccang.com/#/documentCenter?docId=111799&catId=0-187-187,0-177

该接口返回报关属性字典（电池类型、磁性物质、木质类等），一般无需业务筛选参数。

运行（在项目根目录）::

    python -m api.eccang.getProductCustomsAttribute
    python -m api.eccang.getProductCustomsAttribute --body "{}"
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


def get_product_customs_attribute(
    *,
    extra: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """调用 ``getProductCustomsAttribute``，返回解析后的完整响应。"""
    from api.eccang import EccangService

    client = EccangService()
    return client.get_product_customs_attribute(extra=extra)


def main(argv: list[str] | None = None) -> int:
    _bootstrap()
    from api.eccang.exceptions import EccangApiError, EccangConfigError

    parser = argparse.ArgumentParser(
        description="易仓 获取产品报关属性 getProductCustomsAttribute（docId=111799）",
    )
    parser.add_argument(
        "--body",
        help="可选 biz_content JSON（默认空对象 {}）",
    )
    args = parser.parse_args(argv)

    extra: dict[str, Any] | None = None
    if args.body:
        try:
            parsed = json.loads(args.body)
        except json.JSONDecodeError as exc:
            print(f"[FAIL] --body JSON 格式错误：{exc}", file=sys.stderr)
            return 2
        if not isinstance(parsed, dict):
            print("[FAIL] --body 必须是 JSON 对象", file=sys.stderr)
            return 2
        extra = parsed

    try:
        response = get_product_customs_attribute(extra=extra)
    except EccangConfigError as exc:
        print(f"[FAIL] 配置错误：{exc}", file=sys.stderr)
        return 2
    except EccangApiError as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        if exc.code:
            print(f"  code={exc.code}", file=sys.stderr)
        if exc.raw_payload is not None:
            print(json.dumps(exc.raw_payload, ensure_ascii=False, indent=2))
        return 1

    printable = {k: v for k, v in response.items() if k != "biz_content"}
    data = printable.get("data") or {}
    attrs = data.get("data") if isinstance(data, dict) else None
    count = len(attrs) if isinstance(attrs, dict) else 0
    print(
        f"[OK] method=getProductCustomsAttribute version=V1.0.0 "
        f"attribute_count={count}"
    )
    print(json.dumps(printable, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
