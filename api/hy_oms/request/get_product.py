"""鸿羽 OMS 获取产品列表 ``getProductList``。

文档：http://oms.gindalogistik.com/api-doc/index.php（产品模块 → 获取产品列表）

必填：``page`` / ``pageSize``。

可选筛选：``product_sku`` / ``product_sku_like`` / ``product_sku_arr``，
以及创建/更新时间区间。

先在 ``api/hy_oms/config.py`` 填写凭证，再执行::

    python -m api.hy_oms.request.get_product
    python -m api.hy_oms.request.get_product --page 1 --page-size 10
    python -m api.hy_oms.request.get_product --product-sku SKU001 --raw
    python -m api.hy_oms.request.get_product --product-sku-arr SKU1,SKU2
    python -m api.hy_oms.request.get_product --all-pages --page-size 100
    python -m api.hy_oms.request.get_product --body path/to/body.json --dry-run
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path
from typing import Any, Mapping, Sequence, Union


# 产品状态（data[].product_status）
PRODUCT_STATUS: dict[str, str] = {
    "X": "废弃",
    "D": "草稿",
    "S": "可用",
    "P": "审核中",
    "R": "审核不通过",
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


def _parse_sku_arr(raw: str | None) -> list[str] | None:
    """``SKU1,SKU2`` 或 ``SKU1;SKU2`` → list。"""
    if raw is None or not str(raw).strip():
        return None
    text = str(raw).strip()
    sep = ";" if ";" in text else ","
    parts = [x.strip() for x in text.split(sep) if x.strip()]
    return parts or None


def build_params(
    *,
    page: int = 1,
    page_size: int = 10,
    product_sku: str | None = None,
    product_sku_like: str | None = None,
    product_sku_arr: Sequence[str] | None = None,
    start_time: str | None = None,
    end_time: str | None = None,
    update_start_time: str | None = None,
    update_end_time: str | None = None,
    validate: bool = True,
    **extra: Any,
) -> dict[str, Any]:
    """组装 ``getProductList`` 的 paramsJson（不发请求）。"""
    page_n = int(page)
    size_n = int(page_size)
    if validate:
        if page_n < 1:
            raise ValueError("page 必须 >= 1")
        if size_n < 1:
            raise ValueError("pageSize 必须 >= 1")

    params: dict[str, Any] = {
        "page": page_n,
        "pageSize": size_n,
        **extra,
    }
    optional: dict[str, Any] = {
        "product_sku": (str(product_sku).strip() if product_sku is not None else None) or None,
        "product_sku_like": (
            str(product_sku_like).strip() if product_sku_like is not None else None
        )
        or None,
        "product_sku_arr": list(product_sku_arr) if product_sku_arr else None,
        "start_time": start_time,
        "end_time": end_time,
        "update_start_time": update_start_time,
        "update_end_time": update_end_time,
    }
    for key, value in optional.items():
        if value is not None and value != "":
            params[key] = value
    return params


def get_product_list(
    *,
    page: int = 1,
    page_size: int = 10,
    product_sku: str | None = None,
    product_sku_like: str | None = None,
    product_sku_arr: Sequence[str] | None = None,
    start_time: str | None = None,
    end_time: str | None = None,
    update_start_time: str | None = None,
    update_end_time: str | None = None,
    all_pages: bool = False,
    **extra: Any,
) -> Union[dict[str, Any], list[Any]]:
    """调用 ``getProductList``。

    ``all_pages=False``（默认）返回完整响应 dict（含 ask / data / count）。
    ``all_pages=True`` 时翻页拉取，返回 ``data`` 行组成的 list。
    """
    from api.hy_oms import HyOmsClient

    client = HyOmsClient.from_config()
    params = build_params(
        page=page,
        page_size=page_size,
        product_sku=product_sku,
        product_sku_like=product_sku_like,
        product_sku_arr=product_sku_arr,
        start_time=start_time,
        end_time=end_time,
        update_start_time=update_start_time,
        update_end_time=update_end_time,
        **extra,
    )
    if all_pages:
        base = {k: v for k, v in params.items() if k not in {"page", "pageSize"}}
        return client.fetch_all("getProductList", base, page_size=page_size)
    return client.call("getProductList", params)


def summarize_product(row: Any) -> str:
    """从单条产品提炼摘要，便于 CLI 一行展示。"""
    if not isinstance(row, Mapping):
        return str(row)
    sku = row.get("product_sku") or ""
    title = row.get("product_title") or row.get("product_title_en") or ""
    status = str(row.get("product_status") or "").strip().upper()
    status_label = PRODUCT_STATUS.get(status, status or "?")
    barcode = row.get("goods_barcode") or ""
    weight = row.get("product_weight") or ""
    return (
        f"sku={sku} status={status}({status_label}) "
        f"title={title!r} barcode={barcode} weight={weight}"
    )


def main(argv=None) -> int:
    _bootstrap()
    from api.hy_oms.exceptions import HyOmsError

    parser = argparse.ArgumentParser(
        description="鸿羽 OMS 获取产品列表 getProductList"
    )
    parser.add_argument("--page", type=int, default=1, help="页码，默认 1")
    parser.add_argument(
        "--page-size",
        type=int,
        default=10,
        dest="page_size",
        help="每页条数，默认 10",
    )
    parser.add_argument(
        "--product-sku",
        dest="product_sku",
        default=None,
        help="SKU 精确匹配",
    )
    parser.add_argument(
        "--product-sku-like",
        dest="product_sku_like",
        default=None,
        help="SKU 模糊搜索",
    )
    parser.add_argument(
        "--product-sku-arr",
        dest="product_sku_arr",
        default=None,
        help="多个 SKU，逗号或分号分隔",
    )
    parser.add_argument(
        "--start-time",
        dest="start_time",
        default=None,
        help="产品创建起始时间 YYYY-MM-DD HH:MM:SS",
    )
    parser.add_argument(
        "--end-time",
        dest="end_time",
        default=None,
        help="产品创建结束时间",
    )
    parser.add_argument(
        "--update-start-time",
        dest="update_start_time",
        default=None,
        help="产品更新起始时间",
    )
    parser.add_argument(
        "--update-end-time",
        dest="update_end_time",
        default=None,
        help="产品更新结束时间",
    )
    parser.add_argument(
        "--all-pages",
        action="store_true",
        help="自动翻页拉取全部，输出 data 列表",
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
            if "page" not in params:
                params["page"] = args.page
            if "pageSize" not in params:
                params["pageSize"] = args.page_size
        else:
            params = build_params(
                page=args.page,
                page_size=args.page_size,
                product_sku=args.product_sku,
                product_sku_like=args.product_sku_like,
                product_sku_arr=_parse_sku_arr(args.product_sku_arr),
                start_time=args.start_time,
                end_time=args.end_time,
                update_start_time=args.update_start_time,
                update_end_time=args.update_end_time,
            )

        if args.dry_run:
            print("[DRY-RUN] getProductList paramsJson:")
            print(json.dumps(params, ensure_ascii=False, indent=2))
            return 0

        from api.hy_oms import HyOmsClient

        client = HyOmsClient.from_config()
        if args.all_pages:
            page_size = int(params.get("pageSize") or args.page_size)
            base = {k: v for k, v in params.items() if k not in {"page", "pageSize"}}
            rows = client.fetch_all("getProductList", base, page_size=page_size)
            print(f"[OK] service=getProductList all_pages=True count={len(rows)}")
            if rows and not args.raw:
                print(summarize_product(rows[0]))
            print(json.dumps(rows if args.raw else rows[:20], ensure_ascii=False, indent=2)[:8000])
            if not args.raw and len(rows) > 20:
                print(f"... 共 {len(rows)} 条，仅展示前 20；加 --raw 输出全部")
            return 0

        result = client.call("getProductList", params)
    except (ValueError, json.JSONDecodeError) as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 2
    except HyOmsError as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        raw = getattr(exc, "raw", None)
        if raw is not None:
            print(json.dumps(raw, ensure_ascii=False, indent=2)[:4000], file=sys.stderr)
        return 1

    data = result.get("data") or []
    if isinstance(data, dict):
        data = [data]
    if not isinstance(data, list):
        data = []
    count = result.get("count", len(data))
    next_page = result.get("nextPage")
    print(
        f"[OK] service=getProductList ask={result.get('ask')} "
        f"page={params.get('page')} pageSize={params.get('pageSize')} "
        f"returned={len(data)} count={count} nextPage={next_page}"
    )
    if data:
        print(summarize_product(data[0]))
    payload = result if args.raw else data
    print(json.dumps(payload, ensure_ascii=False, indent=2)[:8000])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
