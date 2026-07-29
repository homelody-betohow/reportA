"""鸿羽 OMS 创建产品 ``createProduct``。

文档：http://oms.gindalogistik.com/api-doc/index.php（产品模块 → 创建产品）

必填：``product_sku`` / ``product_title`` / ``product_weight`` /
``product_length`` / ``product_width`` / ``product_height`` /
``product_declared_value`` / ``product_declared_name``。

``verify``：不传或 0 → 草稿；1 → 正式产品（审核通过后不可编辑）。

先在 ``api/hy_oms/config.py`` 填写凭证，再执行::

    python -m api.hy_oms.request.create_product ^
      --product-sku TEST001 --product-title "测试产品" ^
      --weight 0.5 --length 10 --width 8 --height 5 ^
      --declared-value 10 --declared-name "Test Product"

    python -m api.hy_oms.request.create_product --body path/to/body.json --dry-run
    python -m api.hy_oms.request.create_product --body path/to/body.json --raw
"""

from __future__ import annotations

import argparse
import base64
import importlib.util
import json
import mimetypes
import sys
from pathlib import Path
from typing import Any, Mapping, Optional, Union


REQUIRED_FIELDS = (
    "product_sku",
    "product_title",
    "product_weight",
    "product_length",
    "product_width",
    "product_height",
    "product_declared_value",
    "product_declared_name",
)

BATTERY_TYPES = {
    "PI970 DHL no more than 2 batteries",
    "PI970 DHL more than 2 batteries",
    "PI967 more than 2 batteries or 4 cells",
    "PI966",
    "PI967",
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


def _set_optional(params: dict[str, Any], mapping: Mapping[str, Any]) -> None:
    for key, value in mapping.items():
        if value is not None:
            params[key] = value


def _as_float(value: Any, name: str) -> float:
    try:
        return float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{name} 必须是数字，收到: {value!r}") from exc


def _as_int(value: Any, name: str) -> int:
    try:
        return int(float(value))
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{name} 必须是整数，收到: {value!r}") from exc


def load_customer_img(
    path: Union[str, Path],
    *,
    file_type: str = "img",
) -> dict[str, Any]:
    """从本地图片文件生成 ``customerImg``（base64 data URL）。"""
    p = Path(path)
    if not p.is_file():
        raise ValueError(f"图片文件不存在: {p}")
    raw = p.read_bytes()
    mime, _ = mimetypes.guess_type(str(p))
    if not mime or not mime.startswith("image/"):
        mime = "image/png"
    b64 = base64.b64encode(raw).decode("ascii")
    return {
        "file_type": file_type,
        "base64_img": f"data:{mime};base64,{b64}",
    }


def validate_params(params: Mapping[str, Any]) -> None:
    """校验 createProduct 必填与枚举字段。"""
    missing = [k for k in REQUIRED_FIELDS if params.get(k) in (None, "")]
    if missing:
        raise ValueError(f"createProduct 缺少必填: {', '.join(missing)}")

    sku = str(params["product_sku"]).strip()
    if not sku:
        raise ValueError("product_sku 不能为空")

    for name in (
        "product_weight",
        "product_length",
        "product_width",
        "product_height",
        "product_declared_value",
    ):
        _as_float(params[name], name)

    battery = params.get("battery_type")
    if battery is not None and str(battery).strip() and str(battery) not in BATTERY_TYPES:
        raise ValueError(
            f"battery_type 非法: {battery!r}，允许值: {sorted(BATTERY_TYPES)}"
        )

    contain = params.get("contain_battery")
    if contain is not None and _as_int(contain, "contain_battery") not in (0, 1):
        raise ValueError("contain_battery 须为 0 或 1")

    verify = params.get("verify")
    if verify is not None and _as_int(verify, "verify") not in (0, 1):
        raise ValueError("verify 须为 0（草稿）或 1（正式）")

    img = params.get("customerImg")
    if img is not None:
        if not isinstance(img, Mapping):
            raise ValueError("customerImg 必须是对象")
        if str(img.get("file_type") or "").strip() != "img":
            raise ValueError('customerImg.file_type 只能为 "img"')
        if not img.get("base64_img"):
            raise ValueError("customerImg.base64_img 必填")


def build_params(
    *,
    product_sku: str,
    product_title: str,
    product_weight: float,
    product_length: float,
    product_width: float,
    product_height: float,
    product_declared_value: float,
    product_declared_name: str,
    reference_no: Optional[str] = None,
    product_title_en: Optional[str] = None,
    product_net_weight: Optional[float] = None,
    contain_battery: Optional[int] = None,
    battery_type: Optional[str] = None,
    product_declared_name_zh: Optional[str] = None,
    hs_code: Optional[str] = None,
    cat_lang: Optional[str] = None,
    warning_qty: Optional[int] = None,
    warning_days: Optional[int] = None,
    product_brand: Optional[str] = None,
    product_model: Optional[str] = None,
    product_origin: Optional[str] = None,
    product_material: Optional[str] = None,
    product_use_en: Optional[str] = None,
    product_material_en: Optional[str] = None,
    product_desc_url: Optional[str] = None,
    cat_id_level0: Optional[int] = None,
    cat_id_level1: Optional[int] = None,
    cat_id_level2: Optional[int] = None,
    verify: Optional[int] = None,
    customer_img: Optional[Mapping[str, Any]] = None,
    product_color: Optional[str] = None,
    shared_product: Optional[int] = None,
    shared_unit_price: Optional[Union[str, float]] = None,
    product_description: Optional[str] = None,
    is_box_more_sku: Optional[int] = None,
    fragile_property: Optional[int] = None,
    product_size_type: Optional[str] = None,
    is_batch_tag: Optional[int] = None,
    ean: Optional[str] = None,
    ncm: Optional[str] = None,
    cest: Optional[str] = None,
    sku_sort_code: Optional[str] = None,
    is_serialized: Optional[int] = None,
    validate: bool = True,
    **extra: Any,
) -> dict[str, Any]:
    """组装 ``createProduct`` 的 paramsJson（不发请求）。"""
    params: dict[str, Any] = {
        "product_sku": str(product_sku or "").strip(),
        "product_title": str(product_title or "").strip(),
        "product_weight": _as_float(product_weight, "product_weight"),
        "product_length": _as_float(product_length, "product_length"),
        "product_width": _as_float(product_width, "product_width"),
        "product_height": _as_float(product_height, "product_height"),
        "product_declared_value": _as_float(
            product_declared_value, "product_declared_value"
        ),
        "product_declared_name": str(product_declared_name or "").strip(),
        **extra,
    }
    _set_optional(
        params,
        {
            "reference_no": reference_no,
            "product_title_en": product_title_en,
            "product_net_weight": product_net_weight,
            "contain_battery": contain_battery,
            "battery_type": battery_type,
            "product_declared_name_zh": product_declared_name_zh,
            "hs_code": hs_code,
            "cat_lang": cat_lang,
            "warning_qty": warning_qty,
            "warning_days": warning_days,
            "product_brand": product_brand,
            "product_model": product_model,
            "product_origin": product_origin,
            "product_material": product_material,
            "product_use_en": product_use_en,
            "product_material_en": product_material_en,
            "product_desc_url": product_desc_url,
            "cat_id_level0": cat_id_level0,
            "cat_id_level1": cat_id_level1,
            "cat_id_level2": cat_id_level2,
            "verify": verify,
            "customerImg": dict(customer_img) if customer_img is not None else None,
            "product_color": product_color,
            "shared_product": shared_product,
            "shared_unit_price": shared_unit_price,
            "product_description": product_description,
            "is_box_more_sku": is_box_more_sku,
            "fragile_property": fragile_property,
            "product_size_type": product_size_type,
            "is_batch_tag": is_batch_tag,
            "EAN": ean,
            "NCM": ncm,
            "CEST": cest,
            "sku_sort_code": sku_sort_code,
            "is_serialized": is_serialized,
        },
    )
    if validate:
        validate_params(params)
    return params


def create_product(
    *,
    product_sku: str,
    product_title: str,
    product_weight: float,
    product_length: float,
    product_width: float,
    product_height: float,
    product_declared_value: float,
    product_declared_name: str,
    **kwargs: Any,
) -> dict[str, Any]:
    """调用 ``createProduct``，返回完整响应 dict（含 ask / product_sku）。"""
    from api.hy_oms import HyOmsClient

    client = HyOmsClient.from_config()
    return client.create_product(
        product_sku=product_sku,
        product_title=product_title,
        product_weight=product_weight,
        product_length=product_length,
        product_width=product_width,
        product_height=product_height,
        product_declared_value=product_declared_value,
        product_declared_name=product_declared_name,
        **kwargs,
    )


def main(argv=None) -> int:
    _bootstrap()
    from api.hy_oms.exceptions import HyOmsError

    parser = argparse.ArgumentParser(
        description="鸿羽 OMS 创建产品 createProduct"
    )
    parser.add_argument("--product-sku", dest="product_sku", default=None, help="SKU（必填）")
    parser.add_argument(
        "--product-title", dest="product_title", default=None, help="产品标题（必填）"
    )
    parser.add_argument(
        "--product-title-en", dest="product_title_en", default=None, help="英文标题"
    )
    parser.add_argument("--reference-no", dest="reference_no", default=None, help="自定义编码")
    parser.add_argument("--weight", dest="product_weight", type=float, default=None, help="重量 KG")
    parser.add_argument(
        "--net-weight", dest="product_net_weight", type=float, default=None, help="净重 KG"
    )
    parser.add_argument("--length", dest="product_length", type=float, default=None, help="长 CM")
    parser.add_argument("--width", dest="product_width", type=float, default=None, help="宽 CM")
    parser.add_argument("--height", dest="product_height", type=float, default=None, help="高 CM")
    parser.add_argument(
        "--declared-value",
        dest="product_declared_value",
        type=float,
        default=None,
        help="申报价值 USD",
    )
    parser.add_argument(
        "--declared-name",
        dest="product_declared_name",
        default=None,
        help="申报名称英文（必填）",
    )
    parser.add_argument(
        "--declared-name-zh",
        dest="product_declared_name_zh",
        default=None,
        help="申报名称中文",
    )
    parser.add_argument("--hs-code", dest="hs_code", default=None)
    parser.add_argument("--cat-lang", dest="cat_lang", default=None, help="zh / en")
    parser.add_argument(
        "--contain-battery",
        dest="contain_battery",
        type=int,
        choices=(0, 1),
        default=None,
    )
    parser.add_argument("--battery-type", dest="battery_type", default=None)
    parser.add_argument(
        "--verify",
        type=int,
        choices=(0, 1),
        default=None,
        help="0草稿 / 1正式产品",
    )
    parser.add_argument("--brand", dest="product_brand", default=None)
    parser.add_argument("--model", dest="product_model", default=None)
    parser.add_argument("--origin", dest="product_origin", default=None)
    parser.add_argument("--material", dest="product_material", default=None)
    parser.add_argument("--color", dest="product_color", default=None)
    parser.add_argument("--description", dest="product_description", default=None)
    parser.add_argument("--ean", dest="ean", default=None)
    parser.add_argument(
        "--fragile",
        dest="fragile_property",
        type=int,
        choices=(0, 1),
        default=None,
        help="0无易碎 / 1易碎",
    )
    parser.add_argument(
        "--image",
        dest="image_path",
        default=None,
        help="本地图片路径，自动转 customerImg base64",
    )
    parser.add_argument(
        "--customer-img",
        dest="customer_img",
        default=None,
        help="customerImg JSON 字符串或文件路径",
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
            validate_params(params)
        else:
            required_cli = {
                "--product-sku": args.product_sku,
                "--product-title": args.product_title,
                "--weight": args.product_weight,
                "--length": args.product_length,
                "--width": args.product_width,
                "--height": args.product_height,
                "--declared-value": args.product_declared_value,
                "--declared-name": args.product_declared_name,
            }
            missing = [k for k, v in required_cli.items() if v is None or v == ""]
            if missing:
                print(f"缺少必填参数: {', '.join(missing)}（或改用 --body）", file=sys.stderr)
                return 2

            customer_img = None
            if args.customer_img:
                customer_img = _load_body(args.customer_img)
            elif args.image_path:
                customer_img = load_customer_img(args.image_path)

            params = build_params(
                product_sku=args.product_sku,
                product_title=args.product_title,
                product_weight=args.product_weight,
                product_length=args.product_length,
                product_width=args.product_width,
                product_height=args.product_height,
                product_declared_value=args.product_declared_value,
                product_declared_name=args.product_declared_name,
                reference_no=args.reference_no,
                product_title_en=args.product_title_en,
                product_net_weight=args.product_net_weight,
                contain_battery=args.contain_battery,
                battery_type=args.battery_type,
                product_declared_name_zh=args.product_declared_name_zh,
                hs_code=args.hs_code,
                cat_lang=args.cat_lang,
                verify=args.verify,
                customer_img=customer_img,
                product_brand=args.product_brand,
                product_model=args.product_model,
                product_origin=args.product_origin,
                product_material=args.product_material,
                product_color=args.product_color,
                product_description=args.product_description,
                fragile_property=args.fragile_property,
                ean=args.ean,
            )

        if args.dry_run:
            # 打印时截断超长 base64，避免刷屏
            printable = dict(params)
            img = printable.get("customerImg")
            if isinstance(img, dict) and isinstance(img.get("base64_img"), str):
                b64 = img["base64_img"]
                if len(b64) > 120:
                    printable["customerImg"] = {
                        **img,
                        "base64_img": b64[:80] + f"...({len(b64)} chars)",
                    }
            print("[DRY-RUN] createProduct paramsJson:")
            print(json.dumps(printable, ensure_ascii=False, indent=2))
            return 0

        from api.hy_oms import HyOmsClient

        result = HyOmsClient.from_config().call("createProduct", params)
    except (ValueError, json.JSONDecodeError) as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 2
    except HyOmsError as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        raw = getattr(exc, "raw", None)
        if raw is not None:
            print(json.dumps(raw, ensure_ascii=False, indent=2)[:4000], file=sys.stderr)
        return 1

    sku = result.get("product_sku") or params.get("product_sku")
    print(
        f"[OK] service=createProduct ask={result.get('ask')} "
        f"f_ask={result.get('f_ask')} product_sku={sku}"
    )
    payload = (
        result
        if args.raw
        else {
            "product_sku": sku,
            "message": result.get("message"),
            "f_ask": result.get("f_ask"),
        }
    )
    print(json.dumps(payload, ensure_ascii=False, indent=2)[:8000])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
