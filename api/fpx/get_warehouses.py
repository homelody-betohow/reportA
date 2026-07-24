"""4PX 查询仓库信息（com.basis.warehouse.getlist），写入 basis/warehouse_list.json。

文档：https://open.4px.com/v2/doc/detail?ids=54,76,153

中文名依赖客户端 ``Accept: application/json;charset=utf-8``（见 ``api/fpx/client.py``）；
否则服务端可能按 ISO-8859-1 回包，把中文变成 ``?``。

先在 ``api/fpx/config.py`` 填写凭证，再执行::

    python -m api.fpx.get_warehouses
    python -m api.fpx.get_warehouses --country FR
    python -m api.fpx.get_warehouses --service-code F --country DE
    python -m api.fpx.get_warehouses --raw
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

LIST_PATH = Path(__file__).resolve().parent / "basis" / "warehouse_list.json"

_SERVICE_CODES = {
    "F": "订单履约",
    "S": "自发服务",
    "T": "转运服务",
    "R": "退件服务",
}


def _bootstrap() -> None:
    root = Path(__file__).resolve().parents[2]
    epr = root / "ensure_project_root.py"
    spec = importlib.util.spec_from_file_location("ensure_project_root", epr)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    mod.bootstrap(__file__)


def _normalize_rows(data: Any) -> list[dict[str, str]]:
    if not isinstance(data, list):
        return []
    out: list[dict[str, str]] = []
    seen: set[str] = set()
    for row in data:
        if not isinstance(row, dict):
            continue
        code = str(row.get("warehouse_code") or "").strip()
        if not code or code in seen:
            continue
        seen.add(code)
        out.append(
            {
                "warehouse_code": code,
                "warehouse_name_cn": str(row.get("warehouse_name_cn") or "").strip(),
                "warehouse_name_en": str(row.get("warehouse_name_en") or "").strip(),
                "country": str(row.get("country") or "").strip().upper(),
                "service_code": str(row.get("service_code") or "").strip(),
            }
        )
    out.sort(key=lambda r: (r.get("country") or "", r["warehouse_code"]))
    return out


def _build_name_map(warehouses: list[dict[str, str]]) -> dict[str, str]:
    """中文名 / 英文名 → warehouse_code（同名保留先出现的）。"""
    mapping: dict[str, str] = {}
    for wh in warehouses:
        code = wh.get("warehouse_code") or ""
        if not code:
            continue
        for key in (wh.get("warehouse_name_cn") or "", wh.get("warehouse_name_en") or ""):
            name = key.strip()
            if name and name not in mapping:
                mapping[name] = code
    return mapping


def _dump_payload(payload: dict[str, Any], path: Path) -> None:
    """写出 JSON：warehouses 每条一行，map 按键排序。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    meta_keys = (
        "version",
        "description",
        "method",
        "service_code",
        "country",
        "updated_at",
    )
    lines = ["{"]

    for key in meta_keys:
        if key not in payload:
            continue
        lines.append(
            f"  {json.dumps(key, ensure_ascii=False)}: "
            f"{json.dumps(payload[key], ensure_ascii=False)},"
        )

    warehouses = (
        payload.get("warehouses") if isinstance(payload.get("warehouses"), list) else []
    )
    lines.append('  "warehouses": [')
    for i, row in enumerate(warehouses):
        suffix = "," if i < len(warehouses) - 1 else ""
        lines.append(
            f"    {json.dumps(row, ensure_ascii=False, separators=(', ', ': '))}{suffix}"
        )
    lines.append("  ],")

    raw = payload.get("map") if isinstance(payload.get("map"), dict) else {}
    ordered = {k: raw[k] for k in sorted(raw.keys(), key=str)}
    lines.append('  "map": {')
    map_items = list(ordered.items())
    for i, (k, v) in enumerate(map_items):
        suffix = "," if i < len(map_items) - 1 else ""
        lines.append(
            f"    {json.dumps(k, ensure_ascii=False)}: "
            f"{json.dumps(v, ensure_ascii=False)}{suffix}"
        )
    lines.append("  }")
    lines.append("}")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main(argv=None) -> int:
    _bootstrap()
    from api.fpx import FpxClient
    from api.fpx.config import FpxConfig
    from api.fpx.exceptions import FpxError

    parser = argparse.ArgumentParser(
        description="4PX 查询仓库信息 com.basis.warehouse.getlist（ids=54,76,153）"
    )
    parser.add_argument(
        "--service-code",
        dest="service_code",
        default=None,
        choices=sorted(_SERVICE_CODES),
        help="业务类型：F 订单履约 / S 自发 / T 转运 / R 退件（默认不过滤）",
    )
    parser.add_argument(
        "--country",
        default=None,
        help="国家二字码过滤，如 FR/DE/CN（默认不过滤）",
    )
    parser.add_argument(
        "--out",
        default=str(LIST_PATH),
        help=f"输出 JSON 路径（默认 {LIST_PATH}）",
    )
    parser.add_argument("--raw", action="store_true", help="打印完整响应 JSON")
    parser.add_argument("--sandbox", action="store_true", help="使用沙箱域名")
    parser.add_argument(
        "--no-write",
        action="store_true",
        help="只拉取打印，不写 warehouse_list.json",
    )
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
        result = client.get_warehouse_list(
            service_code=args.service_code,
            country=args.country,
        )
    except FpxError as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 1

    warehouses = _normalize_rows(result.get("data"))
    name_map = _build_name_map(warehouses)
    sc_label = (
        f"{args.service_code}({_SERVICE_CODES[args.service_code]})"
        if args.service_code
        else "-"
    )
    print(
        f"[OK] method=com.basis.warehouse.getlist "
        f"service_code={sc_label} country={args.country or '-'} "
        f"result={result.get('result')} msg={result.get('msg')} "
        f"warehouses={len(warehouses)}"
    )

    if args.raw:
        print(json.dumps(result, ensure_ascii=False, indent=2)[:8000])
    else:
        preview = warehouses[:15]
        if preview:
            print(json.dumps(preview, ensure_ascii=False, indent=2))
            if len(warehouses) > len(preview):
                print(f"... 共 {len(warehouses)} 条")

    if args.no_write:
        return 0

    out_path = Path(args.out)
    payload = {
        "version": 1,
        "description": "4PX 开放平台仓库列表（com.basis.warehouse.getlist）；"
        "map=中文名/英文名→warehouse_code。",
        "method": "com.basis.warehouse.getlist",
        "service_code": args.service_code or "",
        "country": (args.country or "").strip().upper(),
        "updated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "warehouses": warehouses,
        "map": name_map,
    }
    _dump_payload(payload, out_path)
    print(f"[OK] wrote {out_path} ({len(warehouses)} warehouses, map={len(name_map)})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
