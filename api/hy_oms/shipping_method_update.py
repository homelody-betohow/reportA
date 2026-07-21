"""从鸿羽 OMS 拉取全部运输方式，保存到 ``shipping_method_map.json``。

不依赖订单 Excel。用法（项目根目录）::

    python -m api.hy_oms.shipping_method_update
    python -m api.hy_oms.shipping_method_update --warehouse-code DEHY
    python -m api.hy_oms.shipping_method_update --warehouse-code DE03
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


MAP_PATH = Path(__file__).resolve().parent / "shipping_method_map.json"


def _bootstrap() -> None:
    root = Path(__file__).resolve().parents[2]
    epr = root / "ensure_project_root.py"
    spec = importlib.util.spec_from_file_location("ensure_project_root", epr)
    mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(mod)
    mod.bootstrap(__file__)


def _load_existing(path: Path) -> dict[str, Any]:
    if not path.is_file():
        return {}
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return payload if isinstance(payload, dict) else {}


def _fetch_channels(warehouse_code: str) -> list[dict[str, str]]:
    from api.hy_oms import HyOmsClient

    client = HyOmsClient.from_env()
    resp = client.call("getShippingMethod", {"warehouseCode": warehouse_code})
    rows = resp.get("data") or []
    out: list[dict[str, str]] = []
    if not isinstance(rows, list):
        return out

    for row in rows:
        if not isinstance(row, dict):
            continue
        code = str(row.get("code") or "").strip()
        if not code:
            continue
        wh = str(row.get("warehouse_code") or "").strip() or warehouse_code
        out.append(
            {
                "code": code,
                "name": str(row.get("name") or "").strip(),
                "name_en": str(row.get("name_en") or "").strip(),
                "warehouse_code": wh,
            }
        )

    # 同一仓库下去重 code
    seen: set[tuple[str, str]] = set()
    uniq: list[dict[str, str]] = []
    for r in out:
        key = (r["warehouse_code"], r["code"])
        if key in seen:
            continue
        seen.add(key)
        uniq.append(r)
    uniq.sort(key=lambda r: (r["warehouse_code"], r["code"]))
    return uniq


def _build_name_map(channels: list[dict[str, str]]) -> dict[str, str]:
    """OMS 中文名 → code（同名保留先出现的）。"""
    mapping: dict[str, str] = {}
    for ch in channels:
        name = ch.get("name") or ""
        code = ch.get("code") or ""
        if name and code and name not in mapping:
            mapping[name] = code
    return mapping


def _dump_payload(payload: dict[str, Any], path: Path) -> None:
    """写出 JSON：channels 每条一行，map/aliases 按键排序。"""
    meta_keys = ("version", "description", "warehouse_code", "updated_at")
    lines = ["{"]

    for key in meta_keys:
        if key not in payload:
            continue
        lines.append(
            f"  {json.dumps(key, ensure_ascii=False)}: "
            f"{json.dumps(payload[key], ensure_ascii=False)},"
        )

    channels = payload.get("channels") if isinstance(payload.get("channels"), list) else []
    lines.append('  "channels": [')
    for i, row in enumerate(channels):
        suffix = "," if i < len(channels) - 1 else ""
        lines.append(
            f"    {json.dumps(row, ensure_ascii=False, separators=(', ', ': '))}{suffix}"
        )
    lines.append("  ],")

    for section in ("map", "aliases"):
        raw = payload.get(section) if isinstance(payload.get(section), dict) else {}
        ordered = {k: raw[k] for k in sorted(raw.keys(), key=str)}
        is_last = section == "aliases"
        lines.append(f'  {json.dumps(section, ensure_ascii=False)}: {{')
        items = list(ordered.items())
        for i, (k, v) in enumerate(items):
            suffix = "," if i < len(items) - 1 else ""
            lines.append(
                f"    {json.dumps(k, ensure_ascii=False)}: "
                f"{json.dumps(v, ensure_ascii=False)}{suffix}"
            )
        lines.append(f"  }}{'' if is_last else ','}")

    lines.extend(["}", ""])
    path.write_text("\n".join(lines), encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _bootstrap()
    parser = argparse.ArgumentParser(
        description="拉取鸿羽 OMS 运输方式并保存到 shipping_method_map.json"
    )
    parser.add_argument(
        "--warehouse-code",
        default="DEHY",
        help="getShippingMethod 仓库代码（默认 DEHY）",
    )
    parser.add_argument(
        "--path",
        type=Path,
        default=MAP_PATH,
        help="输出 JSON 路径",
    )
    args = parser.parse_args(argv)
    path: Path = args.path

    try:
        channels = _fetch_channels(args.warehouse_code)
    except Exception as exc:  # noqa: BLE001
        print(f"[FAIL] getShippingMethod 失败: {exc}", file=sys.stderr)
        return 1

    if not channels:
        print(f"[FAIL] 未返回任何运输方式 warehouse={args.warehouse_code}", file=sys.stderr)
        return 1

    existing = _load_existing(path)
    # 人工维护的订单名 → code，刷新 OMS 快照时保留
    aliases = existing.get("aliases") if isinstance(existing.get("aliases"), dict) else {}
    # 兼容旧版：整份 map 若像订单别名则并入 aliases（键不在本次 OMS name 中）
    old_map = existing.get("map") if isinstance(existing.get("map"), dict) else {}
    oms_map = _build_name_map(channels)
    for k, v in old_map.items():
        key, val = str(k).strip(), str(v).strip()
        if key and val and key not in oms_map and key not in aliases:
            aliases[key] = val

    aliases = {
        str(k).strip(): str(v).strip()
        for k, v in aliases.items()
        if str(k).strip() and str(v).strip()
    }

    payload = {
        "version": 1,
        "description": (
            "鸿羽 OMS getShippingMethod 快照。"
            "map=OMS中文名→code；aliases=订单运输方式中文→code（人工维护，刷新时保留）。"
        ),
        "warehouse_code": args.warehouse_code,
        "updated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "channels": channels,
        "map": oms_map,
        "aliases": aliases,
    }
    _dump_payload(payload, path)

    print(f"[OK] warehouse={args.warehouse_code}")
    print(f"  channels={len(channels)}  map={len(oms_map)}  aliases={len(aliases)}")
    print(f"  saved: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
