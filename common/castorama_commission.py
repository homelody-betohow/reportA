"""Castorama SKU 类目佣金：本机 JSON 映射（取代桌面 xlsx）。"""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

from common.style import Color

_COL_SKU = "SKU"
_COL_RATE = "佣金比"
_COL_MAPPED_RATE = "映射佣金比"

_DEFAULT_DESCRIPTION = (
    "Castorama SKU 类目佣金比例本机映射"
    "（取代桌面 castorama - SKU类目佣金比例.xlsx）。"
    "字段：SKU、佣金比。"
)


def castorama_commission_path(project_root: Path) -> Path:
    return project_root / "runtime" / "local" / "castorama_commission.json"


def normalize_sku(sku) -> str:
    """与 sku_mapping 一致：strip；剥尾缀 -NW。"""
    if sku is None or (isinstance(sku, float) and pd.isna(sku)):
        return ""
    s = str(sku).strip()
    if s.endswith("-NW"):
        s = s[:-3]
    return s


def _parse_rate(val) -> float | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def dump_commission_json(payload: dict, json_path: Path) -> None:
    """写出 castorama 佣金 JSON：items 中每个 {} 占一行，便于对照编辑。"""
    items = payload.get("items") if isinstance(payload, dict) else None
    if not isinstance(items, list):
        json_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        return

    meta = {k: v for k, v in payload.items() if k != "items"}
    lines = ["{"]
    for key, val in meta.items():
        lines.append(
            f"  {json.dumps(key, ensure_ascii=False)}: "
            f"{json.dumps(val, ensure_ascii=False)},"
        )
    item_field_order = (_COL_SKU, _COL_RATE)
    lines.append('  "items": [')
    for i, row in enumerate(items):
        if isinstance(row, dict):
            ordered = {k: row.get(k) for k in item_field_order}
            for k, v in row.items():
                if k not in ordered:
                    ordered[k] = v
            row = ordered
        row_json = json.dumps(row, ensure_ascii=False, separators=(", ", ": "))
        suffix = "," if i < len(items) - 1 else ""
        lines.append(f"    {row_json}{suffix}")
    lines.append("  ]")
    lines.append("}")
    lines.append("")
    json_path.write_text("\n".join(lines), encoding="utf-8")


def load_castorama_commission(json_path: Path, *, log_tag: str = "castorama") -> dict[str, float]:
    """
    读取 castorama_commission.json → {规范化SKU: 佣金比}。
    items 形如：[{"SKU": "...", "佣金比": 0.1}, ...]
    """
    if not json_path.is_file():
        return {}
    try:
        payload = json.loads(json_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        print(f"{Color.YELLOW}[{log_tag}] 无法读取 castorama 佣金 JSON {json_path}：{exc}{Color.RESET}")
        return {}

    items = payload.get("items") if isinstance(payload, dict) else None
    if not isinstance(items, list):
        print(f"{Color.YELLOW}[{log_tag}] JSON 缺少 items 列表，已跳过：{json_path}{Color.RESET}")
        return {}

    out: dict[str, float] = {}
    for row in items:
        if not isinstance(row, dict):
            continue
        rate = _parse_rate(row.get(_COL_RATE))
        if rate is None:
            continue
        sku = normalize_sku(row.get(_COL_SKU))
        if not sku:
            continue
        out[sku] = rate
    return out


def read_commission_payload(json_path: Path, *, log_tag: str = "castorama") -> dict:
    """读取 castorama_commission.json；文件不存在或损坏时返回空骨架。"""
    default = {
        "version": 1,
        "description": _DEFAULT_DESCRIPTION,
        "items": [],
    }
    if not json_path.is_file():
        return default
    try:
        payload = json.loads(json_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        print(f"{Color.YELLOW}[{log_tag}] 读取 {json_path} 失败，将重建：{exc}{Color.RESET}")
        return default
    if not isinstance(payload, dict):
        return default
    if not isinstance(payload.get("items"), list):
        payload["items"] = []
    payload.setdefault("version", 1)
    payload.setdefault("description", default["description"])
    return payload


def apply_castorama_commission_from_json(
    df: pd.DataFrame,
    json_path: Path,
    *,
    sku_col: str = _COL_SKU,
    mapped_col: str = _COL_MAPPED_RATE,
    log_tag: str = "castorama",
) -> pd.DataFrame:
    """用本机 JSON 按 SKU 填充「映射佣金比」（取代 Excel sku_mappings）。"""
    out = df.copy()
    rate_map = load_castorama_commission(json_path, log_tag=log_tag)
    keys = out[sku_col].map(normalize_sku)
    mapped = keys.map(rate_map)
    out[mapped_col] = pd.to_numeric(mapped, errors="coerce")

    hit = int(out[mapped_col].notna().sum())
    total = len(out)
    print(
        f"{Color.CYAN}[{log_tag}] castorama_commission.json 映射：{hit}/{total} 行命中「{mapped_col}」"
        f"\n  文件：{json_path}{Color.RESET}"
    )
    if not rate_map:
        print(
            f"{Color.YELLOW}[{log_tag}] JSON 为空或未启用；castorama 将依赖后续兜底规则{Color.RESET}"
        )
    return out


def merge_missing_into_castorama_commission_json(
    df: pd.DataFrame,
    json_path: Path,
    *,
    platform_col: str = "平台",
    fee_col: str = "映射平台费（佣金）",
    sku_col: str = _COL_SKU,
    log_tag: str = "castorama",
) -> int:
    """
    将 castorama 仍缺平台费的 SKU 追加进 JSON（佣金比=null，待手工填写）。
    返回新追加条数。
    """
    castorama = df[platform_col].astype(str).str.lower() == "castorama"
    empty = castorama & (
        df[fee_col].isna()
        | df[fee_col].astype(str).str.strip().isin(["", "nan", "None"])
    )
    miss_df = df.loc[empty]
    if miss_df.empty:
        return 0

    pending: dict[str, dict] = {}
    for _, r in miss_df.iterrows():
        sku = normalize_sku(r.get(sku_col))
        if not sku or sku in pending:
            continue
        pending[sku] = {_COL_SKU: sku, _COL_RATE: None}
    if not pending:
        return 0

    json_path.parent.mkdir(parents=True, exist_ok=True)
    payload = read_commission_payload(json_path, log_tag=log_tag)
    existing_items: list[dict] = []
    existing_keys: set[str] = set()

    for row in payload["items"]:
        if not isinstance(row, dict):
            continue
        key = normalize_sku(row.get(_COL_SKU))
        if not key:
            continue
        existing_keys.add(key)
        row[_COL_SKU] = key
        existing_items.append(row)

    n_added = 0
    for key, row in pending.items():
        if key in existing_keys:
            continue
        existing_items.append(row)
        n_added += 1

    existing_items.sort(key=lambda x: normalize_sku(x.get(_COL_SKU)))
    payload["items"] = existing_items
    dump_commission_json(payload, json_path)
    print(
        f"{Color.YELLOW}[{log_tag}] 已写入 {json_path}："
        f"新增待填 {n_added} 条，合计 {len(existing_items)} 条"
        f"（请填写「{_COL_RATE}」后重跑）{Color.RESET}"
    )
    return n_added
