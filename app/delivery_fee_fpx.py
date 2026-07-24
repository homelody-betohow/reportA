"""4PX 派送费试算：从（已完成-4）筛 4PX- 行，调用 com.css.price_calculator，写入 goods_delivery_fee。

用法（项目根目录）::

    python -m app.delivery_fee_fpx
    python -m app.delivery_fee_fpx --sku E52011003
    python -m app.delivery_fee_fpx --sku E52011003,E59032000 --limit 20

传 ``--sku`` 时忽略 goods_delivery_fee 缓存，强制调 API 并 UPSERT。

依赖映射：
  - ``api/fpx/warehouse_map.json``：ERP 仓码 → 4PX warehouse_code
  - ``api/fpx/logistics_product_map.json``：运输方式中文 → product_code
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import re
import sys
import time
from collections import Counter
from datetime import date, timedelta
from pathlib import Path
from typing import Any

import pandas as pd
import pymysql.cursors

# 须在 import config/common 之前：加载项目根到 sys.path
_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
assert _spec.loader is not None
_spec.loader.exec_module(_epr_mod)
_PROJECT_ROOT = _epr_mod.bootstrap(__file__)

from api.fpx import FpxClient  # noqa: E402
from api.fpx.exceptions import FpxError  # noqa: E402
from common.style import Color  # noqa: E402
from config.A0_paths import DESKTOP_ROOT  # noqa: E402
from config.A0_set_date import folder_name, shared_date  # noqa: E402
from database.db_connection import get_db_manager  # noqa: E402

PROVIDER_CODE = "4PX"
SERVICE_CODE = "FB4"
FEE_TABLE = "goods_delivery_fee"
ZIPCODE_STORE = "*"
CACHE_DAYS = 30
_KEY_CHUNK = 200
_WAREHOUSE_BRACKET_RE = re.compile(r"\[.*\]\s*$")

# 缓存命中键：(product_sku, dispatch_warehouse, destination_country, dispatch_channel)
FeeKey = tuple[str, str, str, str]

WAREHOUSE_MAP_PATH = Path(_PROJECT_ROOT) / "api" / "fpx" / "warehouse_map.json"
PRODUCT_MAP_PATH = Path(_PROJECT_ROOT) / "api" / "fpx" / "logistics_product_map.json"

# 订单无邮编时，按目的国给 API 用的默认邮编（入库仍写 destination_zipcode='*'）
_DEFAULT_POSTCODE: dict[str, str] = {
    "FR": "75011",
    "DE": "10115",
    "ES": "28001",
    "IT": "00118",
    "BE": "1000",
    "NL": "1011",
    "PL": "00-001",
    "CZ": "11000",
    "AT": "1010",
    "PT": "1000-001",
}

UPSERT_SQL = f"""
INSERT INTO `{FEE_TABLE}` (
    provider_code, product_sku, length_cm, width_cm, height_cm, weight_kg,
    dispatch_warehouse, dispatch_channel, destination_country, destination_zipcode,
    dispatch_date, dispatch_fee, currency
) VALUES (
    %s, %s, %s, %s, %s, %s,
    %s, %s, %s, %s,
    %s, %s, %s
)
ON DUPLICATE KEY UPDATE
    length_cm = VALUES(length_cm),
    width_cm = VALUES(width_cm),
    height_cm = VALUES(height_cm),
    weight_kg = VALUES(weight_kg),
    dispatch_channel = VALUES(dispatch_channel),
    dispatch_fee = VALUES(dispatch_fee),
    currency = VALUES(currency)
"""


def _normalize_currency(val: Any) -> str | None:
    """API currency → 入库 currency（如 EUR）；空则 None。"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip().upper()
    return s or None


def _print_api_fail_body(label: str, body: Any) -> None:
    """接口失败时打印返回内容。"""
    if body is None:
        return
    if isinstance(body, (dict, list)):
        text = json.dumps(body, ensure_ascii=False, indent=2)
    else:
        text = str(body)
    print(f"{Color.RED}[API返回-失败] {label}\n{text}{Color.RESET}")


def _normalize_sku(sku: Any) -> str:
    if sku is None or (isinstance(sku, float) and pd.isna(sku)):
        return ""
    s = str(sku).strip()
    if s.endswith("-NW"):
        s = s[:-3]
    return s


def _load_name_code_map(path: Path, *, label: str) -> dict[str, str]:
    """读取 JSON：aliases（人工）优先，其次 map → code。"""
    if not path.is_file():
        raise FileNotFoundError(f"缺少{label}映射文件: {path}")
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"无法解析{label}映射 {path}: {exc}") from exc

    if not isinstance(payload, dict):
        raise ValueError(f"{path} 格式错误，应为 JSON 对象")

    out: dict[str, str] = {}
    for section in ("map", "aliases"):
        raw = payload.get(section)
        if not isinstance(raw, dict):
            continue
        for k, v in raw.items():
            key = str(k).strip()
            val = str(v).strip() if v is not None else ""
            if key and val:
                out[key] = val  # aliases 后写，覆盖同名 map

    if not out:
        raise ValueError(f"{path} 的 map/aliases 均为空")
    return out


def _norm_text(val: Any) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).strip()


def _erp_warehouse_code(warehouse: Any) -> str:
    """4PX-LM-BC-FR[中文名] → 4PX-LM-BC-FR"""
    s = _norm_text(warehouse)
    if not s:
        return ""
    return _WAREHOUSE_BRACKET_RE.sub("", s).strip()


def _fpx_warehouse_code(erp_code: str, warehouse_map: dict[str, str]) -> str:
    """ERP 仓码 → 4PX warehouse_code；先查 map，再按国别启发式。"""
    code = _norm_text(erp_code)
    if not code:
        return ""
    hit = warehouse_map.get(code) or warehouse_map.get(code.upper())
    if hit:
        return hit

    upper = code.upper()
    if "-FR" in upper or upper.endswith("FR") or "FRANCE" in upper:
        return "FRCDGA"
    if "-CZ" in upper or "JK-" in upper or upper.endswith("CZ"):
        return "CZPRGA"
    if "-BLM" in upper:
        return "DEKARA"
    if "4PX" in upper or "-DE" in upper or upper == "DE":
        return "DEFRAA"
    return ""


def _destination_country(country: Any, fee_class: Any) -> str:
    c = _norm_text(country).upper()
    if c:
        return c
    fc = _norm_text(fee_class)
    # 4PX-FR / HY-DE → FR / DE
    for prefix in ("4PX-", "HY-"):
        if fc.upper().startswith(prefix) and len(fc) > len(prefix):
            return fc.split("-", 1)[1].strip().upper()
    return ""


def _api_postcode(postcode: str | None, country: str) -> str | None:
    """试算用邮编：订单邮编优先，否则国别默认。"""
    pc = _norm_text(postcode)
    if pc:
        return pc.upper() if any(ch.isalpha() for ch in pc) else pc
    return _DEFAULT_POSTCODE.get(country.upper())


def _completed4_path() -> Path:
    return Path(
        fr"{DESKTOP_ROOT}\{folder_name}{shared_date}\订单统计"
        fr"\(已完成-4)订单统计-{shared_date}.xlsx"
    )


def _parse_sku_filter(raw: str | None) -> set[str] | None:
    if not raw or not raw.strip():
        return None
    out = {_normalize_sku(x) for x in raw.split(",") if _normalize_sku(x)}
    return out or None


def _load_fpx_rows(path: Path, sku_filter: set[str] | None) -> pd.DataFrame:
    if not path.is_file():
        raise FileNotFoundError(f"找不到已完成-4 文件: {path}")
    df = pd.read_excel(path)
    required = {"派送费-映射分类", "SKU", "仓库", "运输方式", "国家"}
    missing = required - set(df.columns)
    if missing:
        raise KeyError(f"已完成-4 缺少列: {sorted(missing)}")

    fpx = df[df["派送费-映射分类"].astype(str).str.startswith("4PX", na=False)].copy()
    fpx["__sku"] = fpx["SKU"].map(_normalize_sku)
    if sku_filter is not None:
        fpx = fpx[fpx["__sku"].isin(sku_filter)].copy()
    return fpx


def _build_jobs(
    fpx: pd.DataFrame,
    warehouse_map: dict[str, str],
    product_map: dict[str, str],
) -> tuple[list[dict[str, Any]], list[str]]:
    """按 (SKU, 4PX仓, 目的国, product_code) 去重，返回待试算 jobs 与跳过原因。"""
    skips: list[str] = []
    seen: set[FeeKey] = set()
    jobs: list[dict[str, Any]] = []

    for _, row in fpx.iterrows():
        sku = _normalize_sku(row.get("SKU"))
        erp_wh = _erp_warehouse_code(row.get("仓库"))
        country = _destination_country(row.get("国家"), row.get("派送费-映射分类"))
        shipping_cn = _norm_text(row.get("运输方式"))
        postcode = _norm_text(row.get("邮编")) if "邮编" in fpx.columns else ""

        if not sku:
            skips.append("空 SKU")
            continue
        if not erp_wh:
            skips.append(f"空仓库 sku={sku}")
            continue
        if not country:
            skips.append(f"空目的国 sku={sku}")
            continue

        product_code = product_map.get(shipping_cn)
        if not product_code:
            skips.append(f"未映射运输方式 {shipping_cn!r} sku={sku}")
            continue

        fpx_wh = _fpx_warehouse_code(erp_wh, warehouse_map)
        if not fpx_wh:
            skips.append(f"未映射仓库 {erp_wh!r} sku={sku}")
            continue

        key = _fee_cache_key(sku, fpx_wh, country, product_code)
        if key in seen:
            continue
        seen.add(key)

        jobs.append(
            {
                "product_sku": sku,
                # 入库 dispatch_warehouse = 4PX warehouse_code（如 FRCDGA）
                "dispatch_warehouse": fpx_wh,
                "fpx_warehouse": fpx_wh,
                "destination_country": country,
                "product_code": product_code,  # → dispatch_channel
                "postcode": postcode or None,
                "shipping_cn": shipping_cn,
                "erp_warehouse": erp_wh,
            }
        )
    return jobs, skips


def _fetch_sku_dims(skus: list[str]) -> dict[str, dict[str, Any]]:
    """product_sku → weight_kg / length_cm / width_cm / height_cm。"""
    unique = sorted({s for s in skus if s})
    out: dict[str, dict[str, Any]] = {}
    if not unique:
        return out

    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        for i in range(0, len(unique), _KEY_CHUNK):
            chunk = unique[i : i + _KEY_CHUNK]
            placeholders = ",".join(["%s"] * len(chunk))
            sql = f"""
                SELECT product_sku, unit_weight_g,
                       inner_box_l_cm, inner_box_w_cm, inner_box_h_cm
                FROM product_sku
                WHERE product_sku IN ({placeholders})
            """
            cur.execute(sql, chunk)
            for row in cur.fetchall():
                sku = _normalize_sku(row.get("product_sku"))
                if not sku or sku in out:
                    continue
                w_g = row.get("unit_weight_g")
                weight_kg = float(w_g) / 1000.0 if w_g is not None else None
                out[sku] = {
                    "weight_kg": weight_kg,
                    "length_cm": float(row["inner_box_l_cm"]) if row.get("inner_box_l_cm") is not None else None,
                    "width_cm": float(row["inner_box_w_cm"]) if row.get("inner_box_w_cm") is not None else None,
                    "height_cm": float(row["inner_box_h_cm"]) if row.get("inner_box_h_cm") is not None else None,
                }
    finally:
        cur.close()
        conn.close()
    return out


def _fee_cache_key(
    sku: str, warehouse: str, country: str, channel: str
) -> FeeKey:
    return (sku, warehouse, country, channel)


def _fetch_fresh_fee_keys(
    jobs: list[dict[str, Any]],
    *,
    within_days: int = CACHE_DAYS,
) -> dict[FeeKey, date]:
    """批量查询 goods_delivery_fee：同 provider+SKU+仓+渠道+国+zip=* 且 dispatch_date 在 N 天内。"""
    skus = sorted({j["product_sku"] for j in jobs if j.get("product_sku")})
    fresh: dict[FeeKey, date] = {}
    if not skus:
        return fresh

    since = date.today() - timedelta(days=within_days)
    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor(pymysql.cursors.DictCursor)
    try:
        for i in range(0, len(skus), _KEY_CHUNK):
            chunk = skus[i : i + _KEY_CHUNK]
            placeholders = ",".join(["%s"] * len(chunk))
            sql = f"""
                SELECT product_sku, dispatch_warehouse, dispatch_channel,
                       destination_country,
                       MAX(dispatch_date) AS max_dispatch_date
                FROM `{FEE_TABLE}`
                WHERE provider_code = %s
                  AND product_sku IN ({placeholders})
                  AND destination_zipcode = %s
                  AND dispatch_date IS NOT NULL
                  AND dispatch_date >= %s
                GROUP BY product_sku, dispatch_warehouse, dispatch_channel,
                         destination_country
            """
            cur.execute(sql, [PROVIDER_CODE, *chunk, ZIPCODE_STORE, since])
            for row in cur.fetchall():
                sku = _normalize_sku(row.get("product_sku"))
                wh = _norm_text(row.get("dispatch_warehouse"))
                channel = _norm_text(row.get("dispatch_channel"))
                country = _norm_text(row.get("destination_country")).upper()
                d = row.get("max_dispatch_date")
                if not sku or not wh or not country or d is None:
                    continue
                key = _fee_cache_key(sku, wh, country, channel)
                if isinstance(d, date):
                    fresh[key] = d
                else:
                    fresh[key] = date.fromisoformat(str(d)[:10])
    finally:
        cur.close()
        conn.close()
    return fresh


def _upsert_fee(row: tuple[Any, ...]) -> None:
    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor()
    try:
        cur.execute(UPSERT_SQL, row)
        conn.commit()
    finally:
        cur.close()
        conn.close()


def _pick_quote(
    data: Any, *, product_code: str
) -> tuple[dict[str, Any] | None, str | None]:
    """从 price_calculator data 列表中选取目标 product_code 报价。"""
    if isinstance(data, dict):
        items = [data]
    elif isinstance(data, list):
        items = [x for x in data if isinstance(x, dict)]
    else:
        return None, "data 非 list/dict"

    if not items:
        return None, "data 为空"

    matched = [
        x for x in items if _norm_text(x.get("product_code")).upper() == product_code.upper()
    ]
    if matched:
        return matched[0], None
    if len(items) == 1:
        return items[0], None
    codes = [_norm_text(x.get("product_code")) for x in items]
    return None, f"未命中 product_code={product_code}，返回={codes}"


def _run_jobs(
    jobs: list[dict[str, Any]],
    dims: dict[str, dict[str, Any]],
    fresh_keys: dict[FeeKey, date],
    *,
    dry_run: bool = False,
    verbose_api: bool = False,
) -> tuple[int, int, int, int, int, list[str]]:
    """返回 (ok, fail, skip_weight, skip_dims, skip_cached, fail_samples)。"""
    client = FpxClient.from_env()
    dispatch_date = date.today()
    ok = fail = skip_weight = skip_dims = skip_cached = 0
    fail_samples: list[str] = []

    for idx, job in enumerate(jobs, 1):
        sku = job["product_sku"]
        channel = job["product_code"]
        cache_key = _fee_cache_key(
            sku, job["dispatch_warehouse"], job["destination_country"], channel
        )
        label = (
            f"[{idx}/{len(jobs)}] {sku} {job['dispatch_warehouse']}→"
            f"{job['destination_country']} channel={channel}"
        )

        cached_date = fresh_keys.get(cache_key)
        if cached_date is not None:
            skip_cached += 1
            continue

        dim = dims.get(sku) or {}
        weight = dim.get("weight_kg")
        length = dim.get("length_cm")
        width = dim.get("width_cm")
        height = dim.get("height_cm")

        if weight is None or weight <= 0:
            skip_weight += 1
            print(f"{Color.YELLOW}[跳过-无重量] {label}{Color.RESET}")
            continue
        if any(v is None or float(v) <= 0 for v in (length, width, height)):
            skip_dims += 1
            print(f"{Color.YELLOW}[跳过-无尺寸] {label}{Color.RESET}")
            continue

        country = job["destination_country"]
        post_code = _api_postcode(job.get("postcode"), country)
        if not post_code:
            fail += 1
            sample = f"{label} 无可用邮编（订单空且无国别默认）"
            if len(fail_samples) < 10:
                fail_samples.append(sample)
            print(f"{Color.RED}[FAIL] {sample}{Color.RESET}")
            continue

        weight_g = round(float(weight) * 1000.0, 3)
        destination = {"country": country, "post_code": post_code}
        req = {
            "service_code": SERVICE_CODE,
            "warehouse_code": job["fpx_warehouse"],
            "weight": weight_g,
            "length": float(length),
            "width": float(width),
            "height": float(height),
            "destination": destination,
            "product_codes": [channel],
            "billing_time": int(time.time() * 1000),
        }
        if verbose_api:
            print(
                f"{Color.CYAN}[API请求] {label}\n"
                f"{json.dumps(req, ensure_ascii=False, indent=2)}{Color.RESET}"
            )

        try:
            resp = client.price_calculator(
                service_code=SERVICE_CODE,
                warehouse_code=job["fpx_warehouse"],
                weight=weight_g,
                length=float(length),
                width=float(width),
                height=float(height),
                destination=destination,
                billing_time=req["billing_time"],
                product_codes=[channel],
            )
        except FpxError as exc:
            fail += 1
            sample = f"{label} ERR={exc}"
            if len(fail_samples) < 10:
                fail_samples.append(sample)
            print(f"{Color.RED}[FAIL] {sample}{Color.RESET}")
            _print_api_fail_body(label, getattr(exc, "raw", None))
            continue

        if verbose_api:
            print(
                f"{Color.CYAN}[API返回] {label}\n"
                f"{json.dumps(resp, ensure_ascii=False, indent=2)}{Color.RESET}"
            )

        quote, pick_err = _pick_quote(resp.get("data"), product_code=channel)
        if quote is None:
            fail += 1
            sample = f"{label} {pick_err or '无报价'}"
            if len(fail_samples) < 10:
                fail_samples.append(sample)
            print(f"{Color.RED}[FAIL] {sample}{Color.RESET}")
            _print_api_fail_body(label, resp)
            continue

        total = quote.get("total_amount")
        currency = _normalize_currency(quote.get("currency"))
        try:
            fee = float(total)
        except (TypeError, ValueError):
            fail += 1
            sample = f"{label} 无 total_amount"
            if len(fail_samples) < 10:
                fail_samples.append(sample)
            print(f"{Color.RED}[FAIL] {sample}{Color.RESET}")
            _print_api_fail_body(label, resp)
            continue

        db_row = (
            PROVIDER_CODE,
            sku,
            length,
            width,
            height,
            weight,
            job["dispatch_warehouse"],
            channel,  # dispatch_channel = 4PX product_code
            country,
            ZIPCODE_STORE,
            dispatch_date,
            fee,
            currency,
        )
        fee_label = f"total_amount={fee}" + (f" {currency}" if currency else "")
        if dry_run:
            print(f"{Color.CYAN}[dry-run] {label} {fee_label}{Color.RESET}")
            ok += 1
            fresh_keys[cache_key] = dispatch_date
            continue

        try:
            _upsert_fee(db_row)
            ok += 1
            fresh_keys[cache_key] = dispatch_date
            print(f"{Color.GREEN}[OK] {label} {fee_label}{Color.RESET}")
        except Exception as exc:  # noqa: BLE001 — 单条失败不中断
            fail += 1
            sample = f"{label} DB={exc}"
            if len(fail_samples) < 10:
                fail_samples.append(sample)
            print(f"{Color.RED}[FAIL] {sample}{Color.RESET}")

    return ok, fail, skip_weight, skip_dims, skip_cached, fail_samples


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="4PX 派送费试算写入 goods_delivery_fee")
    parser.add_argument(
        "--sku",
        default="",
        help="只处理指定 SKU，逗号分隔（规范化：strip，剥 -NW）；"
        "传入时忽略库内缓存，强制调 API 并更新 goods_delivery_fee",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=0,
        help="调试：限制去重后试算条数（0=不限制）",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只调 API 不写库",
    )
    args = parser.parse_args(argv)

    path = _completed4_path()
    sku_filter = _parse_sku_filter(args.sku)
    print(f"{Color.CYAN}输入: {path}{Color.RESET}")
    if sku_filter:
        print(f"{Color.CYAN}SKU 过滤: {sorted(sku_filter)}{Color.RESET}")

    fpx = _load_fpx_rows(path, sku_filter)
    print(f"4PX- 行数: {len(fpx)}")
    if fpx.empty:
        print(f"{Color.YELLOW}无 4PX- 行可处理{Color.RESET}")
        return 0

    warehouse_map = _load_name_code_map(WAREHOUSE_MAP_PATH, label="仓库")
    product_map = _load_name_code_map(PRODUCT_MAP_PATH, label="物流产品")
    print(
        f"{Color.CYAN}仓库映射: {WAREHOUSE_MAP_PATH} ({len(warehouse_map)} 条){Color.RESET}"
    )
    print(
        f"{Color.CYAN}物流产品映射: {PRODUCT_MAP_PATH} ({len(product_map)} 条){Color.RESET}"
    )

    jobs, map_skips = _build_jobs(fpx, warehouse_map, product_map)
    if args.limit and args.limit > 0:
        jobs = jobs[: args.limit]

    if map_skips:
        for reason, cnt in Counter(map_skips).most_common(15):
            print(f"{Color.YELLOW}[映射跳过×{cnt}] {reason}{Color.RESET}")

    print(f"去重后试算数: {len(jobs)}")
    if not jobs:
        print(f"{Color.YELLOW}无有效试算任务{Color.RESET}")
        return 0

    if sku_filter is not None:
        fresh_keys: dict[FeeKey, date] = {}
        print(f"{Color.CYAN}--sku 模式：忽略库内缓存，强制请求 API 并更新{Color.RESET}")
    else:
        fresh_keys = _fetch_fresh_fee_keys(jobs, within_days=CACHE_DAYS)
        print(
            f"{Color.CYAN}库内 {CACHE_DAYS} 天内已有记录（按 SKU+仓+渠道+国）: "
            f"{len(fresh_keys)}{Color.RESET}"
        )

    dims = _fetch_sku_dims([j["product_sku"] for j in jobs])
    print(f"命中 product_sku 数量: {len(dims)}/{len({j['product_sku'] for j in jobs})}")

    ok, fail, skip_weight, skip_dims, skip_cached, fail_samples = _run_jobs(
        jobs,
        dims,
        fresh_keys,
        dry_run=args.dry_run,
        verbose_api=sku_filter is not None,
    )
    print(
        f"\n{Color.CYAN}汇总: 成功={ok} 失败={fail} "
        f"跳过(已缓存)={skip_cached} 跳过(无重量)={skip_weight} "
        f"跳过(无尺寸)={skip_dims} 试算任务={len(jobs)}{Color.RESET}"
    )
    if fail_samples:
        print(f"{Color.RED}失败样例:{Color.RESET}")
        for s in fail_samples:
            print(f"  - {s}")

    return 1 if fail and ok == 0 else 0


if __name__ == "__main__":
    raise SystemExit(main())
