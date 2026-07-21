"""HY 派送费试算：从（已完成-4）筛 HY- 行，调用 getCalculateFee，写入 goods_delivery_fee。

用法（项目根目录）::

    python -m app.delivery_fee_hy
    python -m app.delivery_fee_hy --sku E52011003
    python -m app.delivery_fee_hy --sku E52011003,E59032000 --limit 20

传 ``--sku`` 时忽略 goods_delivery_fee 缓存，强制调 API 并 UPSERT。
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import re
import sys
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

from api.hy_oms import HyOmsClient  # noqa: E402
from api.hy_oms.exceptions import HyOmsError  # noqa: E402
from common.style import Color  # noqa: E402
from config.A0_paths import DESKTOP_ROOT  # noqa: E402
from config.A0_set_date import folder_name, shared_date  # noqa: E402

from database.db_connection import get_db_manager  # noqa: E402

PROVIDER_CODE = "HY"
FEE_TABLE = "goods_delivery_fee"
ZIPCODE_STORE = "*"
CACHE_DAYS = 30
_KEY_CHUNK = 200
_WAREHOUSE_BRACKET_RE = re.compile(r"\[.*\]\s*$")

# 缓存命中键：(product_sku, dispatch_warehouse, destination_country, dispatch_channel)
FeeKey = tuple[str, str, str, str]

SHIPPING_METHOD_MAP_PATH = (
    Path(_PROJECT_ROOT) / "api" / "hy_oms" / "shipping_method_map.json"
)

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
    """OMS data.currency_code → 入库 currency（如 EUR）；空则 None。"""
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


def _load_shipping_method_map(path: Path = SHIPPING_METHOD_MAP_PATH) -> dict[str, str]:
    """读取 shipping_method_map.json：aliases（订单名）优先，其次 map（OMS 中文名）→ code。"""
    if not path.is_file():
        raise FileNotFoundError(
            f"缺少运输方式映射文件: {path}\n"
            "请先执行: python -m api.hy_oms.shipping_method_update"
        )
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"无法解析运输方式映射 {path}: {exc}") from exc

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
        raise ValueError(f"{path} 的 map/aliases 均为空，请先运行 shipping_method_update")
    return out


def _norm_text(val: Any) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).strip()


def _erp_warehouse_code(warehouse: Any) -> str:
    """HY-OTTO-DE-01[中文名] → HY-OTTO-DE-01"""
    s = _norm_text(warehouse)
    if not s:
        return ""
    return _WAREHOUSE_BRACKET_RE.sub("", s).strip()


def _oms_warehouse_code(erp_code: str) -> str:
    """ERP 仓码 → OMS warehouse_code：-03 → DE03，其余默认 DEHY。"""
    code = erp_code.upper()
    if code.endswith("-03") or code in {"DE03", "DEHY-03", "DEHY03"}:
        return "DE03"
    return "DEHY"


def _destination_country(country: Any, fee_class: Any) -> str:
    c = _norm_text(country).upper()
    if c:
        return c
    fc = _norm_text(fee_class)
    if fc.upper().startswith("HY-") and len(fc) > 3:
        return fc.split("-", 1)[1].strip().upper()
    return ""


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


def _load_hy_rows(path: Path, sku_filter: set[str] | None) -> pd.DataFrame:
    if not path.is_file():
        raise FileNotFoundError(f"找不到已完成-4 文件: {path}")
    df = pd.read_excel(path)
    required = {"派送费-映射分类", "SKU", "仓库", "运输方式", "国家"}
    missing = required - set(df.columns)
    if missing:
        raise KeyError(f"已完成-4 缺少列: {sorted(missing)}")

    hy = df[df["派送费-映射分类"].astype(str).str.startswith("HY-", na=False)].copy()
    hy["__sku"] = hy["SKU"].map(_normalize_sku)
    if sku_filter is not None:
        hy = hy[hy["__sku"].isin(sku_filter)].copy()
    return hy


def _build_jobs(
    hy: pd.DataFrame,
    shipping_map: dict[str, str],
) -> tuple[list[dict[str, Any]], list[str]]:
    """按 (SKU, OMS仓, 目的国, 渠道) 去重，返回待试算 jobs 与跳过原因。"""
    skips: list[str] = []
    seen: set[FeeKey] = set()
    jobs: list[dict[str, Any]] = []

    for _, row in hy.iterrows():
        sku = _normalize_sku(row.get("SKU"))
        erp_wh = _erp_warehouse_code(row.get("仓库"))
        country = _destination_country(row.get("国家"), row.get("派送费-映射分类"))
        shipping_cn = _norm_text(row.get("运输方式"))
        postcode = _norm_text(row.get("邮编")) if "邮编" in hy.columns else ""

        if not sku:
            skips.append("空 SKU")
            continue
        if not erp_wh:
            skips.append(f"空仓库 sku={sku}")
            continue
        if not country:
            skips.append(f"空目的国 sku={sku}")
            continue

        method = shipping_map.get(shipping_cn)
        if not method:
            skips.append(f"未映射运输方式 {shipping_cn!r} sku={sku}")
            continue

        oms_wh = _oms_warehouse_code(erp_wh)
        key = _fee_cache_key(sku, oms_wh, country, method)
        if key in seen:
            continue
        seen.add(key)

        jobs.append(
            {
                "product_sku": sku,
                # 入库 dispatch_warehouse = OMS warehouse_code（DEHY / DE03）
                "dispatch_warehouse": oms_wh,
                "oms_warehouse": oms_wh,
                "destination_country": country,
                "shipping_method": method,  # → dispatch_channel
                "postcode": postcode or None,
                "shipping_cn": shipping_cn,
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
    """批量查询 goods_delivery_fee：同 provider+SKU+仓+渠道+国+zip=* 且 dispatch_date 在 N 天内。

    返回命中键 → 最新 dispatch_date（同键多行取最新）。
    """
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


def _run_jobs(
    jobs: list[dict[str, Any]],
    dims: dict[str, dict[str, Any]],
    fresh_keys: dict[FeeKey, date],
    *,
    dry_run: bool = False,
    verbose_api: bool = False,
) -> tuple[int, int, int, int, list[str]]:
    """返回 (ok, fail, skip_weight, skip_cached, fail_samples)。

    verbose_api=True（传了 --sku）时打印每次 getCalculateFee 的请求与返回。
    """
    client = HyOmsClient.from_env()
    dispatch_date = date.today()
    ok = fail = skip_weight = skip_cached = 0
    fail_samples: list[str] = []

    for idx, job in enumerate(jobs, 1):
        sku = job["product_sku"]
        channel = job["shipping_method"]
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
            # print(
            #     # f"{Color.YELLOW}[跳过-已有{CACHE_DAYS}天内记录] {label} "
            #     # f"dispatch_date={cached_date}{Color.RESET}"
            # )
            continue

        dim = dims.get(sku) or {}
        weight = dim.get("weight_kg")
        if weight is None or weight <= 0:
            skip_weight += 1
            print(f"{Color.YELLOW}[跳过-无重量] {label}{Color.RESET}")
            continue

        req = {
            "warehouse_code": job["oms_warehouse"],
            "country_code": job["destination_country"],
            "shipping_method": channel,
            "weight": float(weight),
            "postcode": job.get("postcode"),
            "length": dim.get("length_cm"),
            "width": dim.get("width_cm"),
            "height": dim.get("height_cm"),
        }
        if verbose_api:
            print(
                f"{Color.CYAN}[API请求] {label}\n"
                f"{json.dumps(req, ensure_ascii=False, indent=2)}{Color.RESET}"
            )

        try:
            resp = client.get_calculate_fee(**req)
        except HyOmsError as exc:
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

        data = resp.get("data") or {}
        if not isinstance(data, dict):
            data = {}
        total = data.get("totalFee")
        currency = _normalize_currency(
            data.get("currency_code") or data.get("currency")
        )
        try:
            fee = float(total)
        except (TypeError, ValueError):
            fail += 1
            sample = f"{label} 无 totalFee"
            if len(fail_samples) < 10:
                fail_samples.append(sample)
            print(f"{Color.RED}[FAIL] {sample}{Color.RESET}")
            _print_api_fail_body(label, resp)
            continue

        db_row = (
            PROVIDER_CODE,
            sku,
            dim.get("length_cm"),
            dim.get("width_cm"),
            dim.get("height_cm"),
            weight,
            job["dispatch_warehouse"],
            channel,  # dispatch_channel = OMS shipping_method
            job["destination_country"],
            ZIPCODE_STORE,
            dispatch_date,
            fee,
            currency,
        )
        fee_label = f"totalFee={fee}" + (f" {currency}" if currency else "")
        if dry_run:
            print(f"{Color.CYAN}[dry-run] {label} {fee_label}{Color.RESET}")
            ok += 1
            # dry-run 也记入 fresh，避免同批重复打 API
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

    return ok, fail, skip_weight, skip_cached, fail_samples


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="HY 派送费试算写入 goods_delivery_fee")
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

    hy = _load_hy_rows(path, sku_filter)
    print(f"HY- 行数: {len(hy)}")
    if hy.empty:
        print(f"{Color.YELLOW}无 HY- 行可处理{Color.RESET}")
        return 0

    shipping_map = _load_shipping_method_map()
    print(
        f"{Color.CYAN}运输方式映射: {SHIPPING_METHOD_MAP_PATH} "
        f"({len(shipping_map)} 条){Color.RESET}"
    )

    jobs, map_skips = _build_jobs(hy, shipping_map)
    if args.limit and args.limit > 0:
        jobs = jobs[: args.limit]

    if map_skips:
        for reason, cnt in Counter(map_skips).most_common(15):
            print(f"{Color.YELLOW}[映射跳过×{cnt}] {reason}{Color.RESET}")

    print(f"去重后试算数: {len(jobs)}")
    if not jobs:
        print(f"{Color.YELLOW}无有效试算任务{Color.RESET}")
        return 0

    # 传了 --sku：强制打 API 写库，不查/不跳过库内缓存
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

    ok, fail, skip_weight, skip_cached, fail_samples = _run_jobs(
        jobs,
        dims,
        fresh_keys,
        dry_run=args.dry_run,
        verbose_api=sku_filter is not None,
    )
    print(
        f"\n{Color.CYAN}汇总: 成功={ok} 失败={fail} "
        f"跳过(已缓存)={skip_cached} 跳过(无重量)={skip_weight} "
        f"试算任务={len(jobs)}{Color.RESET}"
    )
    if fail_samples:
        print(f"{Color.RED}失败样例:{Color.RESET}")
        for s in fail_samples:
            print(f"  - {s}")

    return 1 if fail and ok == 0 else 0


if __name__ == "__main__":
    raise SystemExit(main())
