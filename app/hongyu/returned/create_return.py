"""从钉钉在线表格读取退件申请，调用鸿羽 OMS ``createReturnBill``。

表格读写委托 ``api.ding_disk.workbook.Workbook``；创建退件委托
``api.hy_oms.request.create_return``。

文档 ID（workbookId / nodeId）::
    EpGBa2Lm8aDaZ57lTwEMk9boJgN7R35y

默认模式为 **回邮退件**（``return_identification=1``），因表格含寄件地址与标签列。
同一 ``ERP订单号``（空则 ``平台订单号``）多行会合并为一条退件单的 ``items``。

推送 API 前先查 ``sales_order_shipped``（``ERP订单号`` = ``order_no``）：
命中则回写 ``OMS原始订单号``=``provider_order_no``、``店铺名``=``shop_name_en``；
表格 ``城市``/``邮编`` 为空时用 ``city``/``postal_code`` 回填，再创建；
未命中则只回写 ``任务信息``，不调 API。

表格列 → API（回邮）::
    ERP订单号 / 平台订单号  →  reference_no
    产品SKU + 退件数量      →  items[].product_sku / quantity
    退件原因                →  return_desc
    收件人/国家/街道/门牌号/城市/邮编 → sender_info
    店铺名                  →  seller_store（优先用 shipped.shop_name_en）
    可选列 邮箱/电话/手机   →  sender_email（非必填）/ sender_phone（空则默认 0000000000）
    是否良品上架            →  items[].process（是→1，否→3，默认 1）

推送前按 OMS 文档本地校验寄件人必填（邮编/电话/街道/门牌号等；邮箱非必填）；
缺项不调 API，错误写入任务信息。注意：空值不得用 ``-`` 占位，否则 OMS 会当成已填而放行。

成功/失败回写::
    OMS原始订单号 / 店铺名 ← sales_order_shipped（命中时）
    城市 / 邮编 ← sales_order_shipped.city / postal_code（表格为空时回填）
    OMS退件订单号 ← return_code（成功时）
    进度          ← 0-100 数值；创建成功写 30
    任务信息      ← 成功摘要，或未找到原单 / API 错误原文

用法（项目根目录）::

    python app/hongyu/returned/create_return.py --dry-run
    python app/hongyu/returned/create_return.py --list-sheets
    python app/hongyu/returned/create_return.py --limit 1 --preview 5
    python app/hongyu/returned/create_return.py --sm-code DEGLS-RMA
    python app/hongyu/returned/create_return.py --mode standard --return-type S
    python app/hongyu/returned/create_return.py --verify 0
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple

import pandas as pd
import pymysql.cursors

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

from api.ding_disk.exceptions import DingDiskError  # noqa: E402
from api.ding_disk.workbook import Workbook, clean_cell  # noqa: E402
from api.hy_oms import HyOmsClient  # noqa: E402
from api.hy_oms.exceptions import HyOmsError  # noqa: E402
from api.hy_oms.request.create_return import build_params  # noqa: E402
from database.db_connection import get_db_manager  # noqa: E402

WORKBOOK_ID = "EpGBa2Lm8aDaZ57lTwEMk9boJgN7R35y"
DEFAULT_SHEET = "Sheet1"
DEFAULT_WAREHOUSE = "DEHY"
DEFAULT_SM_CODE = "DEGLS-RMA"
DEFAULT_PROCESS = "1"
DEFAULT_RETURN_TYPE = "S"
DEFAULT_SENDER_PHONE = "0000000000"
SHIPPED_TABLE = "sales_order_shipped"
DB_BATCH_SIZE = 500

# 分组键优先级
GROUP_KEY_COLS = ("ERP订单号", "平台订单号")

# 回写列
COL_ERP = "ERP订单号"
COL_SHOP = "店铺名"
COL_CITY = "城市"
COL_ZIP = "邮编"
COL_OMS_RETURN = "OMS退件订单号"
COL_PROGRESS = "进度"
COL_TASK = "任务信息"
COL_OMS_ORDER = "OMS原始订单号"
COL_LABEL_TRACKING = "标签跟踪号"
TASK_MSG_MAX = 500
# 进度为 0-100 数值；创建退件成功后记 30
PROGRESS_CREATED = 30

# sender_info 必填（OMS 文档 createReturnBill 回邮）→ 表格列名（用于任务信息）
# 邮箱按业务约定为非必填，不列入校验
SENDER_REQUIRED_FIELDS: Mapping[str, str] = {
    "sender_name": "收件人",
    "sender_country": "国家代码",
    "sender_phone": "电话",
    "sender_city": "城市",
    "sender_zipcode": "邮编",
    "sender_address1": "街道",
    "sender_address2": "门牌号",
}

PROCESS_MAP: Mapping[str, str] = {
    "是": "1",
    "良品": "1",
    "重新上架": "1",
    "1": "1",
    "否": "3",
    "不良": "3",
    "不良品": "3",
    "3": "3",
    "待检查": "5",
    "5": "5",
    "销毁": "4",
    "4": "4",
}


@dataclass
class SheetRow:
    excel_row: int  # 1-based，含表头 → 数据行从 2 起
    values: Dict[str, str]


@dataclass
class ReturnGroup:
    key: str
    rows: List[SheetRow] = field(default_factory=list)


@dataclass
class RunStats:
    sheet_rows: int = 0
    groups: int = 0
    skipped_done: int = 0
    skipped_empty: int = 0
    skipped_no_shipped: int = 0
    ok: int = 0
    fail: int = 0


@dataclass(frozen=True)
class ShippedOrder:
    order_no: str
    provider_order_no: str
    shop_name_en: str
    city: str = ""
    postal_code: str = ""


def fetch_shipped_orders(order_nos: Sequence[str]) -> Dict[str, ShippedOrder]:
    """按 order_no 批量查 sales_order_shipped；同单多行取首条，并补齐空的 city/postal。"""
    wanted = sorted({clean_cell(x) for x in order_nos if clean_cell(x)})
    if not wanted:
        return {}

    db = get_db_manager()
    conn = db.get_connection()
    cur = conn.cursor(pymysql.cursors.DictCursor)
    mapping: Dict[str, ShippedOrder] = {}
    try:
        for i in range(0, len(wanted), DB_BATCH_SIZE):
            chunk = wanted[i : i + DB_BATCH_SIZE]
            placeholders = ",".join(["%s"] * len(chunk))
            sql = f"""
                SELECT order_no, provider_order_no, shop_name_en, city, postal_code
                FROM `{SHIPPED_TABLE}`
                WHERE order_no IN ({placeholders})
                ORDER BY id ASC
            """
            cur.execute(sql, chunk)
            for row in cur.fetchall():
                order_no = clean_cell(row.get("order_no"))
                if not order_no:
                    continue
                city = clean_cell(row.get("city"))
                postal = clean_cell(row.get("postal_code"))
                prev = mapping.get(order_no)
                if prev is None:
                    mapping[order_no] = ShippedOrder(
                        order_no=order_no,
                        provider_order_no=clean_cell(row.get("provider_order_no")),
                        shop_name_en=clean_cell(row.get("shop_name_en")),
                        city=city,
                        postal_code=postal,
                    )
                    continue
                # 后续行仅补齐空的 city / postal_code
                if (not prev.city and city) or (not prev.postal_code and postal):
                    mapping[order_no] = ShippedOrder(
                        order_no=prev.order_no,
                        provider_order_no=prev.provider_order_no,
                        shop_name_en=prev.shop_name_en,
                        city=prev.city or city,
                        postal_code=prev.postal_code or postal,
                    )
    finally:
        cur.close()
        conn.close()
    return mapping


def _erp_order_no(group: ReturnGroup) -> str:
    return group.rows[0].values.get(COL_ERP, "")


def _apply_shipped_to_group(group: ReturnGroup, shipped: ShippedOrder) -> Dict[str, bool]:
    """把 shipped 字段写进内存行；城市/邮编仅在表格为空时回填。

    返回是否实际回填了城市/邮编（供写回钉钉表）。
    """
    filled_city = False
    filled_zip = False
    for row in group.rows:
        if shipped.provider_order_no:
            row.values[COL_OMS_ORDER] = shipped.provider_order_no
        if shipped.shop_name_en:
            row.values[COL_SHOP] = shipped.shop_name_en
        if not row.values.get(COL_CITY, "") and shipped.city:
            row.values[COL_CITY] = shipped.city
            filled_city = True
        if not row.values.get(COL_ZIP, "") and shipped.postal_code:
            row.values[COL_ZIP] = shipped.postal_code
            filled_zip = True
    return {"city": filled_city, "postal_code": filled_zip}


def _cell(row: Mapping[str, Any], *names: str) -> str:
    for name in names:
        if name in row:
            return clean_cell(row[name])
    return ""


def _sheet_col_index(df: pd.DataFrame, col_name: str) -> int:
    cols = [clean_cell(c) for c in df.columns]
    try:
        return cols.index(clean_cell(col_name))
    except ValueError as exc:
        raise KeyError(f"表格缺少列: [{col_name}]；实际列={list(df.columns)}") from exc


def _group_key(values: Mapping[str, str]) -> str:
    for col in GROUP_KEY_COLS:
        v = values.get(col, "")
        if v:
            return v
    sku = values.get("产品SKU", "")
    return f"SKU:{sku}" if sku else ""


def _process_code(values: Mapping[str, str], default: str) -> str:
    raw = values.get("是否良品上架", "")
    if not raw:
        return default
    mapped = PROCESS_MAP.get(raw) or PROCESS_MAP.get(raw.lower())
    return mapped or default


def dataframe_to_rows(df: pd.DataFrame) -> List[SheetRow]:
    """DataFrame → SheetRow；保留原 index 以定位 Excel 行号。"""
    rows: List[SheetRow] = []
    for idx, series in df.iterrows():
        values = {clean_cell(c): clean_cell(series[c]) for c in df.columns}
        excel_row = int(idx) + 2  # header=1
        rows.append(SheetRow(excel_row=excel_row, values=values))
    return rows


def build_groups(
    rows: Sequence[SheetRow],
    *,
    force: bool = False,
    only_keys: Optional[Sequence[str]] = None,
) -> Tuple[List[ReturnGroup], RunStats]:
    stats = RunStats(sheet_rows=len(rows))
    wanted = {clean_cell(k) for k in (only_keys or []) if clean_cell(k)}
    buckets: Dict[str, ReturnGroup] = {}

    for row in rows:
        key = _group_key(row.values)
        if not key:
            stats.skipped_empty += 1
            continue
        if wanted and key not in wanted:
            continue
        if not force and row.values.get(COL_OMS_RETURN, ""):
            stats.skipped_done += 1
            continue
        if not row.values.get("产品SKU", ""):
            stats.skipped_empty += 1
            continue
        grp = buckets.get(key)
        if grp is None:
            grp = ReturnGroup(key=key)
            buckets[key] = grp
        grp.rows.append(row)

    groups = list(buckets.values())
    stats.groups = len(groups)
    return groups, stats


def _sender_info(values: Mapping[str, str]) -> Dict[str, str]:
    """从表格组装 sender_info；必填项为空则保持空串，由本地校验拦截（不再用 "-" 占位）。"""
    name = values.get("收件人", "") or values.get("Buyer ID", "")
    street = values.get("街道", "")
    door = values.get("门牌号", "")
    city = values.get("城市", "")
    zipcode = values.get("邮编", "")
    country = values.get("国家代码", "").upper()
    email = values.get("邮箱", "") or values.get("Email", "") or values.get("email", "")
    phone = (
        values.get("电话", "")
        or values.get("手机", "")
        or values.get("Phone", "")
        or DEFAULT_SENDER_PHONE
    )

    # address1=街道；address2=门牌号（二者均为文档必填，空则留给校验）
    info: Dict[str, str] = {
        "sender_name": name,
        "sender_country": country,
        "sender_email": email,
        "sender_phone": phone,
        "sender_city": city,
        "sender_zipcode": zipcode,
        "sender_address1": street,
        "sender_address2": door,
    }
    if door:
        info["sender_doorplate"] = door
    if city:
        info["sender_state"] = city
    return info


def _missing_sender_fields(sender: Mapping[str, Any]) -> List[str]:
    missing: List[str] = []
    for key, label in SENDER_REQUIRED_FIELDS.items():
        if not clean_cell(sender.get(key)):
            missing.append(label)
    return missing


def _validate_params_before_api(params: Mapping[str, Any], *, mode: str) -> None:
    """按 OMS createReturnBill 文档校验必填；失败则不调 API，错误写入任务信息。"""
    if mode == "mail":
        top_required = ("reference_no", "warehouse_code", "sm_code", "items")
        missing_top = [k for k in top_required if not params.get(k)]
        if missing_top:
            raise ValueError(f"回邮退件缺少必填: {', '.join(missing_top)}")
        if int(params.get("return_identification") or 0) != 1:
            raise ValueError("回邮退件 return_identification 须为 1")
        sender = params.get("sender_info")
        if not isinstance(sender, Mapping):
            raise ValueError("回邮退件 sender_info 必填")
        missing = _missing_sender_fields(sender)
        if missing:
            raise ValueError(f"寄件人必填项为空: {', '.join(missing)}")
        return

    # standard
    missing_top = [
        k
        for k in ("tracking_no", "warehouse_code", "return_type", "items")
        if not params.get(k)
    ]
    if missing_top:
        raise ValueError(f"标准退件缺少必填: {', '.join(missing_top)}")
    rt = str(params.get("return_type") or "").strip().upper()
    if rt not in ("S", "L", "C"):
        raise ValueError(f"return_type 须为 S/L/C，收到: {params.get('return_type')}")
    if rt == "S" and not params.get("order_code"):
        raise ValueError("return_type=S（买家退件）时 OMS原始订单号/order_code 必填")
    if rt == "C" and not params.get("claim_code"):
        raise ValueError("return_type=C（认领退件）时 认领单号/claim_code 必填")
    items = params.get("items") or []
    for i, item in enumerate(items):
        if not isinstance(item, Mapping):
            raise ValueError(f"items[{i}] 非法")
        for key in ("product_sku", "quantity", "process"):
            if item.get(key) in (None, ""):
                raise ValueError(f"items[{i}].{key} 必填")


def _format_api_error(exc: HyOmsError) -> str:
    """优先取 OMS 响应里的业务错误文案，供回写「任务信息」。"""
    raw = getattr(exc, "raw", None)
    if isinstance(raw, dict):
        err = raw.get("Error")
        if isinstance(err, dict):
            for key in ("errMessage", "message", "errMsg"):
                text = clean_cell(err.get(key))
                if text:
                    return text
        for key in ("message", "errMessage", "ask"):
            text = clean_cell(raw.get(key))
            if text:
                return text
    return str(exc)


def _task_msg(text: str) -> str:
    return text[:TASK_MSG_MAX]


def _items_from_group(group: ReturnGroup, default_process: str) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    for row in group.rows:
        sku = row.values.get("产品SKU", "")
        qty_raw = row.values.get("退件数量", "") or "1"
        try:
            qty = int(float(qty_raw))
        except ValueError as exc:
            raise ValueError(f"退件数量非法 key={group.key} sku={sku}: {qty_raw}") from exc
        if qty <= 0:
            raise ValueError(f"退件数量须 > 0 key={group.key} sku={sku}")
        process = _process_code(row.values, default_process)
        item: Dict[str, Any] = {
            "product_sku": sku,
            "quantity": qty,
            "process": str(process),
        }
        note = row.values.get("备注", "")
        if note:
            item["note"] = note[:255]
        items.append(item)
    return items


def build_mail_params(
    group: ReturnGroup,
    *,
    warehouse_code: str,
    sm_code: str,
    verify: Optional[int],
    default_process: str,
    return_type: str = DEFAULT_RETURN_TYPE,
) -> Dict[str, Any]:
    """组装回邮参数；必填校验见 ``_validate_params_before_api``。"""
    head = group.rows[0].values
    return build_params(
        warehouse_code=warehouse_code,
        items=_items_from_group(group, default_process),
        reference_no=group.key[:32],
        return_identification=1,
        return_type=return_type or DEFAULT_RETURN_TYPE,
        sm_code=sm_code,
        verify=verify,
        return_desc=(head.get("退件原因", "") or None),
        operation_desc=(head.get("备注", "") or None),
        seller_store=(head.get("店铺名", "") or head.get("账号", "") or None),
        sender_info=_sender_info(head),
        validate=False,
    )


def build_standard_params(
    group: ReturnGroup,
    *,
    warehouse_code: str,
    return_type: str,
    verify: Optional[int],
    default_process: str,
) -> Dict[str, Any]:
    head = group.rows[0].values
    tracking = head.get(COL_LABEL_TRACKING, "")
    order_code = head.get(COL_OMS_ORDER, "") or head.get("ERP订单号", "")
    claim_code = head.get("认领单号", "")
    return build_params(
        warehouse_code=warehouse_code,
        items=_items_from_group(group, default_process),
        tracking_no=tracking or None,
        return_type=return_type,
        verify=verify,
        reference_no=group.key[:32],
        order_code=order_code or None,
        claim_code=claim_code or None,
        return_desc=(head.get("退件原因", "") or None),
        operation_desc=(head.get("备注", "") or None),
        buyer_name=(head.get("收件人", "") or head.get("Buyer ID", "") or None),
        seller_store=(head.get("店铺名", "") or head.get("账号", "") or None),
        validate=False,
    )


def _cell_write_value(val: Any) -> str:
    """钉钉 ranges 写入要求字符串（否则报 String is mandatory）。"""
    if val is None:
        return ""
    if isinstance(val, float) and pd.isna(val):
        return ""
    if isinstance(val, bool):
        return "1" if val else "0"
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    if isinstance(val, (int, float)):
        return str(val)
    return str(val)


def _write_back(
    wb: Workbook,
    sheet: str,
    df: pd.DataFrame,
    updates_by_col: Mapping[str, Sequence[Tuple[int, Any]]],
) -> int:
    written = 0
    for col_name, updates in updates_by_col.items():
        if not updates:
            continue
        if col_name not in [clean_cell(c) for c in df.columns]:
            print(f"[WARN] 回写跳过，表格无列: [{col_name}]", file=sys.stderr)
            continue
        col_idx = _sheet_col_index(df, col_name)
        safe_updates = [(row, _cell_write_value(val)) for row, val in updates]
        try:
            written += wb.write_column_updates(sheet, col_idx, safe_updates)
        except DingDiskError as exc:
            print(f"[FAIL] 回写列[{col_name}] 失败: {exc}", file=sys.stderr)
            raise
    return written


def process_groups(
    wb: Workbook,
    sheet: str,
    df: pd.DataFrame,
    groups: Sequence[ReturnGroup],
    stats: RunStats,
    *,
    mode: str,
    warehouse_code: str,
    sm_code: str,
    return_type: str,
    verify: Optional[int],
    default_process: str,
    dry_run: bool,
    write_back: bool,
) -> int:
    """处理分组；返回进程退出码（有失败则为 1）。"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    back: Dict[str, List[Tuple[int, Any]]] = {
        COL_OMS_ORDER: [],
        COL_SHOP: [],
        COL_CITY: [],
        COL_ZIP: [],
        COL_OMS_RETURN: [],
        COL_PROGRESS: [],
        COL_TASK: [],
    }

    erp_nos = [_erp_order_no(g) for g in groups]
    shipped_map = fetch_shipped_orders(erp_nos)
    print(
        f"[DB] sales_order_shipped 查询 ERP={len({x for x in erp_nos if x})} "
        f"命中={len(shipped_map)}"
    )

    for i, group in enumerate(groups, 1):
        label = f"[{i}/{len(groups)}] key={group.key} rows={len(group.rows)}"
        erp = _erp_order_no(group)
        if not erp:
            stats.skipped_no_shipped += 1
            msg = "缺少 ERP订单号，无法匹配 sales_order_shipped.order_no"
            print(f"[SKIP] {label} {msg}", file=sys.stderr)
            for row in group.rows:
                back[COL_TASK].append((row.excel_row, _task_msg(msg)))
            continue

        shipped = shipped_map.get(erp)
        if shipped is None:
            stats.skipped_no_shipped += 1
            msg = f"sales_order_shipped 未找到原单 order_no={erp}"
            print(f"[SKIP] {label} {msg}", file=sys.stderr)
            for row in group.rows:
                back[COL_TASK].append((row.excel_row, _task_msg(msg)))
            continue

        filled = _apply_shipped_to_group(group, shipped)
        for row in group.rows:
            if shipped.provider_order_no:
                back[COL_OMS_ORDER].append((row.excel_row, shipped.provider_order_no))
            if shipped.shop_name_en:
                back[COL_SHOP].append((row.excel_row, shipped.shop_name_en))
            if filled["city"]:
                back[COL_CITY].append((row.excel_row, shipped.city))
            if filled["postal_code"]:
                back[COL_ZIP].append((row.excel_row, shipped.postal_code))

        try:
            if mode == "mail":
                params = build_mail_params(
                    group,
                    warehouse_code=warehouse_code,
                    sm_code=sm_code,
                    verify=verify,
                    default_process=default_process,
                    return_type=return_type,
                )
            else:
                params = build_standard_params(
                    group,
                    warehouse_code=warehouse_code,
                    return_type=return_type,
                    verify=verify,
                    default_process=default_process,
                )
            _validate_params_before_api(params, mode=mode)
        except ValueError as exc:
            stats.fail += 1
            msg = str(exc)
            print(f"[FAIL] {label} 参数: {msg}", file=sys.stderr)
            for row in group.rows:
                back[COL_TASK].append((row.excel_row, _task_msg(msg)))
            continue

        if dry_run:
            print(f"[DRY-RUN] {label} oms={shipped.provider_order_no} shop={shipped.shop_name_en}")
            print(json.dumps(params, ensure_ascii=False, indent=2))
            stats.ok += 1
            continue

        try:
            result = HyOmsClient.from_config().call("createReturnBill", params)
        except HyOmsError as exc:
            stats.fail += 1
            msg = _format_api_error(exc)
            print(f"[FAIL] {label} API: {msg}", file=sys.stderr)
            raw = getattr(exc, "raw", None)
            if raw is not None:
                print(json.dumps(raw, ensure_ascii=False, indent=2)[:2000], file=sys.stderr)
            for row in group.rows:
                back[COL_TASK].append((row.excel_row, _task_msg(msg)))
            continue

        return_code = clean_cell(result.get("return_code"))
        task = f"{now} 创建成功"
        print(f"[OK] {label} return_code={return_code}")
        stats.ok += 1
        for row in group.rows:
            if return_code:
                back[COL_OMS_RETURN].append((row.excel_row, return_code))
            back[COL_PROGRESS].append((row.excel_row, PROGRESS_CREATED))
            back[COL_TASK].append((row.excel_row, _task_msg(task)))

    # dry-run 不改表；正式运行回写 OMS 原单/店铺名/进度/任务信息/退件单号
    if write_back and not dry_run:
        n = _write_back(wb, sheet, df, back)
        print(f"[WRITE-BACK] cells={n}")

    print(
        f"[DONE] sheet_rows={stats.sheet_rows} groups={stats.groups} "
        f"skipped_done={stats.skipped_done} skipped_empty={stats.skipped_empty} "
        f"skipped_no_shipped={stats.skipped_no_shipped} "
        f"ok={stats.ok} fail={stats.fail}"
    )
    return 1 if (stats.fail or stats.skipped_no_shipped) else 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="读取钉钉退件表格并调用鸿羽 OMS createReturnBill"
    )
    parser.add_argument(
        "--workbook-id",
        default=WORKBOOK_ID,
        help=f"钉钉表格文档 ID；默认 {WORKBOOK_ID}",
    )
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help=f"工作表名；默认 {DEFAULT_SHEET}")
    parser.add_argument("--list-sheets", action="store_true", help="仅列出工作表")
    parser.add_argument(
        "--mode",
        choices=("mail", "standard"),
        default="mail",
        help="mail=回邮退件（默认）；standard=标准退件",
    )
    parser.add_argument(
        "--warehouse-code",
        default=DEFAULT_WAREHOUSE,
        help=f"仓库编码；默认 {DEFAULT_WAREHOUSE}",
    )
    parser.add_argument(
        "--sm-code",
        default=DEFAULT_SM_CODE,
        help=f"回邮物流产品代码；默认 {DEFAULT_SM_CODE}",
    )
    parser.add_argument(
        "--return-type",
        default=DEFAULT_RETURN_TYPE,
        choices=("S", "L", "C"),
        help=f"退件类型；默认 {DEFAULT_RETURN_TYPE}（买家退件）",
    )
    parser.add_argument(
        "--verify",
        type=int,
        choices=(0, 1),
        default=1,
        help="1确认审核 / 0草稿；默认 1",
    )
    parser.add_argument(
        "--process",
        default=DEFAULT_PROCESS,
        help=f"默认处理方式；默认 {DEFAULT_PROCESS}=重新上架",
    )
    parser.add_argument(
        "--key",
        action="append",
        dest="keys",
        default=None,
        help="仅处理指定 ERP/平台订单号（可重复）",
    )
    parser.add_argument("--limit", type=int, default=None, help="最多处理 N 组")
    parser.add_argument("--force", action="store_true", help="已有 OMS退件订单号 也重新创建")
    parser.add_argument("--dry-run", action="store_true", help="只打印 paramsJson，不调用 API")
    parser.add_argument(
        "--no-write-back",
        action="store_true",
        help="不回写钉钉表格（进度/退件单号）",
    )
    parser.add_argument(
        "--preview",
        type=int,
        default=None,
        help="预览表格前 N 行后退出（不创建）",
    )
    parser.add_argument("--raw", action="store_true", help="--list-sheets 时缩进 JSON")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)

    try:
        workbook_id = (args.workbook_id or WORKBOOK_ID).strip()
        if not workbook_id:
            print("[FAIL] 未指定 workbookId", file=sys.stderr)
            return 2

        wb = Workbook(workbook_id)
        sheets = wb.list_sheets()

        if args.list_sheets:
            payload = {"workbookId": workbook_id, "sheets": sheets}
            print(json.dumps(payload, ensure_ascii=False, indent=2 if args.raw else None))
            return 0

        if not sheets:
            print(f"[WARN] 表格无工作表: {workbook_id}", file=sys.stderr)
            return 1

        sheet = args.sheet or DEFAULT_SHEET
        df = wb.read_sheet(sheet)
        print(f"[READ] workbook={workbook_id} sheet={sheet} rows={len(df)} cols={len(df.columns)}")

        if args.preview is not None:
            n = args.preview if args.preview > 0 else len(df)
            preview = df.head(n).fillna("").astype(str)
            print(preview.to_string())
            return 0

        rows = dataframe_to_rows(df)
        groups, stats = build_groups(rows, force=bool(args.force), only_keys=args.keys)
        if args.limit is not None and args.limit >= 0:
            groups = groups[: args.limit]
            stats.groups = len(groups)

        if not groups:
            print(
                f"[DONE] 无可处理分组 sheet_rows={stats.sheet_rows} "
                f"skipped_done={stats.skipped_done} skipped_empty={stats.skipped_empty}"
            )
            return 0

        return process_groups(
            wb,
            sheet,
            df,
            groups,
            stats,
            mode=args.mode,
            warehouse_code=args.warehouse_code,
            sm_code=args.sm_code,
            return_type=args.return_type,
            verify=args.verify,
            default_process=str(args.process),
            dry_run=bool(args.dry_run),
            write_back=not args.no_write_back,
        )
    except (DingDiskError, HyOmsError, ValueError, KeyError) as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
