"""从钉钉在线表格读取退件申请，调用鸿羽 OMS ``createReturnBill``。

表格读写委托 ``api.ding_disk.workbook.Workbook``；创建退件委托
``api.hy_oms.request.create_return``。

文档 ID（workbookId / nodeId）::
    EpGBa2Lm8aDaZ57lTwEMk9boJgN7R35y

默认模式为 **回邮退件**（``return_identification=1``），因表格含寄件地址与标签列。
同一 ``ERP订单号``（空则 ``平台订单号``）多行会合并为一条退件单的 ``items``。
仅当表格 ``确认提交`` 已打勾（√）时才调用 createReturnBill。

推送 API 前先查 ``sales_order_shipped``（``ERP订单号`` = ``order_no``）：
命中则回写 ``OMS原始订单号``=``provider_order_no``、
``店铺名称``=``shop_name_en``（有值即覆盖）、
``平台订单号``=``ref_no``（有值即覆盖）；
表格 ``国家代码``/``城市``/``邮编`` 为空时用 ``country``/``city``/``postal_code`` 回填。
DB 未命中时，再调易仓 WMS ``getOrders``（``code``=ERP 订单号）补查，
同样回填 ``OMS原始订单号`` / ``店铺名称`` / ``平台订单号``（``reference_no``）等；
仍无结果则只回写 ``任务信息``，不调 API。

表格列 → API（回邮）::
    ERP订单号 / 平台订单号  →  reference_no
    产品SKU + 退件数量      →  items[].product_sku / quantity
        产品SKU 含 ``+`` 时按多产品解析：``SKU*数量+SKU*数量``，
        先 ``+`` 再 ``*``；无 ``*`` 数量时回退该行「退件数量」
        非 ``-NW`` 结尾的 SKU 会自动补 ``-NW``；创建退件前
        用 getProductList 查 ``SKU-NW``，不存在则按原 SKU 资料
        createProduct 新建后再推 createReturnBill
    退件原因                →  return_desc
    （配置）operation_desc  →  operation_desc（returned_config / --operation-desc）
    收件人/国家/街道/门牌号/城市/邮编 → sender_info
    店铺名称                →  seller_store（优先用 shipped.shop_name_en）
    可选列 邮箱/电话/手机   →  sender_email（非必填）/ sender_phone（空则默认 0000000000）
    是否良品上架            →  items[].process（是→1，否→3，默认 1）
    备注                    →  items[].note（不再映射 operation_desc）

推送前按 OMS 文档本地校验寄件人必填（邮编/电话/街道/门牌号等；邮箱非必填）；
缺项不调 API，错误写入任务信息。注意：空值不得用 ``-`` 占位，否则 OMS 会当成已填而放行。

成功/失败回写::
    OMS原始订单号 / 店铺名称 / 平台订单号 ← sales_order_shipped 或易仓 getOrders
    国家代码 / 城市 / 邮编 ← country / city / postal_code（表格为空时回填）
    OMS退件订单号 ← return_code（成功时）
    进度          ← 0-100 数值；创建成功写 30
    任务信息      ← 成功摘要，或未找到原单 / API 错误原文

用法（项目根目录）::

    python app/hongyu/returned/create_return.py --dry-run
    python app/hongyu/returned/create_return.py --list-sheets
    python app/hongyu/returned/create_return.py --limit 1 --preview 5
    python app/hongyu/returned/create_return.py --sm-code DEGLS-RMA
    # 未传 --sm-code 时：国家代码 DE → DEGLS-RMA，其它 → DEDHL-RMA
    python app/hongyu/returned/create_return.py --mode standard --return-type S
    python app/hongyu/returned/create_return.py --verify 0

Windows 可执行文件（见 packaging/build_returned.ps1）::

    .\\dist\\returned\\create_return.exe --dry-run
    .\\dist\\returned\\run_task.exe          # 日常一键：创建→更新→下载→刷新
    # 配置放在 dist\\config\\（多模块共享），不是 exe 同级
"""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass, field, replace
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple, Union

import pandas as pd
import pymysql.cursors

_RETURNED_DIR = Path(__file__).resolve().parent
if str(_RETURNED_DIR) not in sys.path:
    sys.path.insert(0, str(_RETURNED_DIR))

import runtime_config as _returned_rt  # noqa: E402

_PROJECT_ROOT, _CFG = _returned_rt.init_script(__file__)

from api.ding_disk.exceptions import DingDiskError  # noqa: E402
from api.ding_disk.workbook import Workbook, clean_cell  # noqa: E402

from sheet_utils import (  # noqa: E402
    SheetRow,
    cell_write_value,
    dataframe_to_rows,
    format_api_error,
    has_col,
    sheet_col_index,
    task_msg,
)

from api.hy_oms import HyOmsClient  # noqa: E402
from api.hy_oms.exceptions import HyOmsError  # noqa: E402
from api.hy_oms.request.create_product import create_product  # noqa: E402
from api.hy_oms.request.create_return import build_params  # noqa: E402
from api.hy_oms.request.get_product import get_product_list  # noqa: E402
from database.db_connection import get_db_manager  # noqa: E402

WORKBOOK_ID = _CFG.workbook_id
DEFAULT_SHEET = _CFG.register_sheet
DEFAULT_WAREHOUSE = _CFG.warehouse_code
DEFAULT_SM_CODE_DE = _CFG.sm_code_de
DEFAULT_SM_CODE_OTHER = _CFG.sm_code_other
DEFAULT_SM_CODE = DEFAULT_SM_CODE_DE  # 兼容旧用法；实际按国家代码自动选择
DEFAULT_PROCESS = "1"
DEFAULT_RETURN_TYPE = "S"
DEFAULT_SENDER_PHONE = "0000000000"
DEFAULT_OPERATION_DESC = _CFG.operation_desc or "买家退件，检查换标上架"
SHIPPED_TABLE = "sales_order_shipped"
DB_BATCH_SIZE = 500
NW_SUFFIX = "-NW"

GROUP_KEY_COLS = ("ERP订单号", "平台订单号")

COL_ERP = "ERP订单号"
COL_PLATFORM = "平台订单号"
COL_SHOP = "店铺名称"
COL_COUNTRY = "国家代码"
COL_CITY = "城市"
COL_ZIP = "邮编"
COL_CONFIRM = "确认提交"
COL_OMS_RETURN = "OMS退件订单号"
COL_PROGRESS = "进度"
COL_TASK = "任务信息"
COL_OMS_ORDER = "OMS原始订单号"
COL_LABEL_TRACKING = "标签跟踪号"
PROGRESS_CREATED = 10

CONFIRM_TRUE = {"1", "true", "yes", "y", "是", "√", "✓", "checked", "勾选"}

# sender_info 必填（OMS 文档）→ 表格列名；邮箱非必填
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

WRITE_BACK_COLS = (
    COL_OMS_ORDER,
    COL_PLATFORM,
    COL_SHOP,
    COL_COUNTRY,
    COL_CITY,
    COL_ZIP,
    COL_OMS_RETURN,
    COL_PROGRESS,
    COL_TASK,
)


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
    skipped_unconfirmed: int = 0
    skipped_no_shipped: int = 0
    ok: int = 0
    fail: int = 0


@dataclass(frozen=True)
class ShippedOrder:
    order_no: str
    provider_order_no: str
    shop_name_en: str
    ref_no: str = ""
    country: str = ""
    city: str = ""
    postal_code: str = ""


@dataclass
class FilledFields:
    """表格空值时从 shipped 回填的字段标记。"""

    country: bool = False
    city: bool = False
    postal_code: bool = False


@dataclass
class CreateOptions:
    mode: str = "mail"
    warehouse_code: str = DEFAULT_WAREHOUSE
    sm_code: Optional[str] = None
    return_type: str = DEFAULT_RETURN_TYPE
    verify: Optional[int] = 1
    default_process: str = DEFAULT_PROCESS
    operation_desc: str = DEFAULT_OPERATION_DESC
    dry_run: bool = False
    write_back: bool = True


@dataclass
class WriteBackBuffer:
    """收集钉钉回写：列名 → (excel_row, value) 列表。"""

    updates: Dict[str, List[Tuple[int, Any]]] = field(default_factory=dict)

    def __post_init__(self) -> None:
        if not self.updates:
            self.updates = {col: [] for col in WRITE_BACK_COLS}

    def task_for_group(self, group: ReturnGroup, msg: str) -> None:
        text = task_msg(msg)
        for row in group.rows:
            self.updates[COL_TASK].append((row.excel_row, text))

    def apply_shipped(
        self,
        group: ReturnGroup,
        shipped: ShippedOrder,
        filled: FilledFields,
    ) -> None:
        for row in group.rows:
            if shipped.provider_order_no:
                self.updates[COL_OMS_ORDER].append(
                    (row.excel_row, shipped.provider_order_no)
                )
            if shipped.ref_no:
                self.updates[COL_PLATFORM].append((row.excel_row, shipped.ref_no))
            if shipped.shop_name_en:
                self.updates[COL_SHOP].append((row.excel_row, shipped.shop_name_en))
            if filled.country:
                self.updates[COL_COUNTRY].append((row.excel_row, shipped.country))
            if filled.city:
                self.updates[COL_CITY].append((row.excel_row, shipped.city))
            if filled.postal_code:
                self.updates[COL_ZIP].append((row.excel_row, shipped.postal_code))

    def mark_created(
        self,
        group: ReturnGroup,
        *,
        return_code: str,
        progress: Any,
        task: str,
    ) -> None:
        text = task_msg(task)
        for row in group.rows:
            if return_code:
                self.updates[COL_OMS_RETURN].append((row.excel_row, return_code))
            self.updates[COL_PROGRESS].append((row.excel_row, progress))
            self.updates[COL_TASK].append((row.excel_row, text))

    def as_updates(self) -> Mapping[str, Sequence[Tuple[int, Any]]]:
        return self.updates


def fetch_shipped_orders(order_nos: Sequence[str]) -> Dict[str, ShippedOrder]:
    """按 order_no 批量查 sales_order_shipped；同单多行取首条，并补齐空的 country/city/postal。"""
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
                SELECT order_no, provider_order_no, shop_name_en, ref_no,
                       country, city, postal_code
                FROM `{SHIPPED_TABLE}`
                WHERE order_no IN ({placeholders})
                ORDER BY id ASC
            """
            cur.execute(sql, chunk)
            for row in cur.fetchall():
                order_no = clean_cell(row.get("order_no"))
                if not order_no:
                    continue
                country = clean_cell(row.get("country")).upper()
                city = clean_cell(row.get("city"))
                postal = clean_cell(row.get("postal_code"))
                ref_no = clean_cell(row.get("ref_no"))
                prev = mapping.get(order_no)
                if prev is None:
                    mapping[order_no] = ShippedOrder(
                        order_no=order_no,
                        provider_order_no=clean_cell(row.get("provider_order_no")),
                        shop_name_en=clean_cell(row.get("shop_name_en")),
                        ref_no=ref_no,
                        country=country,
                        city=city,
                        postal_code=postal,
                    )
                    continue
                if (
                    (not prev.ref_no and ref_no)
                    or (not prev.country and country)
                    or (not prev.city and city)
                    or (not prev.postal_code and postal)
                ):
                    mapping[order_no] = replace(
                        prev,
                        ref_no=prev.ref_no or ref_no,
                        country=prev.country or country,
                        city=prev.city or city,
                        postal_code=prev.postal_code or postal,
                    )
    finally:
        cur.close()
        conn.close()
    return mapping


def _eccang_order_matches_code(order: Mapping[str, Any], lookup_code: str) -> bool:
    """``getOrders`` 返回行是否匹配查询单号（仓库单号/参考号/跟踪号）。"""
    code = clean_cell(lookup_code)
    if not code:
        return False
    for key in ("warehouse_order_code", "reference_no", "tracking_number"):
        if clean_cell(order.get(key)) == code:
            return True
    refs = order.get("platform_ref_no")
    if isinstance(refs, list):
        return any(clean_cell(x) == code for x in refs)
    return False


def _pick_best_eccang_order(
    items: Sequence[Mapping[str, Any]],
    lookup_code: str,
) -> Optional[Mapping[str, Any]]:
    """同一参考号可能多条；优先已出库(8)且发货时间最新。"""
    matches = [
        row
        for row in items
        if isinstance(row, Mapping) and _eccang_order_matches_code(row, lookup_code)
    ]
    if not matches:
        return None
    if len(matches) == 1:
        return matches[0]

    def _rank(row: Mapping[str, Any]) -> tuple[int, str]:
        status = str(row.get("order_status") or "").strip()
        shipped = 1 if status == "8" else 0
        ship_time = str(row.get("ship_time") or "")
        return shipped, ship_time

    return max(matches, key=_rank)


def _eccang_ref_no(order: Mapping[str, Any]) -> str:
    """易仓订单 → 平台参考号（优先 ``reference_no``，否则 ``platform_ref_no[0]``）。"""
    ref = clean_cell(order.get("reference_no"))
    if ref:
        return ref
    refs = order.get("platform_ref_no")
    if isinstance(refs, list):
        for item in refs:
            text = clean_cell(item)
            if text:
                return text
    return ""


def _shipped_from_eccang_order(order_no: str, order: Mapping[str, Any]) -> ShippedOrder:
    """易仓 ``getOrders`` 单行 → ``ShippedOrder``。"""
    addr = order.get("address")
    if not isinstance(addr, Mapping):
        addr = {}
    shop = clean_cell(order.get("seller_id")) or clean_cell(order.get("platform_user_name"))
    return ShippedOrder(
        order_no=order_no,
        provider_order_no=clean_cell(order.get("warehouse_order_code")),
        shop_name_en=shop,
        ref_no=_eccang_ref_no(order),
        country=clean_cell(addr.get("oab_country")).upper(),
        city=clean_cell(addr.get("oab_city")),
        postal_code=clean_cell(addr.get("oab_postcode")),
    )


def fetch_eccang_orders_fallback(order_nos: Sequence[str]) -> Dict[str, ShippedOrder]:
    """DB 未命中时，按 ERP 订单号调易仓 ``getOrders``（``code`` 精确查）。"""
    wanted = sorted({clean_cell(x) for x in order_nos if clean_cell(x)})
    if not wanted:
        return {}

    from api.eccang.exceptions import EccangApiError, EccangConfigError
    from api.eccang.request.getOrders import _extract_page, get_orders

    mapping: Dict[str, ShippedOrder] = {}
    fail_n = 0
    for order_no in wanted:
        try:
            result = get_orders(page=1, page_size=20, code=[order_no])
        except EccangConfigError as exc:
            fail_n += 1
            print(f"[WARN] 易仓 getOrders 配置错误: {exc}", file=sys.stderr)
            break
        except EccangApiError as exc:
            fail_n += 1
            print(
                f"[WARN] 易仓 getOrders 失败 order_no={order_no}: {format_api_error(exc)}",
                file=sys.stderr,
            )
            continue
        except Exception as exc:
            fail_n += 1
            print(
                f"[WARN] 易仓 getOrders 异常 order_no={order_no}: {exc}",
                file=sys.stderr,
            )
            continue

        items, *_ = _extract_page(result)
        picked = _pick_best_eccang_order(items, order_no)
        if picked is None:
            continue
        mapping[order_no] = _shipped_from_eccang_order(order_no, picked)

    if mapping:
        print(
            f"[ECCANG] getOrders 命中 {len(mapping)}/{len(wanted)} "
            f"(warehouse_order_code→OMS原始订单号)"
        )
    elif wanted and fail_n == 0:
        print(f"[ECCANG] getOrders 未命中 {len(wanted)} 个 ERP 订单号")
    return mapping


def resolve_shipped_orders(order_nos: Sequence[str]) -> Dict[str, ShippedOrder]:
    """先查 ``sales_order_shipped``，未命中再易仓 ``getOrders`` 补查。"""
    mapping = fetch_shipped_orders(order_nos)
    db_hit = len(mapping)
    missing = sorted(
        code
        for code in {clean_cell(x) for x in order_nos if clean_cell(x)}
        if code not in mapping
    )
    eccang_hit = 0
    if missing:
        eccang_map = fetch_eccang_orders_fallback(missing)
        eccang_hit = len(eccang_map)
        mapping.update(eccang_map)
    erp_cnt = len({clean_cell(x) for x in order_nos if clean_cell(x)})
    print(
        f"[SHIP] ERP={erp_cnt} db_hit={db_hit} eccang_hit={eccang_hit} "
        f"total_hit={len(mapping)}"
    )
    return mapping


def _erp_order_no(group: ReturnGroup) -> str:
    return group.rows[0].values.get(COL_ERP, "")


def resolve_sm_code(
    country: str,
    override: Optional[str] = None,
) -> str:
    """回邮物流产品：显式 override 优先；否则 DE→DEGLS-RMA，其它→DEDHL-RMA。"""
    text = clean_cell(override)
    if text:
        return text
    if clean_cell(country).upper() == "DE":
        return DEFAULT_SM_CODE_DE
    return DEFAULT_SM_CODE_OTHER


def _apply_shipped_to_group(group: ReturnGroup, shipped: ShippedOrder) -> FilledFields:
    """把 shipped 写进内存行；OMS/店铺/平台订单号有值即覆盖，国家/城市/邮编仅空时回填。"""
    filled = FilledFields()
    for row in group.rows:
        if shipped.provider_order_no:
            row.values[COL_OMS_ORDER] = shipped.provider_order_no
        if shipped.ref_no:
            row.values[COL_PLATFORM] = shipped.ref_no
        if shipped.shop_name_en:
            row.values[COL_SHOP] = shipped.shop_name_en
        if not row.values.get(COL_COUNTRY, "") and shipped.country:
            row.values[COL_COUNTRY] = shipped.country
            filled.country = True
        if not row.values.get(COL_CITY, "") and shipped.city:
            row.values[COL_CITY] = shipped.city
            filled.city = True
        if not row.values.get(COL_ZIP, "") and shipped.postal_code:
            row.values[COL_ZIP] = shipped.postal_code
            filled.postal_code = True
    return filled


def _is_confirmed(val: Any) -> bool:
    """「确认提交」是否已打勾。"""
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)) and not (isinstance(val, float) and pd.isna(val)):
        return int(val) == 1
    text = clean_cell(val).lower()
    return text in CONFIRM_TRUE


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


def _seller_store(values: Mapping[str, str]) -> Optional[str]:
    return (
        values.get(COL_SHOP, "")
        or values.get("店铺名", "")
        or values.get("账号", "")
        or None
    )


def build_groups(
    rows: Sequence[SheetRow],
    *,
    force: bool = False,
    only_keys: Optional[Sequence[str]] = None,
) -> Tuple[List[ReturnGroup], RunStats]:
    """仅处理「确认提交」已勾选、且尚未创建成功的行。"""
    stats = RunStats(sheet_rows=len(rows))
    wanted = {clean_cell(k) for k in (only_keys or []) if clean_cell(k)}
    buckets: Dict[str, ReturnGroup] = {}

    for row in rows:
        if not _is_confirmed(row.values.get(COL_CONFIRM, "")):
            stats.skipped_unconfirmed += 1
            continue
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
    """从表格组装 sender_info；必填项为空则保持空串，由本地校验拦截。"""
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
    """按 OMS createReturnBill 文档校验必填；失败则不调 API。"""
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


def _parse_sku_products(sku_raw: str, default_qty: int) -> List[Tuple[str, int]]:
    """解析「产品SKU」列。

    含 ``+`` 时视为多产品：``SKU*数量+SKU*数量``，先按 ``+`` 再按 ``*``
   （``*`` 取最后一段为数量，便于 SKU 本身含 ``*``）。
    片段无 ``*`` 或数量为空时用 ``default_qty``（行内退件数量）。
    不含 ``+`` 时整段作为单个 SKU，数量用 ``default_qty``。
    """
    text = clean_cell(sku_raw)
    if not text:
        return []
    if "+" not in text:
        return [(text, default_qty)]

    products: List[Tuple[str, int]] = []
    for part in text.split("+"):
        part = clean_cell(part)
        if not part:
            continue
        if "*" in part:
            sku, qty_part = part.rsplit("*", 1)
            sku = clean_cell(sku)
            qty_part = clean_cell(qty_part)
            if not sku:
                raise ValueError(f"产品SKU片段缺少SKU: {part}")
            if qty_part:
                try:
                    qty = int(float(qty_part))
                except ValueError as exc:
                    raise ValueError(f"产品SKU数量非法: {part}") from exc
            else:
                qty = default_qty
        else:
            sku = part
            qty = default_qty
        if qty <= 0:
            raise ValueError(f"产品数量须 > 0: {part}")
        products.append((sku, qty))
    return products


def _with_nw_suffix(sku: str) -> str:
    """非 ``-NW`` 结尾则追加后缀；已是 ``-NW`` 则原样返回。"""
    text = clean_cell(sku)
    if not text:
        return text
    if text.endswith(NW_SUFFIX):
        return text
    return f"{text}{NW_SUFFIX}"


def _nw_base_sku(nw_sku: str) -> str:
    """``SKU-NW`` → ``SKU``；否则原样。"""
    text = clean_cell(nw_sku)
    if text.endswith(NW_SUFFIX):
        return text[: -len(NW_SUFFIX)]
    return text


def _as_optional_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _as_optional_int(value: Any) -> Optional[int]:
    if value is None or value == "":
        return None
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


def _fetch_product_exact(sku: str) -> Optional[Dict[str, Any]]:
    """``getProductList`` 按 product_sku 精确查一条；未命中返回 None。"""
    target = clean_cell(sku)
    if not target:
        return None
    result = get_product_list(product_sku=target, page=1, page_size=10)
    rows = result.get("data") or []
    if isinstance(rows, dict):
        rows = [rows]
    if not isinstance(rows, list):
        return None
    for row in rows:
        if isinstance(row, Mapping) and clean_cell(row.get("product_sku")) == target:
            return dict(row)
    return None


def _create_params_from_product(
    source: Mapping[str, Any],
    *,
    nw_sku: str,
) -> Dict[str, Any]:
    """按原 SKU 资料组装 createProduct 参数（目标 SKU 为 ``nw_sku``）。"""
    title = clean_cell(source.get("product_title")) or clean_cell(
        source.get("product_title_en")
    )
    declared_name = clean_cell(source.get("product_declared_name")) or title
    weight = _as_optional_float(source.get("product_weight"))
    length = _as_optional_float(source.get("product_length"))
    width = _as_optional_float(source.get("product_width"))
    height = _as_optional_float(source.get("product_height"))
    declared_value = _as_optional_float(source.get("product_declared_value"))

    missing: List[str] = []
    if not title:
        missing.append("product_title")
    if weight is None:
        missing.append("product_weight")
    if length is None:
        missing.append("product_length")
    if width is None:
        missing.append("product_width")
    if height is None:
        missing.append("product_height")
    if declared_value is None:
        missing.append("product_declared_value")
    if not declared_name:
        missing.append("product_declared_name")
    if missing:
        base = clean_cell(source.get("product_sku")) or _nw_base_sku(nw_sku)
        raise ValueError(
            f"原SKU={base} 资料缺少 createProduct 必填: {', '.join(missing)}"
        )

    ean = clean_cell(source.get("EAN")) or clean_cell(source.get("ean")) or None
    ncm = clean_cell(source.get("NCM")) or clean_cell(source.get("ncm")) or None
    cest = clean_cell(source.get("CEST")) or clean_cell(source.get("cest")) or None
    desc = clean_cell(source.get("product_description")) or clean_cell(
        source.get("product_desc")
    )

    kwargs: Dict[str, Any] = {
        "product_sku": nw_sku,
        "product_title": title,
        "product_weight": weight,
        "product_length": length,
        "product_width": width,
        "product_height": height,
        "product_declared_value": declared_value,
        "product_declared_name": declared_name,
        "verify": 1,
    }
    optional = {
        "reference_no": clean_cell(source.get("reference_no")) or None,
        "product_title_en": clean_cell(source.get("product_title_en")) or None,
        "contain_battery": _as_optional_int(source.get("contain_battery")),
        "product_declared_name_zh": clean_cell(source.get("product_declared_name_zh"))
        or None,
        "hs_code": clean_cell(source.get("hs_code")) or None,
        "cat_lang": clean_cell(source.get("cat_lang")) or None,
        "product_brand": clean_cell(source.get("product_brand")) or None,
        "product_model": clean_cell(source.get("product_model")) or None,
        "product_origin": clean_cell(source.get("product_origin")) or None,
        "product_material": clean_cell(source.get("product_material")) or None,
        "product_desc_url": clean_cell(source.get("product_desc_url")) or None,
        "cat_id_level0": _as_optional_int(source.get("cat_id_level0")),
        "cat_id_level1": _as_optional_int(source.get("cat_id_level1")),
        "cat_id_level2": _as_optional_int(source.get("cat_id_level2")),
        "product_color": clean_cell(source.get("product_color")) or None,
        "product_description": desc or None,
        "fragile_property": _as_optional_int(source.get("fragile_property")),
        "product_size_type": clean_cell(source.get("product_size_type")) or None,
        "is_batch_tag": _as_optional_int(source.get("is_batch_tag")),
        "ean": ean,
        "ncm": ncm,
        "cest": cest,
        "sku_sort_code": clean_cell(source.get("sku_sort_code")) or None,
        "is_serialized": _as_optional_int(source.get("is_serialized")),
    }
    for key, value in optional.items():
        if value is not None and value != "":
            kwargs[key] = value
    return kwargs


def ensure_nw_skus_on_items(
    items: Sequence[Dict[str, Any]],
    *,
    dry_run: bool = False,
    cache: Optional[Dict[str, str]] = None,
) -> None:
    """将 items 中非 ``-NW`` SKU 改为 ``SKU-NW``，并确保 OMS 已有该产品。

    - 已以 ``-NW`` 结尾：不改动、不自动建品。
    - 否则：改写为 ``SKU-NW``；非 dry_run 时 getProductList 查询，
      不存在则按原 SKU 资料 createProduct（verify=1）。
    ``cache``：``nw_sku → exists|created``，同进程去重。
    """
    store = cache if cache is not None else {}
    for item in items:
        if not isinstance(item, dict):
            continue
        raw_sku = clean_cell(item.get("product_sku"))
        if not raw_sku:
            continue
        if raw_sku.endswith(NW_SUFFIX):
            continue

        base_sku = raw_sku
        nw_sku = _with_nw_suffix(base_sku)
        item["product_sku"] = nw_sku

        if dry_run:
            print(f"[DRY-RUN] ensure product {nw_sku} ← {base_sku}")
            continue

        cached = store.get(nw_sku)
        if cached:
            print(f"[SKU-NW] {nw_sku} cache={cached}")
            continue

        existing = _fetch_product_exact(nw_sku)
        if existing is not None:
            store[nw_sku] = "exists"
            print(f"[SKU-NW] {nw_sku} 已存在，跳过创建")
            continue

        source = _fetch_product_exact(base_sku)
        if source is None:
            raise ValueError(
                f"OMS 无原产品 {base_sku}，无法自动创建 {nw_sku}"
            )

        create_kwargs = _create_params_from_product(source, nw_sku=nw_sku)
        try:
            result = create_product(**create_kwargs)
        except HyOmsError as exc:
            raise ValueError(
                f"创建产品 {nw_sku} 失败（来源 {base_sku}）: {format_api_error(exc)}"
            ) from exc

        created_sku = clean_cell(result.get("product_sku")) or nw_sku
        store[nw_sku] = "created"
        print(
            f"[SKU-NW] 已创建 product_sku={created_sku} "
            f"ask={result.get('ask')} from={base_sku}"
        )


def _items_from_group(group: ReturnGroup, default_process: str) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    for row in group.rows:
        sku_raw = row.values.get("产品SKU", "")
        qty_raw = row.values.get("退件数量", "") or "1"
        try:
            default_qty = int(float(qty_raw))
        except ValueError as exc:
            raise ValueError(
                f"退件数量非法 key={group.key} sku={sku_raw}: {qty_raw}"
            ) from exc
        if default_qty <= 0:
            raise ValueError(f"退件数量须 > 0 key={group.key} sku={sku_raw}")

        try:
            products = _parse_sku_products(sku_raw, default_qty)
        except ValueError as exc:
            raise ValueError(f"{exc} key={group.key}") from exc
        if not products:
            raise ValueError(f"产品SKU为空 key={group.key}")

        process = _process_code(row.values, default_process)
        note = row.values.get("备注", "")
        for sku, qty in products:
            item: Dict[str, Any] = {
                "product_sku": sku,
                "quantity": qty,
                "process": str(process),
            }
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
    operation_desc: str = DEFAULT_OPERATION_DESC,
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
        operation_desc=operation_desc or DEFAULT_OPERATION_DESC,
        seller_store=_seller_store(head),
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
    operation_desc: str = DEFAULT_OPERATION_DESC,
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
        operation_desc=operation_desc or DEFAULT_OPERATION_DESC,
        buyer_name=(head.get("收件人", "") or head.get("Buyer ID", "") or None),
        seller_store=_seller_store(head),
        validate=False,
    )


def _resolve_shipped(
    group: ReturnGroup,
    shipped_map: Mapping[str, ShippedOrder],
) -> Union[ShippedOrder, str]:
    """成功返回 ShippedOrder；失败返回任务信息文案。"""
    erp = _erp_order_no(group)
    if not erp:
        return "缺少 ERP订单号，无法匹配原单"
    shipped = shipped_map.get(erp)
    if shipped is None:
        return (
            f"未找到原单 order_no={erp} "
            f"（sales_order_shipped / 易仓 getOrders）"
        )
    return shipped


def _build_params(group: ReturnGroup, options: CreateOptions) -> Dict[str, Any]:
    """按 mode 组装 params 并做本地必填校验。"""
    if options.mode == "mail":
        country = group.rows[0].values.get(COL_COUNTRY, "")
        params = build_mail_params(
            group,
            warehouse_code=options.warehouse_code,
            sm_code=resolve_sm_code(country, options.sm_code),
            verify=options.verify,
            default_process=options.default_process,
            return_type=options.return_type,
            operation_desc=options.operation_desc,
        )
    else:
        params = build_standard_params(
            group,
            warehouse_code=options.warehouse_code,
            return_type=options.return_type,
            verify=options.verify,
            default_process=options.default_process,
            operation_desc=options.operation_desc,
        )
    _validate_params_before_api(params, mode=options.mode)
    return params


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
        if not has_col(df, col_name):
            print(f"[WARN] 回写跳过，表格无列: [{col_name}]", file=sys.stderr)
            continue
        col_idx = sheet_col_index(df, col_name)
        safe_updates = [(row, cell_write_value(val)) for row, val in updates]
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
    options: CreateOptions,
) -> int:
    """处理分组；返回进程退出码（有失败则为 1）。"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    back = WriteBackBuffer()

    erp_nos = [_erp_order_no(g) for g in groups]
    shipped_map = resolve_shipped_orders(erp_nos)
    nw_cache: Dict[str, str] = {}

    for i, group in enumerate(groups, 1):
        label = f"[{i}/{len(groups)}] key={group.key} rows={len(group.rows)}"

        resolved = _resolve_shipped(group, shipped_map)
        if isinstance(resolved, str):
            stats.skipped_no_shipped += 1
            print(f"[SKIP] {label} {resolved}", file=sys.stderr)
            back.task_for_group(group, resolved)
            continue
        shipped = resolved

        filled = _apply_shipped_to_group(group, shipped)
        back.apply_shipped(group, shipped, filled)

        try:
            params = _build_params(group, options)
            ensure_nw_skus_on_items(
                params.get("items") or [],
                dry_run=options.dry_run,
                cache=nw_cache,
            )
        except ValueError as exc:
            stats.fail += 1
            msg = str(exc)
            print(f"[FAIL] {label} 参数: {msg}", file=sys.stderr)
            back.task_for_group(group, msg)
            continue
        except HyOmsError as exc:
            stats.fail += 1
            msg = format_api_error(exc)
            print(f"[FAIL] {label} SKU-NW: {msg}", file=sys.stderr)
            back.task_for_group(group, msg)
            continue

        if options.dry_run:
            print(
                f"[DRY-RUN] {label} oms={shipped.provider_order_no} "
                f"shop={shipped.shop_name_en}"
            )
            print(json.dumps(params, ensure_ascii=False, indent=2))
            stats.ok += 1
            continue

        try:
            result = HyOmsClient.from_config().call("createReturnBill", params)
        except HyOmsError as exc:
            stats.fail += 1
            msg = format_api_error(exc)
            print(f"[FAIL] {label} API: {msg}", file=sys.stderr)
            raw = getattr(exc, "raw", None)
            if raw is not None:
                print(
                    json.dumps(raw, ensure_ascii=False, indent=2)[:2000],
                    file=sys.stderr,
                )
            back.task_for_group(group, msg)
            continue

        return_code = clean_cell(result.get("return_code"))
        task = f"{now} 创建成功"
        print(f"[OK] {label} return_code={return_code}")
        stats.ok += 1
        back.mark_created(
            group,
            return_code=return_code,
            progress=PROGRESS_CREATED,
            task=task,
        )

    if options.write_back and not options.dry_run:
        n = _write_back(wb, sheet, df, back.as_updates())
        print(f"[WRITE-BACK] cells={n}")

    print(
        f"[DONE] sheet_rows={stats.sheet_rows} groups={stats.groups} "
        f"skipped_done={stats.skipped_done} skipped_empty={stats.skipped_empty} "
        f"skipped_unconfirmed={stats.skipped_unconfirmed} "
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
        default=None,
        help=(
            "回邮物流产品代码；默认按国家代码自动选择："
            f"DE→{DEFAULT_SM_CODE_DE}，其它→{DEFAULT_SM_CODE_OTHER}"
        ),
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
        "--operation-desc",
        dest="operation_desc",
        default=None,
        help=(
            f"OMS operation_desc；默认取 returned_config.operation_desc "
            f"或内置 {DEFAULT_OPERATION_DESC!r}"
        ),
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

        print(
            f"[CFG] workbook_id={workbook_id} "
            f"register_sheet={args.sheet or DEFAULT_SHEET} "
            f"warehouse_code={args.warehouse_code or DEFAULT_WAREHOUSE} "
            f"sm_code_de={DEFAULT_SM_CODE_DE} sm_code_other={DEFAULT_SM_CODE_OTHER} "
            f"operation_desc="
            f"{(args.operation_desc or DEFAULT_OPERATION_DESC)!r} "
            f"source={_CFG.config_path or '(defaults)'}"
        )
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

        rows = dataframe_to_rows(df, bool_as_true_false=True)
        groups, stats = build_groups(rows, force=bool(args.force), only_keys=args.keys)
        if args.limit is not None and args.limit >= 0:
            groups = groups[: args.limit]
            stats.groups = len(groups)

        if not groups:
            print(
                f"[DONE] 无可处理分组 sheet_rows={stats.sheet_rows} "
                f"skipped_done={stats.skipped_done} skipped_empty={stats.skipped_empty} "
                f"skipped_unconfirmed={stats.skipped_unconfirmed}"
            )
            return 0

        options = CreateOptions(
            mode=args.mode,
            warehouse_code=args.warehouse_code,
            sm_code=args.sm_code,
            return_type=args.return_type,
            verify=args.verify,
            default_process=str(args.process),
            operation_desc=(
                str(args.operation_desc).strip()
                if args.operation_desc
                else DEFAULT_OPERATION_DESC
            ),
            dry_run=bool(args.dry_run),
            write_back=not args.no_write_back,
        )
        return process_groups(wb, sheet, df, groups, stats, options)
    except (DingDiskError, HyOmsError, ValueError, KeyError) as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
