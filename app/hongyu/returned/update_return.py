"""从钉钉「退件登记表」拉取 OMS 退件详情并回填；同步写入「退件标签」。

表格读写委托 ``api.ding_disk.workbook.Workbook``；查询委托
``api.hy_oms.request.get_return`` / ``getReturnBill``。

文档 ID（workbookId / nodeId）::
    EpGBa2Lm8aDaZ57lTwEMk9boJgN7R35y

筛选条件（退件登记表）::
    ``OMS退件订单号`` 非空 且 ``标签跟踪号`` 为空

回写「退件登记表」::
    tracking_no / logistics_labels / 进度 / 任务信息 等

同步「退件标签」列::
    退件订单号 / 标签跟踪号 / ERP订单号 / 平台订单号 /
    销售平台←platform_shop.market_code /
    销售站点←market_region
      （店铺名称=shop_name_en 且 国家代码=platform_site）/
    店铺名称 / 自动发送←登记表 checkbox（是/否）
    （不同步标签地址）

    各类单号写入前去除两端空白；数据库不可用时：销售平台/销售站点留空，
    仍写入「退件标签」其余字段。

用法（项目根目录）::

    python app/hongyu/returned/update_return.py --dry-run
    python app/hongyu/returned/update_return.py --limit 1
    python app/hongyu/returned/update_return.py --list-sheets
    python app/hongyu/returned/update_return.py --return-code RMA900008-260727-0004

Windows 可执行文件（见 packaging/build_returned.ps1）::

    .\\dist\\returned\\update_return.exe --dry-run
    .\\dist\\returned\\run_task.exe
    # 配置放在 dist\\config\\（多模块共享），不是 exe 同级
"""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional, Sequence, Tuple, Union

import pandas as pd

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
    normalize_col_name,
    sheet_col_index,
    task_msg,
)

from api.hy_oms import HyOmsClient  # noqa: E402
from api.hy_oms.exceptions import HyOmsError  # noqa: E402
from api.hy_oms.request.get_return import (  # noqa: E402
    RETURN_STATUS,
    summarize_return,
)
from common.platform_shop import build_shop_maps, fetch_platform_shop_rows  # noqa: E402

WORKBOOK_ID = _CFG.workbook_id
DEFAULT_SHEET = _CFG.register_sheet
LABEL_SHEET = _CFG.label_sheet

COL_OMS_RETURN = "OMS退件订单号"
COL_LABEL_TRACKING = "标签跟踪号"
COL_LABEL_URL = "标签地址"
COL_LABEL_TIME = "标签制作时间"
COL_OMS_ORDER = "OMS原始订单号"
COL_SHOP = "店铺名称"
COL_COUNTRY = "国家代码"
COL_ERP = "ERP订单号"
COL_PLATFORM = "平台订单号"
COL_AUTO_SEND = "自动发送"
COL_PROGRESS = "进度"
COL_TASK = "任务信息"

LBL_RETURN_CODE = "退件订单号"
LBL_TRACKING = "标签跟踪号"
LBL_ERP = "ERP订单号"
LBL_PLATFORM_ORDER = "平台订单号"
LBL_PLATFORM = "销售平台"
LBL_SITE = "销售站点"
LBL_SHOP = "店铺名称"
LBL_AUTO_SEND = "自动发送"
LABEL_SHEET_COLS = (
    LBL_RETURN_CODE,
    LBL_TRACKING,
    LBL_ERP,
    LBL_PLATFORM_ORDER,
    LBL_PLATFORM,
    LBL_SITE,
    LBL_SHOP,
    LBL_AUTO_SEND,
)

WRITE_BACK_COLS = (
    COL_LABEL_TRACKING,
    COL_LABEL_TIME,
    COL_OMS_ORDER,
    COL_SHOP,
    COL_PROGRESS,
    COL_TASK,
)

# 进度 0-100：拿到标签跟踪号后记 30（创建成功为 10）
PROGRESS_TRACKED = 30

# 登记表 checkbox → 退件标签下拉「是」「否」
_AUTO_SEND_TRUE = frozenset(
    {"1", "true", "yes", "y", "是", "√", "✓", "checked", "勾选"}
)


@dataclass
class ReturnGroup:
    return_code: str
    rows: List[SheetRow] = field(default_factory=list)


@dataclass
class LabelSyncRow:
    """待同步到「退件标签」的一行。"""

    return_code: str
    tracking_no: str
    erp_order_no: str
    platform_order_no: str
    platform: str
    site: str
    shop_name: str
    auto_send: str  # 是 / 否


@dataclass
class RunStats:
    sheet_rows: int = 0
    candidates: int = 0
    groups: int = 0
    ok: int = 0
    pending: int = 0  # API 成功但尚无跟踪号
    fail: int = 0
    label_upserted: int = 0


@dataclass
class UpdateOptions:
    dry_run: bool = False
    write_back: bool = True


@dataclass
class WriteBackBuffer:
    """登记表回写缓冲：普通列 + 标签地址超链接。"""

    updates: Dict[str, List[Tuple[int, Any]]] = field(default_factory=dict)
    links: List[Tuple[int, Mapping[str, Any]]] = field(default_factory=list)

    def __post_init__(self) -> None:
        if not self.updates:
            self.updates = {col: [] for col in WRITE_BACK_COLS}

    def task_for_group(self, group: ReturnGroup, msg: str) -> None:
        text = task_msg(msg)
        for row in group.rows:
            self.updates[COL_TASK].append((row.excel_row, text))

    def apply_api_data(
        self,
        group: ReturnGroup,
        data: Mapping[str, Any],
        *,
        now: str,
    ) -> str:
        """按 API data 写入登记表缓冲；返回任务信息文案。"""
        tracking = clean_cell(data.get("tracking_no"))
        label_url = _logistics_labels_text(data)
        order_code = clean_cell(data.get("order_code"))
        shop = clean_cell(data.get("spo_seller_store"))
        status = str(data.get("return_status") or "").strip().upper()
        status_label = RETURN_STATUS.get(status, status or "?")

        if tracking:
            task = f"{now} 已回填跟踪号"
        else:
            task = f"{now} 查询成功但跟踪号为空 status={status}({status_label})"

        text = task_msg(task)
        for row in group.rows:
            if tracking:
                self.updates[COL_LABEL_TRACKING].append((row.excel_row, tracking))
                self.updates[COL_PROGRESS].append(
                    (row.excel_row, str(PROGRESS_TRACKED))
                )
                if label_url:
                    self.links.append((row.excel_row, _label_hyperlink(label_url)))
                if not row.values.get(COL_LABEL_TIME, ""):
                    self.updates[COL_LABEL_TIME].append((row.excel_row, now))
            if order_code and not row.values.get(COL_OMS_ORDER, ""):
                self.updates[COL_OMS_ORDER].append((row.excel_row, order_code))
            if shop and not row.values.get(COL_SHOP, ""):
                self.updates[COL_SHOP].append((row.excel_row, shop))
            self.updates[COL_TASK].append((row.excel_row, text))
        return task

    def as_updates(self) -> Mapping[str, Sequence[Tuple[int, Any]]]:
        return self.updates


def _logistics_labels_text(data: Mapping[str, Any]) -> str:
    """取 logistics_labels 首条下载 URL（单行）。"""
    labels = data.get("logistics_labels")
    if labels is None:
        return ""
    if isinstance(labels, str):
        text = clean_cell(labels)
        return text.splitlines()[0].strip() if text else ""
    if isinstance(labels, (list, tuple)):
        for item in labels:
            text = clean_cell(item)
            if text:
                return text.splitlines()[0].strip()
        return ""
    return clean_cell(labels)


def _label_link_title(url: str) -> str:
    """超链接标题：文件名主体后 8 位 + 扩展名，如 ``219efe0a.pdf``。"""
    name = clean_cell(url).rstrip("/").rsplit("/", 1)[-1]
    if not name:
        return "label"
    if "." in name:
        stem, ext = name.rsplit(".", 1)
        short = stem[-8:] if stem else name
        return f"{short}.{ext}" if ext else short
    return name[-8:] if len(name) > 8 else name


def _label_hyperlink(url: str) -> Dict[str, str]:
    """钉钉 ranges.hyperlinks 元素：type=path / link=URL / text=标题。"""
    link = clean_cell(url)
    return {
        "type": "path",
        "link": link,
        "text": _label_link_title(link),
    }


def build_groups(
    rows: Sequence[SheetRow],
    *,
    only_codes: Optional[Sequence[str]] = None,
    force: bool = False,
) -> Tuple[List[ReturnGroup], RunStats]:
    """OMS退件订单号非空 且（默认）标签跟踪号为空 → 按退件单号分组。"""
    stats = RunStats(sheet_rows=len(rows))
    wanted = {clean_cell(c) for c in (only_codes or []) if clean_cell(c)}
    buckets: Dict[str, ReturnGroup] = {}

    for row in rows:
        code = row.values.get(COL_OMS_RETURN, "")
        tracking = row.values.get(COL_LABEL_TRACKING, "")
        if not code:
            continue
        if tracking and not force:
            continue
        if wanted and code not in wanted:
            continue
        stats.candidates += 1
        grp = buckets.get(code)
        if grp is None:
            grp = ReturnGroup(return_code=code)
            buckets[code] = grp
        grp.rows.append(row)

    groups = list(buckets.values())
    stats.groups = len(groups)
    return groups, stats


def _shop_site_key(shop: str, country: str) -> str:
    """platform_shop 精确键：shop_name_en-platform_site（国家代码统一大写）。"""
    return f"{shop}-{str(country or '').strip().upper()}"


def _load_shop_market_maps() -> Tuple[
    Dict[str, str], Dict[str, str], Dict[str, str], Dict[str, str]
]:
    """加载店铺映射。

    返回::
        platform_by_shop, region_by_shop,
        platform_by_shop_site, region_by_shop_site

    ``*_by_shop_site`` 的 key 为 ``店铺英文名-国家代码``（对应 shop_name_en-platform_site）。
    数据库连接/配置失败时返回空映射，销售平台与销售站点留空，
    不阻断后续「退件标签」同步。
    """
    try:
        rows = fetch_platform_shop_rows()
        platform_by_shop, region_by_shop, _ = build_shop_maps(rows)
        platform_by_shop_site: Dict[str, str] = {}
        region_by_shop_site: Dict[str, str] = {}
        for r in rows:
            shop = str(r.get("shop_name_en") or "").strip()
            site = str(r.get("platform_site") or "").strip()
            region = str(r.get("market_region") or "").strip()
            code = str(r.get("market_code") or "").strip()
            if not (shop and site):
                continue
            key = _shop_site_key(shop, site)
            if code:
                platform_by_shop_site[key] = code
            if region:
                region_by_shop_site[key] = region
        return (
            platform_by_shop,
            region_by_shop,
            platform_by_shop_site,
            region_by_shop_site,
        )
    except Exception as exc:
        print(
            f"[WARN] 加载 platform_shop 失败，销售平台/销售站点将留空: {exc}",
            file=sys.stderr,
        )
        return {}, {}, {}, {}


def _resolve_shop_name(group: ReturnGroup, data: Mapping[str, Any]) -> str:
    head = group.rows[0].values
    return clean_cell(data.get("spo_seller_store")) or clean_cell(
        head.get(COL_SHOP, "")
    )


def _resolve_country(group: ReturnGroup, data: Mapping[str, Any]) -> str:
    """登记表「国家代码」优先；空则尝试 API 字段。"""
    head = group.rows[0].values
    country = clean_cell(head.get(COL_COUNTRY, ""))
    if country:
        return country.upper()
    for key in ("country", "country_code", "spo_country"):
        text = clean_cell(data.get(key)).upper()
        if text:
            return text
    return ""


def _resolve_platform_order_no(
    head: Mapping[str, str],
    data: Mapping[str, Any],
) -> str:
    """登记表「平台订单号」优先；无则 reference_no（与 ERP 不同或 ERP 为空时）。"""
    platform = clean_cell(head.get(COL_PLATFORM, ""))
    if platform:
        return platform
    ref = clean_cell(data.get("reference_no"))
    erp = clean_cell(head.get(COL_ERP, ""))
    if ref and (not erp or ref != erp):
        return ref
    return ""


def _auto_send_yes_no(val: Any) -> str:
    """登记表「自动发送」checkbox → 退件标签下拉「是」「否」。"""
    if isinstance(val, bool):
        return "是" if val else "否"
    if isinstance(val, (int, float)) and not (
        isinstance(val, float) and pd.isna(val)
    ):
        return "是" if int(val) == 1 else "否"
    text = clean_cell(val).lower()
    if text in _AUTO_SEND_TRUE:
        return "是"
    return "否"


def _build_label_sync(
    group: ReturnGroup,
    data: Mapping[str, Any],
    *,
    platform_by_shop: Mapping[str, str],
    region_by_shop: Mapping[str, str],
    platform_by_shop_site: Mapping[str, str],
    region_by_shop_site: Mapping[str, str],
) -> Optional[LabelSyncRow]:
    """从登记表行 + API + platform_shop 组装「退件标签」同步行；无跟踪号则 None。

    销售站点：店铺名称=shop_name_en 且 国家代码=platform_site → market_region。
    销售平台：同一精确键优先；无国家代码时回退按店铺英文名。
    platform/site 可为空（数据库不可用或未命中时）。
    各类单号写入前均 ``clean_cell`` 去两端空白。
    """
    tracking = clean_cell(data.get("tracking_no"))
    if not tracking:
        return None
    head = group.rows[0].values
    shop = _resolve_shop_name(group, data)
    country = _resolve_country(group, data)
    platform = ""
    site = ""
    if shop and country:
        key = _shop_site_key(shop, country)
        platform = platform_by_shop_site.get(key, "")
        site = region_by_shop_site.get(key, "")
    elif shop:
        platform = platform_by_shop.get(shop, "")
    # 仅在已加载到映射数据时提示未命中；DB 失败导致空映射则不逐条告警
    maps_loaded = bool(
        platform_by_shop
        or region_by_shop
        or platform_by_shop_site
        or region_by_shop_site
    )
    return_code = clean_cell(group.return_code)
    if shop and maps_loaded and not (platform or site):
        print(
            f"[WARN] platform_shop 未命中 "
            f"shop={shop} country={country or '(空)'} "
            f"return_code={return_code}",
            file=sys.stderr,
        )
    erp = clean_cell(head.get(COL_ERP, "")) or clean_cell(data.get("reference_no"))
    return LabelSyncRow(
        return_code=return_code,
        tracking_no=tracking,
        erp_order_no=erp,
        platform_order_no=_resolve_platform_order_no(head, data),
        platform=clean_cell(platform),
        site=clean_cell(site),
        shop_name=clean_cell(shop),
        auto_send=_auto_send_yes_no(head.get(COL_AUTO_SEND, "")),
    )


def _label_existing_codes(df: pd.DataFrame) -> Dict[str, int]:
    """退件订单号 → excel 行号（首条）。"""
    code_col = normalize_col_name(LBL_RETURN_CODE)
    norm_cols = [normalize_col_name(c) for c in df.columns]
    raw_code_col = df.columns[norm_cols.index(code_col)]
    existing: Dict[str, int] = {}
    for idx, val in df[raw_code_col].items():
        code = clean_cell(val)
        if code and code not in existing:
            existing[code] = int(idx) + 2
    return existing


def _label_row_values(rec: LabelSyncRow, df: pd.DataFrame) -> List[Any]:
    """按「退件标签」列序拼追加行；未同步列（如标签地址）留空。"""
    row_vals = [""] * len(df.columns)
    mapping = {
        LBL_RETURN_CODE: rec.return_code,
        LBL_TRACKING: rec.tracking_no,
        LBL_ERP: rec.erp_order_no,
        LBL_PLATFORM_ORDER: rec.platform_order_no,
        LBL_PLATFORM: rec.platform,
        LBL_SITE: rec.site,
        LBL_SHOP: rec.shop_name,
        LBL_AUTO_SEND: rec.auto_send,
    }
    for col_name, val in mapping.items():
        if not has_col(df, col_name):
            continue
        row_vals[sheet_col_index(df, col_name)] = cell_write_value(clean_cell(val))
    return row_vals


def sync_label_sheet(
    wb: Workbook,
    records: Sequence[LabelSyncRow],
    *,
    dry_run: bool = False,
) -> int:
    """upsert「退件标签」：按退件订单号匹配；不同步标签地址。"""
    if not records:
        return 0

    try:
        df = wb.read_sheet(LABEL_SHEET)
    except Exception as exc:
        print(f"[WARN] 无法读取工作表 [{LABEL_SHEET}]: {exc}", file=sys.stderr)
        return 0

    for col in LABEL_SHEET_COLS:
        if not has_col(df, col):
            print(f"[WARN] [{LABEL_SHEET}] 缺少列 [{col}]，跳过同步", file=sys.stderr)
            return 0

    existing = _label_existing_codes(df)
    updates: Dict[str, List[Tuple[int, Any]]] = {
        LBL_TRACKING: [],
        LBL_ERP: [],
        LBL_PLATFORM_ORDER: [],
        LBL_PLATFORM: [],
        LBL_SITE: [],
        LBL_SHOP: [],
        LBL_AUTO_SEND: [],
    }
    append_rows: List[List[Any]] = []

    for rec in records:
        if dry_run:
            print(
                f"[DRY-RUN][{LABEL_SHEET}] "
                f"{rec.return_code} tracking={rec.tracking_no} "
                f"shop={rec.shop_name} auto_send={rec.auto_send}"
            )
            continue

        if rec.return_code in existing:
            excel_row = existing[rec.return_code]
            updates[LBL_TRACKING].append((excel_row, clean_cell(rec.tracking_no)))
            if rec.erp_order_no:
                updates[LBL_ERP].append((excel_row, clean_cell(rec.erp_order_no)))
            if rec.platform_order_no:
                updates[LBL_PLATFORM_ORDER].append(
                    (excel_row, clean_cell(rec.platform_order_no))
                )
            if rec.platform:
                updates[LBL_PLATFORM].append((excel_row, clean_cell(rec.platform)))
            if rec.site:
                updates[LBL_SITE].append((excel_row, clean_cell(rec.site)))
            if rec.shop_name:
                updates[LBL_SHOP].append((excel_row, clean_cell(rec.shop_name)))
            updates[LBL_AUTO_SEND].append((excel_row, clean_cell(rec.auto_send)))
        else:
            append_rows.append(_label_row_values(rec, df))

    if dry_run:
        return len(records)

    written = 0
    updated_n = 0
    for col_name, col_updates in updates.items():
        if not col_updates:
            continue
        col_idx = sheet_col_index(df, col_name)
        safe = [(r, cell_write_value(clean_cell(v))) for r, v in col_updates]
        written += wb.write_column_updates(LABEL_SHEET, col_idx, safe)
        if col_name == LBL_TRACKING:
            updated_n = len(col_updates)

    if append_rows:
        wb.append_rows(LABEL_SHEET, append_rows)
        written += len(append_rows)

    print(
        f"[SYNC][{LABEL_SHEET}] total={len(records)} "
        f"updated={updated_n} appended={len(append_rows)} cells≈{written}"
    )
    return len(records)


def _write_back(
    wb: Workbook,
    sheet: str,
    df: pd.DataFrame,
    updates_by_col: Mapping[str, Sequence[Tuple[int, Any]]],
    *,
    hyperlink_updates: Optional[Sequence[Tuple[int, Mapping[str, Any]]]] = None,
) -> int:
    written = 0
    if hyperlink_updates:
        if not has_col(df, COL_LABEL_URL):
            print(f"[WARN] 回写跳过，表格无列: [{COL_LABEL_URL}]", file=sys.stderr)
        else:
            col_idx = sheet_col_index(df, COL_LABEL_URL)
            try:
                written += wb.write_hyperlink_column_updates(
                    sheet, col_idx, hyperlink_updates
                )
            except DingDiskError as exc:
                print(
                    f"[FAIL] 回写列[{COL_LABEL_URL}] 超链接失败: {exc}",
                    file=sys.stderr,
                )
                raise

    for col_name, updates in updates_by_col.items():
        if not updates:
            continue
        if col_name == COL_LABEL_URL:
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


def _fetch_return_data(
    client: HyOmsClient,
    return_code: str,
) -> Union[Mapping[str, Any], str]:
    """成功返回 getReturnBill.data；失败返回任务信息文案。"""
    try:
        result = client.get_return_bill(return_code=return_code)
    except HyOmsError as exc:
        msg = format_api_error(exc)
        raw = getattr(exc, "raw", None)
        if raw is not None:
            print(
                json.dumps(raw, ensure_ascii=False, indent=2)[:2000],
                file=sys.stderr,
            )
        return msg

    data = result.get("data")
    if not isinstance(data, Mapping):
        return "getReturnBill 响应无 data"
    return data


def process_groups(
    wb: Workbook,
    sheet: str,
    df: pd.DataFrame,
    groups: Sequence[ReturnGroup],
    stats: RunStats,
    options: UpdateOptions,
) -> int:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    back = WriteBackBuffer()
    label_records: List[LabelSyncRow] = []
    client = HyOmsClient.from_config()

    (
        platform_by_shop,
        region_by_shop,
        platform_by_shop_site,
        region_by_shop_site,
    ) = _load_shop_market_maps()
    if (
        platform_by_shop
        or region_by_shop
        or platform_by_shop_site
        or region_by_shop_site
    ):
        print(
            f"[DB] platform_shop 映射 "
            f"by_shop platform={len(platform_by_shop)} region={len(region_by_shop)}；"
            f"by_shop+country platform={len(platform_by_shop_site)} "
            f"region={len(region_by_shop_site)}"
        )
    else:
        print(
            "[DB] platform_shop 不可用或为空；"
            "退件标签将继续写入，销售平台/销售站点留空"
        )

    for i, group in enumerate(groups, 1):
        label = (
            f"[{i}/{len(groups)}] return_code={group.return_code} "
            f"rows={len(group.rows)}"
        )
        if options.dry_run:
            print(f"[DRY-RUN] {label} → getReturnBill")
            print(json.dumps({"return_code": group.return_code}, ensure_ascii=False))
            stats.ok += 1
            continue

        fetched = _fetch_return_data(client, group.return_code)
        if isinstance(fetched, str):
            stats.fail += 1
            print(f"[FAIL] {label} {fetched}", file=sys.stderr)
            back.task_for_group(group, fetched)
            continue
        data = fetched

        tracking = clean_cell(data.get("tracking_no"))
        task = back.apply_api_data(group, data, now=now)
        sync_row = _build_label_sync(
            group,
            data,
            platform_by_shop=platform_by_shop,
            region_by_shop=region_by_shop,
            platform_by_shop_site=platform_by_shop_site,
            region_by_shop_site=region_by_shop_site,
        )
        if sync_row is not None:
            label_records.append(sync_row)
        print(f"[OK] {label} {summarize_return(data)}")
        if tracking:
            stats.ok += 1
        else:
            stats.pending += 1
            print(f"[WARN] {label} {task}", file=sys.stderr)

    if options.write_back and not options.dry_run:
        n = _write_back(
            wb,
            sheet,
            df,
            back.as_updates(),
            hyperlink_updates=back.links,
        )
        print(f"[WRITE-BACK][{sheet}] cells={n}")
        stats.label_upserted = sync_label_sheet(wb, label_records, dry_run=False)

    print(
        f"[DONE] sheet_rows={stats.sheet_rows} candidates={stats.candidates} "
        f"groups={stats.groups} ok={stats.ok} pending={stats.pending} "
        f"fail={stats.fail} label_sync={stats.label_upserted}"
    )
    return 1 if stats.fail else 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="读取钉钉退件表，按 OMS退件订单号查询 getReturnBill 并回填跟踪号"
    )
    parser.add_argument(
        "--workbook-id",
        default=WORKBOOK_ID,
        help=f"钉钉表格文档 ID；默认 {WORKBOOK_ID}",
    )
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help=f"工作表名；默认 {DEFAULT_SHEET}")
    parser.add_argument("--list-sheets", action="store_true", help="仅列出工作表")
    parser.add_argument(
        "--return-code",
        action="append",
        dest="return_codes",
        default=None,
        help="仅处理指定 OMS退件订单号（可重复）",
    )
    parser.add_argument("--limit", type=int, default=None, help="最多处理 N 个退件单号")
    parser.add_argument(
        "--force",
        action="store_true",
        help="已有标签跟踪号也重新查询并回写",
    )
    parser.add_argument("--dry-run", action="store_true", help="只打印将查询的单号，不调 API/不回写")
    parser.add_argument(
        "--no-write-back",
        action="store_true",
        help="调用 API 但不回写钉钉表格",
    )
    parser.add_argument(
        "--preview",
        type=int,
        default=None,
        help="预览表格前 N 行后退出",
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
            f"label_sheet={LABEL_SHEET} "
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

        for required in (COL_OMS_RETURN, COL_LABEL_TRACKING):
            if not has_col(df, required):
                print(f"[FAIL] 表格缺少列: [{required}]", file=sys.stderr)
                return 2

        if args.preview is not None:
            n = args.preview if args.preview > 0 else len(df)
            print(df.head(n).fillna("").astype(str).to_string())
            return 0

        rows = dataframe_to_rows(df)
        groups, stats = build_groups(
            rows,
            only_codes=args.return_codes,
            force=bool(args.force),
        )
        if args.limit is not None and args.limit >= 0:
            groups = groups[: args.limit]
            stats.groups = len(groups)

        if not groups:
            print(
                f"[DONE] 无可更新行 sheet_rows={stats.sheet_rows} "
                f"（需 OMS退件订单号非空且标签跟踪号为空）"
            )
            return 0

        options = UpdateOptions(
            dry_run=bool(args.dry_run),
            write_back=not args.no_write_back,
        )
        return process_groups(wb, sheet, df, groups, stats, options)
    except (DingDiskError, HyOmsError, ValueError, KeyError) as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
