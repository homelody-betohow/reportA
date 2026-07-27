"""从钉钉在线表格拉取 OMS 退件详情，回填标签跟踪号等字段。

表格读写委托 ``api.ding_disk.workbook.Workbook``；查询委托
``api.hy_oms.request.get_return`` / ``getReturnBill``。

文档 ID（workbookId / nodeId）::
    EpGBa2Lm8aDaZ57lTwEMk9boJgN7R35y

筛选条件::
    ``OMS退件订单号`` 非空 且 ``标签跟踪号`` 为空

API ``data`` → 表格回写::
    tracking_no       → 标签跟踪号
    logistics_labels  → 标签地址（超链接：标题=文件名后8位+.ext，地址=完整下载 URL）
    order_code        → OMS原始订单号（表格为空时）
    spo_seller_store  → 店铺名（表格为空时）
    进度              → "50"（成功拿到跟踪号时；字符串写入，避免钉钉报 String is mandatory）
    任务信息          → 状态摘要或 API 错误原文

用法（项目根目录）::

    python app/hongyu/returned/update_return.py --dry-run
    python app/hongyu/returned/update_return.py --limit 1
    python app/hongyu/returned/update_return.py --list-sheets
    python app/hongyu/returned/update_return.py --return-code RMA900008-260727-0004
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
from api.hy_oms.request.get_return import (  # noqa: E402
    RETURN_STATUS,
    summarize_return,
)

WORKBOOK_ID = "EpGBa2Lm8aDaZ57lTwEMk9boJgN7R35y"
DEFAULT_SHEET = "Sheet1"

COL_OMS_RETURN = "OMS退件订单号"
COL_LABEL_TRACKING = "标签跟踪号"
COL_LABEL_URL = "标签地址"
COL_LABEL_TIME = "标签制作时间"
COL_OMS_ORDER = "OMS原始订单号"
COL_SHOP = "店铺名"
COL_PROGRESS = "进度"
COL_TASK = "任务信息"
TASK_MSG_MAX = 500

# 进度 0-100：拿到标签跟踪号后记 50（创建成功为 30）
PROGRESS_TRACKED = 50


@dataclass
class SheetRow:
    excel_row: int
    values: Dict[str, str]


@dataclass
class ReturnGroup:
    return_code: str
    rows: List[SheetRow] = field(default_factory=list)


@dataclass
class RunStats:
    sheet_rows: int = 0
    candidates: int = 0
    groups: int = 0
    ok: int = 0
    pending: int = 0  # API 成功但尚无跟踪号
    fail: int = 0


def _task_msg(text: str) -> str:
    return text[:TASK_MSG_MAX]


def _sheet_col_index(df: pd.DataFrame, col_name: str) -> int:
    cols = [clean_cell(c) for c in df.columns]
    try:
        return cols.index(clean_cell(col_name))
    except ValueError as exc:
        raise KeyError(f"表格缺少列: [{col_name}]；实际列={list(df.columns)}") from exc


def _format_api_error(exc: HyOmsError) -> str:
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


def dataframe_to_rows(df: pd.DataFrame) -> List[SheetRow]:
    rows: List[SheetRow] = []
    for idx, series in df.iterrows():
        values = {clean_cell(c): clean_cell(series[c]) for c in df.columns}
        rows.append(SheetRow(excel_row=int(idx) + 2, values=values))
    return rows


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


def _write_back(
    wb: Workbook,
    sheet: str,
    df: pd.DataFrame,
    updates_by_col: Mapping[str, Sequence[Tuple[int, Any]]],
    *,
    hyperlink_updates: Optional[Sequence[Tuple[int, Mapping[str, Any]]]] = None,
) -> int:
    written = 0
    # 标签地址用超链接写入（标题简化 + 完整 URL）
    if hyperlink_updates:
        if COL_LABEL_URL not in [clean_cell(c) for c in df.columns]:
            print(f"[WARN] 回写跳过，表格无列: [{COL_LABEL_URL}]", file=sys.stderr)
        else:
            col_idx = _sheet_col_index(df, COL_LABEL_URL)
            try:
                written += wb.write_hyperlink_column_updates(
                    sheet, col_idx, hyperlink_updates
                )
            except DingDiskError as exc:
                print(f"[FAIL] 回写列[{COL_LABEL_URL}] 超链接失败: {exc}", file=sys.stderr)
                raise

    for col_name, updates in updates_by_col.items():
        if not updates:
            continue
        if col_name == COL_LABEL_URL:
            continue  # 已由 hyperlinks 写入
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


def _apply_data_updates(
    group: ReturnGroup,
    data: Mapping[str, Any],
    back: Dict[str, List[Tuple[int, Any]]],
    link_back: List[Tuple[int, Mapping[str, Any]]],
    *,
    now: str,
) -> str:
    """根据 API data 组装回写；返回任务信息文案。"""
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

    for row in group.rows:
        if tracking:
            back[COL_LABEL_TRACKING].append((row.excel_row, tracking))
            back[COL_PROGRESS].append((row.excel_row, str(PROGRESS_TRACKED)))
            if label_url:
                link_back.append((row.excel_row, _label_hyperlink(label_url)))
            # 首次拿到跟踪号时补标签制作时间（表格为空才写）
            if not row.values.get(COL_LABEL_TIME, ""):
                back[COL_LABEL_TIME].append((row.excel_row, now))
        if order_code and not row.values.get(COL_OMS_ORDER, ""):
            back[COL_OMS_ORDER].append((row.excel_row, order_code))
        if shop and not row.values.get(COL_SHOP, ""):
            back[COL_SHOP].append((row.excel_row, shop))
        back[COL_TASK].append((row.excel_row, _task_msg(task)))
    return task


def process_groups(
    wb: Workbook,
    sheet: str,
    df: pd.DataFrame,
    groups: Sequence[ReturnGroup],
    stats: RunStats,
    *,
    dry_run: bool,
    write_back: bool,
) -> int:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    back: Dict[str, List[Tuple[int, Any]]] = {
        COL_LABEL_TRACKING: [],
        COL_LABEL_TIME: [],
        COL_OMS_ORDER: [],
        COL_SHOP: [],
        COL_PROGRESS: [],
        COL_TASK: [],
    }
    link_back: List[Tuple[int, Mapping[str, Any]]] = []
    client = HyOmsClient.from_config()

    for i, group in enumerate(groups, 1):
        label = (
            f"[{i}/{len(groups)}] return_code={group.return_code} "
            f"rows={len(group.rows)}"
        )
        if dry_run:
            print(f"[DRY-RUN] {label} → getReturnBill")
            print(json.dumps({"return_code": group.return_code}, ensure_ascii=False))
            stats.ok += 1
            continue

        try:
            result = client.get_return_bill(return_code=group.return_code)
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

        data = result.get("data")
        if not isinstance(data, Mapping):
            stats.fail += 1
            msg = "getReturnBill 响应无 data"
            print(f"[FAIL] {label} {msg}", file=sys.stderr)
            for row in group.rows:
                back[COL_TASK].append((row.excel_row, _task_msg(msg)))
            continue

        tracking = clean_cell(data.get("tracking_no"))
        task = _apply_data_updates(group, data, back, link_back, now=now)
        print(f"[OK] {label} {summarize_return(data)}")
        if tracking:
            stats.ok += 1
        else:
            stats.pending += 1
            print(f"[WARN] {label} {task}", file=sys.stderr)

    if write_back and not dry_run:
        n = _write_back(wb, sheet, df, back, hyperlink_updates=link_back)
        print(f"[WRITE-BACK] cells={n}")

    print(
        f"[DONE] sheet_rows={stats.sheet_rows} candidates={stats.candidates} "
        f"groups={stats.groups} ok={stats.ok} pending={stats.pending} fail={stats.fail}"
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
            if required not in [clean_cell(c) for c in df.columns]:
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

        return process_groups(
            wb,
            sheet,
            df,
            groups,
            stats,
            dry_run=bool(args.dry_run),
            write_back=not args.no_write_back,
        )
    except (DingDiskError, HyOmsError, ValueError, KeyError) as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
