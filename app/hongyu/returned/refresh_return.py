"""从钉钉「退件登记表」拉取 OMS 退件详情，刷新「仓库是否收到」。

表格读写委托 ``api.ding_disk.workbook.Workbook``；查询委托
``api.hy_oms.request.get_return`` / ``getReturnBill``。

文档 ID（workbookId / nodeId）::
    EpGBa2Lm8aDaZ57lTwEMk9boJgN7R35y

筛选条件（退件登记表）::
    ``OMS退件订单号`` 非空 且 ``仓库是否收到`` 为空或 <> F（已完成）

回写「退件登记表」::
    return_status=D → 仓库是否收到=到货，进度=80
    return_status=F → 仓库是否收到=已完成，进度=100
    任务信息

用法（项目根目录）::

    python app/hongyu/returned/refresh_return.py --dry-run
    python app/hongyu/returned/refresh_return.py --limit 1
    python app/hongyu/returned/refresh_return.py --list-sheets
    python app/hongyu/returned/refresh_return.py --return-code RMA900008-260727-0004

Windows 可执行文件（见 packaging/build_returned.ps1）::

    .\\dist\\returned\\refresh_return.exe --dry-run
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
    sheet_col_index,
    task_msg,
)

from api.hy_oms import HyOmsClient  # noqa: E402
from api.hy_oms.exceptions import HyOmsError  # noqa: E402
from api.hy_oms.request.get_return import (  # noqa: E402
    RETURN_STATUS,
    summarize_return,
)

WORKBOOK_ID = _CFG.workbook_id
DEFAULT_SHEET = _CFG.register_sheet

COL_OMS_RETURN = "OMS退件订单号"
COL_WAREHOUSE_RECV = "仓库是否收到"
COL_PROGRESS = "进度"
COL_TASK = "任务信息"

# return_status → 表格「仓库是否收到」文案
WAREHOUSE_RECV_BY_STATUS: Mapping[str, str] = {
    "D": "到货",
    "F": "已完成",
}

# return_status → 进度 0-100（创建 10 / 跟踪号 30 / 标签 50 / 到货 80 / 完成 100）
PROGRESS_BY_STATUS: Mapping[str, int] = {
    "D": 80,
    "F": 100,
}

# 筛选时视为「已完成」而不再刷新（除非 --force）
WAREHOUSE_DONE = frozenset({"F", "已完成"})

WRITE_BACK_COLS = (
    COL_WAREHOUSE_RECV,
    COL_PROGRESS,
    COL_TASK,
)


@dataclass
class ReturnGroup:
    return_code: str
    rows: List[SheetRow] = field(default_factory=list)


@dataclass
class RunStats:
    sheet_rows: int = 0
    candidates: int = 0
    groups: int = 0
    ok: int = 0  # 已回填 D/F
    skip: int = 0  # API 成功但 status 非 D/F
    fail: int = 0


@dataclass
class RefreshOptions:
    dry_run: bool = False
    write_back: bool = True


@dataclass
class WriteBackBuffer:
    updates: Dict[str, List[Tuple[int, Any]]] = field(default_factory=dict)

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
    ) -> Tuple[str, bool]:
        """按 API data 写入缓冲；返回 (任务信息, 是否已回填仓库是否收到)。"""
        status = str(data.get("return_status") or "").strip().upper()
        status_label = RETURN_STATUS.get(status, status or "?")
        warehouse_val = WAREHOUSE_RECV_BY_STATUS.get(status, "")
        progress = PROGRESS_BY_STATUS.get(status)

        if warehouse_val:
            task = (
                f"{now} 仓库是否收到={warehouse_val} "
                f"进度={progress} status={status}({status_label})"
            )
        else:
            task = f"{now} 查询成功 status={status}({status_label})"

        text = task_msg(task)
        for row in group.rows:
            if warehouse_val:
                self.updates[COL_WAREHOUSE_RECV].append(
                    (row.excel_row, warehouse_val)
                )
            if progress is not None:
                self.updates[COL_PROGRESS].append(
                    (row.excel_row, str(progress))
                )
            self.updates[COL_TASK].append((row.excel_row, text))
        return task, bool(warehouse_val)

    def as_updates(self) -> Mapping[str, Sequence[Tuple[int, Any]]]:
        return self.updates


def _is_warehouse_done(value: str) -> bool:
    """仓库是否收到 为 F / 已完成。"""
    text = clean_cell(value).strip()
    if not text:
        return False
    return text.upper() == "F" or text in WAREHOUSE_DONE


def build_groups(
    rows: Sequence[SheetRow],
    *,
    only_codes: Optional[Sequence[str]] = None,
    force: bool = False,
) -> Tuple[List[ReturnGroup], RunStats]:
    """OMS退件订单号非空 且（默认）仓库是否收到为空或 <> F → 按退件单号分组。"""
    stats = RunStats(sheet_rows=len(rows))
    wanted = {clean_cell(c) for c in (only_codes or []) if clean_cell(c)}
    buckets: Dict[str, ReturnGroup] = {}

    for row in rows:
        code = row.values.get(COL_OMS_RETURN, "")
        warehouse = row.values.get(COL_WAREHOUSE_RECV, "")
        if not code:
            continue
        if _is_warehouse_done(warehouse) and not force:
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
    options: RefreshOptions,
) -> int:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    back = WriteBackBuffer()
    client = HyOmsClient.from_config()

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

        task, updated = back.apply_api_data(group, data, now=now)
        print(f"[OK] {label} {summarize_return(data)}")
        if updated:
            stats.ok += 1
        else:
            stats.skip += 1
            print(f"[SKIP] {label} {task}")

    if options.write_back and not options.dry_run:
        n = _write_back(wb, sheet, df, back.as_updates())
        print(f"[WRITE-BACK][{sheet}] cells={n}")

    print(
        f"[DONE] sheet_rows={stats.sheet_rows} candidates={stats.candidates} "
        f"groups={stats.groups} ok={stats.ok} skip={stats.skip} fail={stats.fail}"
    )
    return 1 if stats.fail else 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="读取钉钉退件表，按 OMS退件订单号查询 getReturnBill 并刷新仓库是否收到"
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
        help="仓库是否收到已为 F/已完成 也重新查询并回写",
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

        for required in (COL_OMS_RETURN, COL_WAREHOUSE_RECV):
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
                f"[DONE] 无可刷新行 sheet_rows={stats.sheet_rows} "
                f"（需 OMS退件订单号非空且仓库是否收到为空或<>F）"
            )
            return 0

        options = RefreshOptions(
            dry_run=bool(args.dry_run),
            write_back=not args.no_write_back,
        )
        return process_groups(wb, sheet, df, groups, stats, options)
    except (DingDiskError, HyOmsError, ValueError, KeyError) as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
