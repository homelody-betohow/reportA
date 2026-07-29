"""用 OMS 模拟登录下载退件标签文件。

流程参考 ``con_test.py``：
1. 仅读取钉钉「退件标签」：``退件订单号`` 非空且 ``标签路径`` 为空 → 待下载
2. SOAP ``logOn`` 拿快捷登录 URL → 访问后获得 Web Cookie
3. ``POST /order/special-orders/download-label``，表单字段 ``code`` = 退件单号
4. 保存到共享盘；若为 zip 则解压并删除压缩包，再遍历解压目录，
   将文件名中的退件订单号替换为平台订单号；解压目录回写 ``标签路径``
5. 下载成功后，将「退件登记表」对应行（按 ``OMS退件订单号``）的 ``进度`` 设为 50

默认保存目录::

    \\\\Betohow\\数据报表\\RPA\\退货标签\\yyyy-mm\\店铺名称

用法（项目根目录）::

    python app/hongyu/returned/download_label.py
    python app/hongyu/returned/download_label.py --return-code RMA900008-260727-0004
    python app/hongyu/returned/download_label.py --month 2026-07 --limit 5
    python app/hongyu/returned/download_label.py --force
    python app/hongyu/returned/download_label.py --dry-run

Windows 可执行文件（见 packaging/build_returned.ps1）::

    .\\dist\\returned\\download_label.exe --dry-run
    .\\dist\\returned\\run_task.exe
    # 配置放在 dist\\config\\（多模块共享），不是 exe 同级
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

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
    has_col,
    sheet_col_index,
)

from api.hy_oms import HyOmsWebSession  # noqa: E402
from api.hy_oms.exceptions import HyOmsError  # noqa: E402

WORKBOOK_ID = _CFG.workbook_id
LABEL_SHEET = _CFG.label_sheet
REGISTER_SHEET = _CFG.register_sheet

LBL_RETURN_CODE = "退件订单号"
LBL_TRACKING = "标签跟踪号"
LBL_ERP = "ERP订单号"
LBL_PLATFORM_ORDER = "平台订单号"
LBL_SHOP = "店铺名称"
LBL_PATH = "标签路径"

COL_OMS_RETURN = "OMS退件订单号"
COL_PROGRESS = "进度"

# 进度 0-100：创建成功 30；标签下载成功 50
PROGRESS_DOWNLOADED = 50

REQUIRED_LABEL_COLS = (LBL_RETURN_CODE, LBL_PATH)

# 共享盘：根目录 \ yyyy-mm \ 店铺名称（默认来自 returned_config.json）
LABEL_SHARE_ROOT = Path(_CFG.label_share_root)
UNKNOWN_SHOP = "_无店铺"


def resolve_month(month: Optional[str] = None) -> str:
    """返回 ``yyyy-mm``。"""
    if month is None or not str(month).strip():
        return datetime.now().strftime("%Y-%m")
    ym = str(month).strip()
    if not re.fullmatch(r"\d{4}-\d{2}", ym):
        raise ValueError(f"month 应为 yyyy-mm，收到: {ym!r}")
    return ym


def month_base_dir(
    *,
    month: Optional[str] = None,
    share_root: Optional[Path] = None,
) -> Path:
    """``share_root\\yyyy-mm``。"""
    root = Path(share_root) if share_root is not None else LABEL_SHARE_ROOT
    return root / resolve_month(month)


def _safe_folder_name(name: str) -> str:
    """去掉 Windows 非法路径字符，空白压成下划线。"""
    text = clean_cell(name)
    if not text:
        return UNKNOWN_SHOP
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", text)
    text = re.sub(r"\s+", "_", text).strip(" ._")
    return text or UNKNOWN_SHOP


def shop_out_dir(
    shop_name: str,
    *,
    month: Optional[str] = None,
    share_root: Optional[Path] = None,
) -> Path:
    """``…\\yyyy-mm\\店铺名称``。"""
    return month_base_dir(month=month, share_root=share_root) / _safe_folder_name(
        shop_name
    )


@dataclass
class DownloadItem:
    """待下载的一条退件（来自「退件标签」）。"""

    return_code: str
    excel_row: int
    shop_name: str = ""
    tracking_no: str = ""
    erp_order_no: str = ""
    platform_order_no: str = ""


@dataclass
class DownloadOk:
    item: DownloadItem
    path: Path


@dataclass
class RunStats:
    sheet_rows: int = 0
    skipped_has_path: int = 0
    skipped_empty_code: int = 0
    candidates: int = 0
    codes: int = 0
    ok: int = 0
    fail: int = 0
    label_updated: int = 0
    progress_updated: int = 0


@dataclass
class DownloadOptions:
    month: str = field(default_factory=lambda: resolve_month())
    share_root: Path = field(default_factory=lambda: LABEL_SHARE_ROOT)
    out_dir_override: Optional[Path] = None  # 若指定则忽略店铺子目录
    dry_run: bool = False
    overwrite: bool = False
    write_back: bool = True
    account: Optional[str] = None
    password: Optional[str] = None

    def dir_for(self, shop_name: str) -> Path:
        if self.out_dir_override is not None:
            return self.out_dir_override
        return shop_out_dir(
            shop_name, month=self.month, share_root=self.share_root
        )


def collect_download_items(
    rows: Sequence[SheetRow],
    *,
    only_codes: Optional[Sequence[str]] = None,
    force: bool = False,
) -> Tuple[List[DownloadItem], RunStats]:
    """从「退件标签」收集待下载项：退件订单号非空，且（默认）标签路径为空。"""
    stats = RunStats(sheet_rows=len(rows))
    wanted = {clean_cell(c) for c in (only_codes or []) if clean_cell(c)}
    seen: set[str] = set()
    items: List[DownloadItem] = []

    for row in rows:
        code = row.values.get(LBL_RETURN_CODE, "")
        path = row.values.get(LBL_PATH, "")
        if not code:
            stats.skipped_empty_code += 1
            continue
        if wanted and code not in wanted:
            continue
        if path and not force:
            stats.skipped_has_path += 1
            continue
        stats.candidates += 1
        if code in seen:
            continue
        seen.add(code)
        items.append(
            DownloadItem(
                return_code=code,
                excel_row=row.excel_row,
                shop_name=row.values.get(LBL_SHOP, ""),
                tracking_no=row.values.get(LBL_TRACKING, ""),
                erp_order_no=row.values.get(LBL_ERP, ""),
                platform_order_no=row.values.get(LBL_PLATFORM_ORDER, ""),
            )
        )

    stats.codes = len(items)
    return items, stats


def _safe_stem(return_code: str) -> str:
    return re.sub(r"[^\w.\-]+", "_", str(return_code).strip()) or "label"


def _unique_path(path: Path, *, overwrite: bool) -> Path:
    """若文件/目录已存在且不允许覆盖，则追加 _2/_3…"""
    if overwrite or not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    i = 2
    while True:
        alt = path.with_name(f"{stem}_{i}{suffix}")
        if not alt.exists():
            return alt
        i += 1


def _looks_like_zip(path: Path, *, content_type: str = "", content: bytes = b"") -> bool:
    ct = (content_type or "").lower()
    if "zip" in ct or path.suffix.lower() == ".zip":
        return True
    head = content[:4] if content else b""
    if not head and path.is_file():
        try:
            with path.open("rb") as f:
                head = f.read(4)
        except OSError:
            return False
    return head.startswith(b"PK\x03\x04") or head.startswith(b"PK\x05\x06")


def extract_zip_then_delete(
    zip_path: Path,
    *,
    overwrite: bool = False,
) -> Path:
    """以压缩包文件名（无扩展名）建目录，解压到该目录后删除压缩包。

    返回解压目录路径（用于回填 ``标签路径``）。
    """
    extract_dir = _unique_path(zip_path.with_suffix(""), overwrite=overwrite)
    extract_dir.mkdir(parents=True, exist_ok=True)
    extracted_n = 0

    with zipfile.ZipFile(zip_path, "r") as zf:
        members = [
            info
            for info in zf.infolist()
            if not info.is_dir() and info.filename and not info.filename.endswith("/")
        ]
        if not members:
            raise ValueError(f"压缩包为空: {zip_path}")

        for info in members:
            base = Path(info.filename.replace("\\", "/")).name
            if not base or base in {".", ".."}:
                continue
            dest = _unique_path(extract_dir / base, overwrite=overwrite)
            with zf.open(info, "r") as src, dest.open("wb") as dst:
                dst.write(src.read())
            extracted_n += 1

    if extracted_n == 0:
        raise ValueError(f"压缩包无可解压文件: {zip_path}")

    try:
        zip_path.unlink()
    except OSError as exc:
        print(f"[WARN] 解压成功但删除压缩包失败: {zip_path} ({exc})", file=sys.stderr)
    return extract_dir


def rename_return_code_in_dir(
    directory: Path,
    *,
    return_code: str,
    platform_order_no: str,
    overwrite: bool = False,
) -> int:
    """遍历目录内文件，将文件名中的退件订单号替换为平台订单号。

    无平台订单号、或文件名不含退件订单号时跳过。返回成功重命名数。
    """
    old = clean_cell(return_code)
    new = clean_cell(platform_order_no)
    if not old or not new or old == new:
        return 0
    if not directory.is_dir():
        return 0

    old_safe = _safe_stem(old)
    new_safe = _safe_stem(new)
    if not new_safe:
        return 0

    renamed = 0
    for path in sorted(directory.iterdir()):
        if not path.is_file():
            continue
        name = path.name
        if old in name:
            new_name = name.replace(old, new_safe)
        elif old_safe != old and old_safe in name:
            new_name = name.replace(old_safe, new_safe)
        else:
            continue
        if new_name == name:
            continue
        dest = _unique_path(path.with_name(new_name), overwrite=overwrite)
        try:
            path.rename(dest)
            renamed += 1
            print(f"[RENAME] {name} -> {dest.name}")
        except OSError as exc:
            print(f"[WARN] 重命名失败 {name}: {exc}", file=sys.stderr)
    return renamed


def save_download(
    *,
    out_dir: Path,
    return_code: str,
    content: bytes,
    filename: str,
    overwrite: bool,
    content_type: str = "",
    platform_order_no: str = "",
) -> Path:
    """保存标签文件；若为 zip 则建同名目录解压并删除压缩包。

    解压后将目录内文件名的退件订单号替换为平台订单号。
    返回路径：解压成功为解压目录；非压缩包为文件本身（供回填 ``标签路径``）。
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    name = Path(filename).name or f"{_safe_stem(return_code)}.bin"
    ext = Path(name).suffix or ".bin"
    if _looks_like_zip(Path(name), content_type=content_type, content=content):
        ext = ".zip"
    target = _unique_path(
        out_dir / f"{_safe_stem(return_code)}{ext}",
        overwrite=overwrite,
    )
    target.write_bytes(content)

    if _looks_like_zip(target, content_type=content_type, content=content):
        try:
            extract_dir = extract_zip_then_delete(target, overwrite=overwrite)
            print(f"[UNZIP] {target.name} -> {extract_dir}\\")
            n = rename_return_code_in_dir(
                extract_dir,
                return_code=return_code,
                platform_order_no=platform_order_no,
                overwrite=overwrite,
            )
            if n:
                print(
                    f"[RENAME] {extract_dir.name}\\ 内 {n} 个文件 "
                    f"{clean_cell(return_code)} → {clean_cell(platform_order_no)}"
                )
            elif clean_cell(platform_order_no):
                print(
                    f"[RENAME] 跳过：解压目录内文件名未含退件订单号 "
                    f"{clean_cell(return_code)}"
                )
            else:
                print(
                    f"[RENAME] 跳过：平台订单号为空 code={clean_cell(return_code)}"
                )
            return extract_dir
        except (zipfile.BadZipFile, OSError, ValueError) as exc:
            print(
                f"[WARN] 解压失败，保留压缩包 {target}: {exc}",
                file=sys.stderr,
            )
            return target
    return target


def _path_text(path: Path) -> str:
    """UNC/本地路径字符串（Windows 反斜杠）。"""
    return str(path)


def sync_label_paths(
    wb: Workbook,
    label_df: pd.DataFrame,
    successes: Sequence[DownloadOk],
    *,
    dry_run: bool = False,
) -> int:
    """把解压目录全路径回写「退件标签」的 ``标签路径``（纯文本，按原行号更新）。"""
    if not successes:
        return 0

    if not has_col(label_df, LBL_PATH):
        print(f"[WARN] [{LABEL_SHEET}] 缺少列 [{LBL_PATH}]，跳过回写", file=sys.stderr)
        return 0

    path_updates: List[Tuple[int, Any]] = []

    for ok in successes:
        item, path = ok.item, ok.path
        path_s = _path_text(path)
        excel_row = item.excel_row
        if dry_run:
            print(
                f"[DRY-RUN][{LABEL_SHEET}] row={excel_row} "
                f"{item.return_code} shop={item.shop_name or UNKNOWN_SHOP} path={path_s}"
            )
            continue
        path_updates.append((excel_row, path_s))

    if dry_run:
        return len(successes)

    col_idx = sheet_col_index(label_df, LBL_PATH)
    safe = [(r, cell_write_value(v)) for r, v in path_updates]
    written = wb.write_column_updates(LABEL_SHEET, col_idx, safe) if safe else 0

    print(
        f"[SYNC][{LABEL_SHEET}] total={len(successes)} "
        f"updated={len(path_updates)} cells≈{written}"
    )
    return len(successes)


def sync_register_progress(
    wb: Workbook,
    successes: Sequence[DownloadOk],
    *,
    progress: int = PROGRESS_DOWNLOADED,
    dry_run: bool = False,
) -> int:
    """下载成功后，将「退件登记表」匹配 ``OMS退件订单号`` 的行 ``进度`` 设为 progress。"""
    if not successes:
        return 0

    codes = {ok.item.return_code for ok in successes if ok.item.return_code}
    if not codes:
        return 0

    try:
        df = wb.read_sheet(REGISTER_SHEET)
    except Exception as exc:
        print(f"[WARN] 无法读取工作表 [{REGISTER_SHEET}]: {exc}", file=sys.stderr)
        return 0

    if not has_col(df, COL_OMS_RETURN):
        print(
            f"[WARN] [{REGISTER_SHEET}] 缺少列 [{COL_OMS_RETURN}]，跳过进度回写",
            file=sys.stderr,
        )
        return 0
    if not has_col(df, COL_PROGRESS):
        print(
            f"[WARN] [{REGISTER_SHEET}] 缺少列 [{COL_PROGRESS}]，跳过进度回写",
            file=sys.stderr,
        )
        return 0

    progress_s = cell_write_value(progress)
    updates: List[Tuple[int, Any]] = []
    for row in dataframe_to_rows(df):
        code = row.values.get(COL_OMS_RETURN, "")
        if code not in codes:
            continue
        if dry_run:
            print(
                f"[DRY-RUN][{REGISTER_SHEET}] row={row.excel_row} "
                f"code={code} {COL_PROGRESS}={progress_s}"
            )
        updates.append((row.excel_row, progress_s))

    if dry_run:
        return len(updates)

    if not updates:
        print(
            f"[WARN] [{REGISTER_SHEET}] 未匹配到 OMS退件订单号，跳过进度回写 "
            f"codes={sorted(codes)}",
            file=sys.stderr,
        )
        return 0

    col_idx = sheet_col_index(df, COL_PROGRESS)
    written = wb.write_column_updates(REGISTER_SHEET, col_idx, updates)
    print(
        f"[SYNC][{REGISTER_SHEET}] {COL_PROGRESS}={progress_s} "
        f"rows={len(updates)} cells≈{written}"
    )
    return len(updates)


def download_items(
    session: HyOmsWebSession,
    items: Sequence[DownloadItem],
    options: DownloadOptions,
    stats: Optional[RunStats] = None,
) -> Tuple[RunStats, List[DownloadOk]]:
    stats = stats or RunStats(codes=len(items))
    stats.codes = len(items)
    successes: List[DownloadOk] = []

    if options.dry_run:
        for i, item in enumerate(items, 1):
            out = options.dir_for(item.shop_name) / _safe_stem(item.return_code)
            print(
                f"[DRY-RUN] [{i}/{len(items)}] row={item.excel_row} "
                f"code={item.return_code} shop={item.shop_name or UNKNOWN_SHOP} -> {out}"
            )
        return stats, successes

    login_kwargs: Dict[str, Any] = {}
    if options.account is not None:
        login_kwargs["user_account"] = options.account
    if options.password is not None:
        login_kwargs["user_password"] = options.password

    print("[LOGIN] OMS 模拟登录…")
    login_info = session.login(**login_kwargs)
    cookie_names = sorted((login_info.get("cookies") or {}).keys())
    print(f"[LOGIN] OK cookies={cookie_names}")

    for i, item in enumerate(items, 1):
        shop = item.shop_name or UNKNOWN_SHOP
        label = (
            f"[{i}/{len(items)}] row={item.excel_row} "
            f"code={item.return_code} shop={shop}"
        )
        try:
            result = session.download_label(item.return_code, auto_login=False)
            path = save_download(
                out_dir=options.dir_for(item.shop_name),
                return_code=item.return_code,
                content=result.content,
                filename=result.filename,
                overwrite=options.overwrite,
                content_type=result.content_type,
                platform_order_no=item.platform_order_no,
            )
            stats.ok += 1
            successes.append(DownloadOk(item=item, path=path))
            print(
                f"[OK] {label} -> {path} "
                f"({len(result.content)} bytes, {result.content_type})"
            )
        except (HyOmsError, OSError, ValueError) as exc:
            stats.fail += 1
            print(f"[FAIL] {label} {exc}", file=sys.stderr)

    return stats, successes


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="鸿羽 OMS：从「退件标签」下载标签路径为空的退件标签"
    )
    parser.add_argument(
        "--workbook-id",
        default=WORKBOOK_ID,
        help=f"钉钉表格文档 ID；默认 {WORKBOOK_ID}",
    )
    parser.add_argument(
        "--sheet",
        default=LABEL_SHEET,
        help=f"工作表名；默认 {LABEL_SHEET}",
    )
    parser.add_argument("--list-sheets", action="store_true", help="仅列出工作表")
    parser.add_argument(
        "--return-code",
        action="append",
        dest="return_codes",
        default=None,
        help="仅处理指定退件订单号（可重复）",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="标签路径已有值也重新下载并覆盖回写",
    )
    parser.add_argument("--limit", type=int, default=None, help="最多下载 N 个单号")
    parser.add_argument(
        "--month",
        default=None,
        help="保存子目录年月 yyyy-mm（默认当月）",
    )
    parser.add_argument(
        "--label-share-root",
        default=str(LABEL_SHARE_ROOT),
        help=f"标签共享盘根目录；默认 {LABEL_SHARE_ROOT}",
    )
    parser.add_argument(
        "--out-dir",
        default=None,
        help=(
            f"覆盖保存目录（不再按店铺分子目录）；"
            f"默认 {{label-share-root}}\\yyyy-mm\\店铺名称"
        ),
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="覆盖已存在的同名解压目录/文件",
    )
    parser.add_argument("--account", default=None, help="OMS 登录账号（缺省用 config）")
    parser.add_argument("--password", default=None, help="OMS 登录密码（缺省用 config）")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只打印计划，不登录/不下载/不回写",
    )
    parser.add_argument(
        "--no-write-back",
        action="store_true",
        help="下载成功后不回写「标签路径」/登记表「进度」",
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
        month = resolve_month(args.month)
        out_override = Path(args.out_dir) if args.out_dir else None
        share_root = Path(
            (args.label_share_root or str(LABEL_SHARE_ROOT)).strip() or str(LABEL_SHARE_ROOT)
        )
    except ValueError as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 2

    options = DownloadOptions(
        month=month,
        share_root=share_root,
        out_dir_override=out_override,
        dry_run=bool(args.dry_run),
        overwrite=bool(args.overwrite),
        write_back=not args.no_write_back,
        account=args.account,
        password=args.password,
    )

    try:
        workbook_id = (args.workbook_id or WORKBOOK_ID).strip()
        if not workbook_id:
            print("[FAIL] 未指定 workbookId", file=sys.stderr)
            return 2

        print(
            f"[CFG] workbook_id={workbook_id} "
            f"label_sheet={args.sheet or LABEL_SHEET} "
            f"label_share_root={share_root} "
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

        sheet = args.sheet or LABEL_SHEET
        df = wb.read_sheet(sheet)
        print(
            f"[READ] workbook={workbook_id} sheet={sheet} "
            f"rows={len(df)} cols={len(df.columns)}"
        )

        for required in REQUIRED_LABEL_COLS:
            if not has_col(df, required):
                print(f"[FAIL] 表格缺少列: [{required}]", file=sys.stderr)
                return 2

        if args.preview is not None:
            n = args.preview if args.preview > 0 else len(df)
            print(df.head(n).fillna("").astype(str).to_string())
            return 0

        rows = dataframe_to_rows(df)
        items, stats = collect_download_items(
            rows,
            only_codes=args.return_codes,
            force=bool(args.force),
        )
        if args.limit is not None and args.limit >= 0:
            items = items[: args.limit]
            stats.codes = len(items)

        if not items:
            print(
                f"[DONE] 无可下载单号 sheet_rows={stats.sheet_rows} "
                f"skipped_has_path={stats.skipped_has_path} "
                f"skipped_empty_code={stats.skipped_empty_code}"
            )
            return 0

        print(
            f"[PLAN] month={options.month} "
            f"base={month_base_dir(month=options.month, share_root=options.share_root)} "
            f"codes={len(items)} skipped_has_path={stats.skipped_has_path}"
        )

        stats, successes = download_items(
            HyOmsWebSession.from_config(), items, options, stats=stats
        )

        if options.write_back and (successes or options.dry_run):
            to_sync = successes
            if options.dry_run:
                to_sync = [
                    DownloadOk(
                        item=it,
                        path=options.dir_for(it.shop_name) / _safe_stem(it.return_code),
                    )
                    for it in items
                ]
            stats.label_updated = sync_label_paths(
                wb, df, to_sync, dry_run=options.dry_run
            )
            stats.progress_updated = sync_register_progress(
                wb, to_sync, dry_run=options.dry_run
            )

        print(
            f"[DONE] ok={stats.ok} fail={stats.fail} codes={stats.codes} "
            f"label_sync={stats.label_updated} "
            f"progress_sync={stats.progress_updated} month={options.month}"
        )
        return 0 if stats.fail == 0 else 1
    except (DingDiskError, HyOmsError, ValueError, KeyError, OSError) as exc:
        print(f"[FAIL] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
