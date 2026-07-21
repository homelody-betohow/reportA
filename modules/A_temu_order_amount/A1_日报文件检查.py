"""
A0_日报文件检查

检查网络盘日报源文件是否齐全，若存在则复制到桌面报表周期目录（如：日报7.1-7.4）。

源目录一：\\\\Betohow\\数据报表\\报表自动化下载\\其它报表\\每天
  - ERP订单、RMA下载
  - transaction交易明细
  - 鸿羽仓二次上架明细
  - 亚马逊利润报表

源目录二：\\\\Betohow\\数据报表\\报表自动化下载\\仓租下载\\每天
  - 4px法国仓仓租下载
  - 鸿羽仓

源目录三：\\\\Betohow\\数据报表\\报表自动化下载\\广告下载\\每天
  - OTTO / Real / DLZ / Mano（子目录 M.D）
  - 测评表（文件名含 report_date）

源目录四：\\\\Betohow\\数据报表\\报表自动化下载\\秒杀下载\\每天
  - 子目录 report_date（YYYY-MM-DD，旧版为 M.D）

日期规则与 A0_set_date.py 一致：
  - ERP / 亚马逊利润 / 二次上架 / 广告 / 秒杀：以 report_date（今天-3天）定位子目录
  - 仓租：子目录 report_date（YYYY-MM-DD，旧版为 M.D）
  - transaction：以当天子目录 + transaction_date 文件名
  - 测评表：以 report_date 的 YYYY-MM-DD 定位文件名
"""
from __future__ import annotations

import argparse
import importlib.util
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

# 须在 import config/common 之前：加载项目根到 sys.path（逻辑见项目根 ensure_project_root.py）
_epr_file = next(
    p / "ensure_project_root.py"
    for p in Path(__file__).resolve().parents
    if (p / "ensure_project_root.py").is_file()
)
_spec = importlib.util.spec_from_file_location("ensure_project_root", _epr_file)
_epr_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_epr_mod)
_epr_mod.bootstrap(__file__)

from config.A0_paths import (  # noqa: E402
    REPORT_PERIOD_DIR,
    SELLERSKU_PROFIT_FILE_NAME,
)
from config.A0_set_date import (  # noqa: E402
    folder_name,
    report_date,
    shared_date,
    transaction_date,
)
from common.runall_utils import setup_console_encoding  # noqa: E402
from common.style import Color  # noqa: E402

SOURCE_BASE = Path(r"\\Betohow\数据报表\报表自动化下载\其它报表\每天")
RENT_SOURCE_BASE = Path(r"\\Betohow\数据报表\报表自动化下载\仓租下载\每天")
AD_SOURCE_BASE = Path(r"\\Betohow\数据报表\报表自动化下载\广告下载\每天")
FLASH_SOURCE_BASE = Path(r"\\Betohow\数据报表\报表自动化下载\秒杀下载\每天")

ERP_SOURCE_DIR = SOURCE_BASE / "ERP订单、RMA下载"
TRANSACTION_SOURCE_DIR = SOURCE_BASE / "transaction交易明细"
RELISTING_SOURCE_DIR = SOURCE_BASE / "鸿羽仓二次上架明细"
AMAZON_PROFIT_SOURCE_DIR = SOURCE_BASE / "亚马逊利润报表"

FPX_RENT_SOURCE_DIR = RENT_SOURCE_BASE / "4px法国仓仓租下载"
HY_RENT_SOURCE_DIR = RENT_SOURCE_BASE / "鸿羽仓"


@dataclass(frozen=True)
class FileCheckItem:
    label: str
    source_dir: Path
    dest_dir: Path
    dest_name: str
    patterns: tuple[str, ...]
    prefer_contains: str = ""
    must_contain: str = ""
    must_not_contain: str = ""
    copy_all: bool = False
    copy_recursive: bool = False


def _is_valid_file(path: Path) -> bool:
    return path.is_file() and not path.name.startswith("~$")


def _pick_source_file(
    directory: Path,
    patterns: tuple[str, ...],
    *,
    prefer_contains: str = "",
    must_contain: str = "",
    must_not_contain: str = "",
) -> Path | None:
    """在目录中按 glob 查找源文件，优先匹配包含指定日期片段的文件名。"""
    if not directory.is_dir():
        return None

    candidates: list[Path] = []
    for pattern in patterns:
        candidates.extend(p for p in directory.glob(pattern) if _is_valid_file(p))

    if not candidates:
        return None

    unique = sorted({p.resolve() for p in candidates}, key=lambda p: p.name)
    if must_contain:
        unique = [p for p in unique if must_contain in p.name]
    if must_not_contain:
        unique = [p for p in unique if must_not_contain not in p.name]
    if not unique:
        return None

    if prefer_contains:
        matched = [p for p in unique if prefer_contains in p.name]
        if matched:
            return matched[0]
    return unique[0]


def _pick_all_source_files(
    directory: Path,
    patterns: tuple[str, ...],
    *,
    prefer_contains: str = "",
    must_contain: str = "",
    must_not_contain: str = "",
) -> list[Path]:
    """在目录中查找全部匹配文件（用于鸿羽仓租等多文件场景）。"""
    if not directory.is_dir():
        return []

    candidates: list[Path] = []
    for pattern in patterns:
        candidates.extend(p for p in directory.glob(pattern) if _is_valid_file(p))

    unique = sorted({p.resolve() for p in candidates}, key=lambda p: p.name)
    if must_contain:
        unique = [p for p in unique if must_contain in p.name]
    if must_not_contain:
        unique = [p for p in unique if must_not_contain not in p.name]
    if prefer_contains:
        matched = [p for p in unique if prefer_contains in p.name]
        if matched:
            return matched
    return unique


def _pick_recursive_source_files(
    directory: Path,
    patterns: tuple[str, ...],
    *,
    prefer_contains: str = "",
) -> list[Path]:
    """递归查找子目录中的匹配文件（用于 MANO 等多级目录广告 CSV）。"""
    if not directory.is_dir():
        return []

    candidates: list[Path] = []
    for pattern in patterns:
        candidates.extend(p for p in directory.rglob(pattern) if _is_valid_file(p))

    unique = sorted({p.resolve() for p in candidates}, key=lambda p: str(p))
    if prefer_contains:
        matched = [p for p in unique if prefer_contains in p.name]
        if matched:
            return matched
    return unique


def _build_other_report_items(period_dir: Path, report_iso: str, report_md: str, today_iso: str) -> list[FileCheckItem]:
    return [
        FileCheckItem(
            label="订单统计",
            source_dir=ERP_SOURCE_DIR / report_iso,
            dest_dir=period_dir / "订单统计",
            dest_name=f"订单统计-{shared_date}.xlsx",
            patterns=("订单统计*.xlsx", "*订单统计*.xlsx"),
            prefer_contains=shared_date,
        ),
        FileCheckItem(
            label="RMA退款",
            source_dir=ERP_SOURCE_DIR / report_iso,
            dest_dir=period_dir / "RMA",
            dest_name=f"RMA-{shared_date}.xlsx",
            patterns=("RMA*.xlsx",),
            prefer_contains=shared_date,
        ),
        FileCheckItem(
            label="transaction-已发放",
            source_dir=TRANSACTION_SOURCE_DIR / today_iso,
            dest_dir=period_dir / "transaction交易明细",
            dest_name=f"transaction交易明细-已发放订单{transaction_date}.xlsx",
            patterns=(
                f"transaction交易明细-已发放订单{transaction_date}.xlsx",
                "transaction交易明细-已发放订单*.xlsx",
            ),
            prefer_contains=transaction_date,
            must_contain="已发放",
        ),
        FileCheckItem(
            label="transaction-已推迟",
            source_dir=TRANSACTION_SOURCE_DIR / today_iso,
            dest_dir=period_dir / "transaction交易明细",
            dest_name=f"transaction交易明细-已推迟订单{transaction_date}.xlsx",
            patterns=(
                f"transaction交易明细-已推迟订单{transaction_date}.xlsx",
                "transaction交易明细-已推迟订单*.xlsx",
            ),
            prefer_contains=transaction_date,
            must_contain="已推迟",
        ),
        FileCheckItem(
            label="鸿羽仓二次上架",
            source_dir=RELISTING_SOURCE_DIR / report_md,
            dest_dir=period_dir / "二次上架",
            dest_name=f"鸿羽仓-二次上架明细-{shared_date}.xls",
            patterns=(
                f"*二次上架明细-{shared_date}.xls",
                f"*二次上架明细-{shared_date}.xlsx",
                "*二次上架明细-*.xls",
                "*二次上架明细-*.xlsx",
            ),
            prefer_contains=shared_date,
        ),
        FileCheckItem(
            label="SellerSku利润报表",
            source_dir=AMAZON_PROFIT_SOURCE_DIR / report_iso,
            dest_dir=period_dir / "SellerSku利润报表",
            dest_name=SELLERSKU_PROFIT_FILE_NAME,
            patterns=(
                f"SellerSku利润报表-{shared_date}.xlsx",
                f"SellerSku利润报表{shared_date}.xlsx",
                "SellerSku利润报表-*.xlsx",
                "SellerSku利润报表*.xlsx",
            ),
            prefer_contains=shared_date,
        ),
    ]


def _build_ad_check_items(period_dir: Path, report_iso: str, report_md: str) -> list[FileCheckItem]:
    return [
        FileCheckItem(
            label="OTTO广告",
            source_dir=AD_SOURCE_BASE / "OTTO" / report_md,
            dest_dir=period_dir / "广告" / "OTTO",
            dest_name=f"OTTO-广告数据-{shared_date}.csv",
            patterns=(
                f"OTTO-广告数据-{shared_date}.csv",
                "OTTO-广告数据-*.csv",
            ),
            prefer_contains=shared_date,
        ),
        FileCheckItem(
            label="REAL广告",
            source_dir=AD_SOURCE_BASE / "Real" / report_md,
            dest_dir=period_dir / "广告" / "REAL",
            dest_name="",
            patterns=(f"*{shared_date}.csv", "REAL*.csv"),
            prefer_contains=shared_date,
            copy_all=True,
        ),
        FileCheckItem(
            label="DLZ广告",
            source_dir=AD_SOURCE_BASE / "DLZ" / report_md,
            dest_dir=period_dir / "广告" / "DLZ",
            dest_name="",
            patterns=(f"*{shared_date}.csv", "DLZ*.csv"),
            prefer_contains=shared_date,
            copy_all=True,
        ),
        FileCheckItem(
            label="MANO广告",
            source_dir=AD_SOURCE_BASE / "Mano" / report_md,
            dest_dir=period_dir / "广告" / "MANO",
            dest_name="",
            patterns=(f"*{shared_date}.csv",),
            prefer_contains=shared_date,
            copy_recursive=True,
        ),
        FileCheckItem(
            label="测评表",
            source_dir=AD_SOURCE_BASE / "测评表",
            dest_dir=period_dir / "测评表",
            dest_name=f"测评表{report_iso}.xlsx",
            patterns=(
                f"测评表{report_iso}.xlsx",
                "测评表*.xlsx",
            ),
            prefer_contains=report_iso,
        ),
    ]


def _resolve_flash_source_dir(report_iso: str, report_md: str) -> Path:
    """秒杀目录优先 YYYY-MM-DD，兼容旧版 M.D 子目录。"""
    iso_dir = FLASH_SOURCE_BASE / report_iso
    if iso_dir.is_dir():
        return iso_dir
    return FLASH_SOURCE_BASE / report_md


def _build_flash_check_items(period_dir: Path, report_iso: str, report_md: str) -> list[FileCheckItem]:
    return [
        FileCheckItem(
            label="秒杀数据",
            source_dir=_resolve_flash_source_dir(report_iso, report_md),
            dest_dir=period_dir / "秒杀",
            dest_name=f"秒杀数据-{shared_date}.xlsx",
            patterns=(
                f"秒杀数据-{shared_date}.xlsx",
                "秒杀数据-*.xlsx",
            ),
            prefer_contains=shared_date,
            must_not_contain="副本",
        ),
    ]


def _resolve_rent_source_dir(base_dir: Path, report_iso: str, report_md: str) -> Path:
    """仓租目录优先 YYYY-MM-DD，兼容旧版 M.D 子目录。"""
    iso_dir = base_dir / report_iso
    if iso_dir.is_dir():
        return iso_dir
    return base_dir / report_md


def _build_rent_check_items(period_dir: Path, report_iso: str, report_md: str) -> list[FileCheckItem]:
    return [
        FileCheckItem(
            label="4PX法国仓仓租",
            source_dir=_resolve_rent_source_dir(FPX_RENT_SOURCE_DIR, report_iso, report_md),
            dest_dir=period_dir / "仓租" / "4PX",
            dest_name=f"4PX法国仓-仓租明细-{shared_date}.xlsx",
            patterns=(
                f"4PX法国仓-仓租明细-{shared_date}.xlsx",
                "4PX法国仓-仓租明细-*.xlsx",
                "4PX*仓租明细*.xlsx",
            ),
            prefer_contains=shared_date,
        ),
        FileCheckItem(
            label="鸿羽仓仓租",
            source_dir=_resolve_rent_source_dir(HY_RENT_SOURCE_DIR, report_iso, report_md),
            dest_dir=period_dir / "仓租" / "鸿羽",
            dest_name="",
            patterns=(
                f"*{shared_date}.xlsx",
                "鸿羽*仓租明细*.xlsx",
            ),
            prefer_contains=shared_date,
            copy_all=True,
        ),
    ]


def _build_check_items() -> list[FileCheckItem]:
    report_iso = report_date.strftime("%Y-%m-%d")
    report_md = f"{report_date.month}.{report_date.day}"
    today_iso = datetime.today().strftime("%Y-%m-%d")
    period_dir = Path(REPORT_PERIOD_DIR)
    return (
        _build_other_report_items(period_dir, report_iso, report_md, today_iso)
        + _build_rent_check_items(period_dir, report_iso, report_md)
        + _build_ad_check_items(period_dir, report_iso, report_md)
        + _build_flash_check_items(period_dir, report_iso, report_md)
    )


def _resolve_relisting_dest_name(source: Path) -> str:
    """源文件多为「鸿羽-」，桌面下游脚本使用「鸿羽仓-」。"""
    suffix = source.suffix.lower()
    if suffix not in {".xls", ".xlsx"}:
        suffix = ".xls"
    return f"鸿羽仓-二次上架明细-{shared_date}{suffix}"


def _check_and_copy(item: FileCheckItem, *, dry_run: bool) -> tuple[bool, str | list[str]]:
    if item.copy_recursive:
        sources = _pick_recursive_source_files(
            item.source_dir,
            item.patterns,
            prefer_contains=item.prefer_contains,
        )
        if not sources:
            if not item.source_dir.is_dir():
                return False, f"源目录不存在：{item.source_dir}"
            return False, f"未找到匹配文件（{', '.join(item.patterns)}）"

        messages: list[str] = []
        for source in sources:
            rel_path = source.relative_to(item.source_dir)
            dest_path = item.dest_dir / rel_path
            if dry_run:
                messages.append(f"[试运行] {rel_path} -> {dest_path}")
            else:
                dest_path.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source, dest_path)
                messages.append(f"已复制：{rel_path} -> {dest_path}")
        return True, messages

    if item.copy_all:
        sources = _pick_all_source_files(
            item.source_dir,
            item.patterns,
            prefer_contains=item.prefer_contains,
            must_contain=item.must_contain,
            must_not_contain=item.must_not_contain,
        )
        if not sources:
            if not item.source_dir.is_dir():
                return False, f"源目录不存在：{item.source_dir}"
            return False, f"未找到匹配文件（{', '.join(item.patterns)}）"

        messages: list[str] = []
        for source in sources:
            dest_path = item.dest_dir / source.name
            if dry_run:
                messages.append(f"[试运行] {source.name} -> {dest_path}")
            else:
                item.dest_dir.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source, dest_path)
                messages.append(f"已复制：{source.name} -> {dest_path}")
        return True, messages

    source = _pick_source_file(
        item.source_dir,
        item.patterns,
        prefer_contains=item.prefer_contains,
        must_contain=item.must_contain,
        must_not_contain=item.must_not_contain,
    )

    dest_name = item.dest_name
    if item.label == "鸿羽仓二次上架" and source is not None:
        dest_name = _resolve_relisting_dest_name(source)

    dest_path = item.dest_dir / dest_name

    if source is None:
        if not item.source_dir.is_dir():
            return False, f"源目录不存在：{item.source_dir}"
        return False, f"未找到匹配文件（{', '.join(item.patterns)}）"

    if dry_run:
        return True, f"[试运行] {source.name} -> {dest_path}"

    item.dest_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, dest_path)
    return True, f"已复制：{source.name} -> {dest_path}"


def main(argv: list[str] | None = None) -> int:
    setup_console_encoding()

    parser = argparse.ArgumentParser(
        description="检查网络盘日报源文件并复制到桌面报表目录"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="仅检查并打印将要复制的文件，不实际复制",
    )
    args = parser.parse_args(argv)

    if folder_name != "日报":
        print(
            f"{Color.RED}当前 A0_set_date.folder_name={folder_name!r}，"
            f"本脚本仅适用于「日报」模式。{Color.RESET}"
        )
        return 1

    period_dir = Path(REPORT_PERIOD_DIR)
    print(f"{Color.CYAN}{'=' * 72}{Color.RESET}")
    print(f"{Color.BOLD}日报文件检查{Color.RESET}")
    print(f"统计区间：{Color.YELLOW}{shared_date}{Color.RESET}")
    print(f"report_date：{report_date.strftime('%Y-%m-%d')}（ERP / 二次上架 / 利润报表 / 仓租 / 广告 / 秒杀）")
    print(f"transaction_date：{Color.YELLOW}{transaction_date}{Color.RESET}")
    print(f"目标目录：{Color.YELLOW}{period_dir}{Color.RESET}")
    print(f"源根目录（其它报表）：{SOURCE_BASE}")
    print(f"源根目录（仓租下载）：{RENT_SOURCE_BASE}")
    print(f"源根目录（广告下载）：{AD_SOURCE_BASE}")
    print(f"源根目录（秒杀下载）：{FLASH_SOURCE_BASE}")
    print(f"{Color.CYAN}{'=' * 72}{Color.RESET}\n")

    if (
        not SOURCE_BASE.is_dir()
        and not RENT_SOURCE_BASE.is_dir()
        and not AD_SOURCE_BASE.is_dir()
        and not FLASH_SOURCE_BASE.is_dir()
    ):
        print(
            f"{Color.RED}无法访问网络源目录：{SOURCE_BASE}、{RENT_SOURCE_BASE}、"
            f"{AD_SOURCE_BASE} 或 {FLASH_SOURCE_BASE}"
            f"\n请确认已连接 \\\\Betohow 共享盘。{Color.RESET}"
        )
        return 1

    items = _build_check_items()
    ok_count = 0
    missing: list[tuple[str, str]] = []

    for item in items:
        print(f"【{item.label}】")
        print(f"  源目录：{item.source_dir}")
        success, message = _check_and_copy(item, dry_run=args.dry_run)
        if success:
            ok_count += 1
            if isinstance(message, list):
                for line in message:
                    print(f"  {Color.GREEN}{line}{Color.RESET}")
            else:
                print(f"  {Color.GREEN}{message}{Color.RESET}")
        else:
            assert isinstance(message, str)
            missing.append((item.label, message))
            print(f"  {Color.RED}{message}{Color.RESET}")
        print()

    print(f"{Color.CYAN}{'=' * 72}{Color.RESET}")
    print(f"检查完成：{ok_count}/{len(items)} 项就绪")
    if missing:
        print(f"{Color.RED}缺少 {len(missing)} 项：{Color.RESET}")
        for label, msg in missing:
            print(f"  - {label}：{msg}")
        print(
            f"\n{Color.YELLOW}请到 ERP 确认上述文件是否已自动生成；"
            f"生成后可重新运行本脚本。{Color.RESET}"
        )
        return 1

    action = "（试运行，未复制）" if args.dry_run else "并已复制到桌面"
    print(f"{Color.GREEN}全部文件已找到{action}。{Color.RESET}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
