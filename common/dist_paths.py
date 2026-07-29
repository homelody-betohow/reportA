"""发布目录布局下的配置路径解析。

推荐布局（多模块共用一份 config）::

    dist/
      config/                 # 共享：db_config / secrets / returned_config ...
      returned/*.exe
      <other_module>/*.exe

冻结运行时查找顺序::

    1. <exe 上级>/config/     # 即 dist/config
    2. <exe 同级>/config/     # 兼容旧布局 dist/returned/config
    3. <cwd>/config/
"""
from __future__ import annotations

import sys
from datetime import date, datetime
from pathlib import Path
from typing import Iterable, List, Optional


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def exe_dir() -> Path:
    """冻结时返回可执行文件所在目录。"""
    return Path(sys.executable).resolve().parent


def source_repo_root(start: Path | None = None) -> Path:
    """自 ``start`` 向上查找含 ``ensure_project_root.py`` 的仓库根。"""
    base = start or Path(__file__).resolve()
    for p in [base, *base.parents]:
        if (p / "ensure_project_root.py").is_file():
            return p
    return Path(__file__).resolve().parents[1]


def config_roots(*, caller_file: Path | str | None = None) -> List[Path]:
    """可能放置 ``config/`` 的根目录（按优先级）。"""
    roots: List[Path] = []
    if is_frozen():
        module_dir = exe_dir()
        dist_root = module_dir.parent
        for r in (dist_root, module_dir):
            if r not in roots:
                roots.append(r)
    else:
        start = Path(caller_file).resolve() if caller_file else Path(__file__).resolve()
        roots.append(source_repo_root(start))
    cwd = Path.cwd()
    if cwd not in roots:
        roots.append(cwd)
    return roots


def iter_config_paths(filename: str, *, caller_file: Path | str | None = None) -> Iterable[Path]:
    for root in config_roots(caller_file=caller_file):
        yield root / "config" / filename


def resolve_config_file(
    filename: str,
    *,
    caller_file: Path | str | None = None,
) -> Optional[Path]:
    """返回首个存在的 ``config/<filename>``；皆无则 None。"""
    for path in iter_config_paths(filename, caller_file=caller_file):
        if path.is_file():
            return path
    return None


def default_config_path(
    filename: str,
    *,
    caller_file: Path | str | None = None,
) -> Path:
    """用于错误提示：优先 dist/config（冻结）或仓库根 config。"""
    roots = config_roots(caller_file=caller_file)
    return roots[0] / "config" / filename


def logs_dir(*, caller_file: Path | str | None = None) -> Path:
    """日志目录：冻结 → ``dist/logs``；源码 → 仓库根 ``logs``。"""
    return config_roots(caller_file=caller_file)[0] / "logs"


def daily_log_path(
    prefix: str,
    *,
    day: date | datetime | None = None,
    caller_file: Path | str | None = None,
) -> Path:
    """``logs/<prefix>_yyyymmdd.log``（自动创建 logs 目录）。"""
    when = day or datetime.now()
    if isinstance(when, datetime):
        stamp = when.strftime("%Y%m%d")
    else:
        stamp = when.strftime("%Y%m%d")
    folder = logs_dir(caller_file=caller_file)
    folder.mkdir(parents=True, exist_ok=True)
    safe = str(prefix).strip() or "app"
    return folder / f"{safe}_{stamp}.log"
