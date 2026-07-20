"""将本仓库根目录注册为可导入包名 A_报表。

各子脚本在 import A_报表.* 之前通过 importlib 加载本文件并调用 bootstrap()。
目录名可以是 A_报表，也可以是 clone 后的其它名（如 reportA）。
"""
from __future__ import annotations

import sys
import types
from pathlib import Path

PACKAGE_NAME = "A_报表"
_ROOT = Path(__file__).resolve().parent


def bootstrap(_caller_file=None) -> Path:
    """注册项目根为包 A_报表，供 from A_报表... 导入。"""
    root = _ROOT

    if root.name == PACKAGE_NAME:
        parent = str(root.parent)
        if parent not in sys.path:
            sys.path.insert(0, parent)
    else:
        # 目录名不是 A_报表 时，手动挂成同名包
        existing = sys.modules.get(PACKAGE_NAME)
        if existing is None or getattr(existing, "__path__", None) is None:
            pkg = types.ModuleType(PACKAGE_NAME)
            pkg.__file__ = str(root / "__init__.py")
            pkg.__path__ = [str(root)]  # type: ignore[attr-defined]
            pkg.__package__ = PACKAGE_NAME
            sys.modules[PACKAGE_NAME] = pkg

    return root
