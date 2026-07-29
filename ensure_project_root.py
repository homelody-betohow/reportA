"""将本仓库根目录加入 sys.path，供 from config / from common 导入。

各子脚本在 import config/common 之前通过 importlib 加载本文件并调用 bootstrap()。

PyInstaller 冻结后：项目根为可执行文件所在目录（外部 config 放此处）。
"""
from __future__ import annotations

import sys
from pathlib import Path


def get_project_root() -> Path:
    """源码运行 → 仓库根；冻结运行 → exe 所在目录。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def bootstrap(_caller_file=None) -> Path:
    """把项目根加入 sys.path，供 from config... / from common... 导入。"""
    root = get_project_root()
    root_s = str(root)
    if root_s not in sys.path:
        sys.path.insert(0, root_s)
    return root
