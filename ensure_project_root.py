"""将本仓库根目录加入 sys.path，供 from config / from common 导入。

各子脚本在 import config/common 之前通过 importlib 加载本文件并调用 bootstrap()。
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent


def bootstrap(_caller_file=None) -> Path:
    """把项目根加入 sys.path，供 from config... / from common... 导入。"""
    root = str(_ROOT)
    if root not in sys.path:
        sys.path.insert(0, root)
    return _ROOT
