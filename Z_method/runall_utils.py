"""
runAll 脚本的公共工具函数
用于处理 Windows 控制台编码、运行子脚本等通用逻辑
"""

import os
import re
import subprocess
import sys
from pathlib import Path


def setup_console_encoding():
    """
    设置 Windows 控制台编码处理
    避免输出中文时出现乱码（Cursor / VS Code / Python 3.13 默认 UTF-8）
    """
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    if sys.platform == "win32":
        try:
            import ctypes

            kernel32 = ctypes.windll.kernel32
            kernel32.SetConsoleOutputCP(65001)
            kernel32.SetConsoleCP(65001)
        except Exception:
            pass


def _subprocess_env() -> dict[str, str]:
    """子进程统一使用 UTF-8 输出，与 runAll 捕获解码方式一致。"""
    env = os.environ.copy()
    env.setdefault("PYTHONIOENCODING", "utf-8")
    return env


def run_script(script_path: Path) -> tuple[int, str]:
    """
    运行单个 Python 脚本
    
    Args:
        script_path: 脚本文件路径
        
    Returns:
        tuple[int, str]: (退出码, 输出内容)
            - 退出码 0 表示成功
            - 输出内容包含脚本的 stdout 和 stderr
    """
    proc = subprocess.Popen(
        [sys.executable, str(script_path)],
        cwd=str(script_path.parent),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=_subprocess_env(),
    )

    out_lines: list[str] = []
    assert proc.stdout is not None
    for line in proc.stdout:
        print(line, end="")
        out_lines.append(line)

    code = proc.wait()
    return int(code), "".join(out_lines)


def get_scripts_in_order(folder: Path, exclude_filename: str) -> list[Path]:
    """
    获取目录下所有 Python 脚本，并按文件名排序
    
    Args:
        folder: 目标目录
        exclude_filename: 要排除的文件名（通常是 runAll 脚本自身）
        
    Returns:
        list[Path]: 排序后的脚本文件列表
    """
    scripts = [
        p
        for p in folder.glob("*.py")
        if p.is_file() and p.name != exclude_filename
    ]
    scripts.sort(key=lambda p: p.name)
    return scripts


def extract_output_file_path(output: str, default: str = "") -> str:
    """
    从脚本输出中提取生成的文件路径
    查找形如 "output_path：xxx.xlsx" 或 "文件另存为：xxx.xlsx" 的行
    
    Args:
        output: 脚本输出内容
        default: 未找到时的默认值
        
    Returns:
        str: 提取到的文件路径，未找到则返回 default
    """
    # 匹配 output_path:xxx.xlsx 或 文件另存为:xxx.xlsx
    patterns = [
        r'output_path[：:]\s*(.+?\.xlsx)',
        r'文件另存为[：:]\s*(.+?\.xlsx)',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, output)
        if match:
            return match.group(1).strip()
    
    return default
