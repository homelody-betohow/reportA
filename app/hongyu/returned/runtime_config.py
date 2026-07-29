"""退件脚本运行时配置：项目根 bootstrap + returned_config.json + 控制台 UTF-8。

查找 ``config/returned_config.json``（发布布局优先 ``dist/config``）；缺文件则用内置默认。
CLI 显式参数应覆盖本模块返回值（由各脚本 argparse 处理）。
"""
from __future__ import annotations

import importlib.util
import json
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Mapping, Optional

DEFAULT_WORKBOOK_ID = "EpGBa2Lm8aDaZ57lTwEMk9boJgN7R35y"
DEFAULT_LABEL_SHARE_ROOT = r"\\Betohow\数据报表\RPA\退货标签"
DEFAULT_REGISTER_SHEET = "退件登记表"
DEFAULT_LABEL_SHEET = "退件标签"
DEFAULT_WAREHOUSE_CODE = "DEHY"
DEFAULT_SM_CODE_DE = "DEGLS-RMA"
DEFAULT_SM_CODE_OTHER = "DEDHL-RMA"
DEFAULT_OPERATION_DESC = "买家退件，检查换标上架"

_CONFIG_NAME = "returned_config.json"


@dataclass(frozen=True)
class ReturnedConfig:
    workbook_id: str = DEFAULT_WORKBOOK_ID
    label_share_root: str = DEFAULT_LABEL_SHARE_ROOT
    register_sheet: str = DEFAULT_REGISTER_SHEET
    label_sheet: str = DEFAULT_LABEL_SHEET
    warehouse_code: str = DEFAULT_WAREHOUSE_CODE
    sm_code_de: str = DEFAULT_SM_CODE_DE
    sm_code_other: str = DEFAULT_SM_CODE_OTHER
    operation_desc: str = DEFAULT_OPERATION_DESC
    # run_task：各步执行前等待秒数，如 {"update_return": 35, "download_label": 4}
    step_delays: Mapping[str, float] = field(default_factory=dict)
    config_path: Optional[Path] = None

    def summary_line(self) -> str:
        path = str(self.config_path) if self.config_path else "(defaults)"
        delays = (
            ",".join(f"{k}={v:g}" for k, v in sorted(self.step_delays.items()))
            if self.step_delays
            else "-"
        )
        return (
            f"[CFG] workbook_id={self.workbook_id} "
            f"label_share_root={self.label_share_root} "
            f"register_sheet={self.register_sheet} "
            f"label_sheet={self.label_sheet} "
            f"warehouse_code={self.warehouse_code} "
            f"sm_code_de={self.sm_code_de} "
            f"sm_code_other={self.sm_code_other} "
            f"operation_desc={self.operation_desc!r} "
            f"step_delays={delays} "
            f"source={path}"
        )


def configure_stdio_utf8() -> None:
    """Windows 控制台中文：尽量把 stdout/stderr 设为 UTF-8。"""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def project_root() -> Path:
    """源码 → 仓库根；冻结 → exe 所在模块目录（如 dist/returned）。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    for p in Path(__file__).resolve().parents:
        if (p / "ensure_project_root.py").is_file():
            return p
    return Path(__file__).resolve().parents[3]


def bootstrap_project_root(_caller_file=None) -> Path:
    """UTF-8 控制台 + 把项目根加入 sys.path（源码还加载 ensure_project_root）。"""
    configure_stdio_utf8()

    if getattr(sys, "frozen", False):
        root = Path(sys.executable).resolve().parent
        root_s = str(root)
        if root_s not in sys.path:
            sys.path.insert(0, root_s)
        return root

    epr_file = next(
        p / "ensure_project_root.py"
        for p in Path(__file__).resolve().parents
        if (p / "ensure_project_root.py").is_file()
    )
    spec = importlib.util.spec_from_file_location("ensure_project_root", epr_file)
    epr_mod = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(epr_mod)
    return epr_mod.bootstrap(_caller_file or __file__)


def resolve_config_path() -> Optional[Path]:
    """返回首个存在的 returned_config.json；皆无则 None。"""
    # 延迟导入，避免源码未 bootstrap 时循环依赖
    from common.dist_paths import resolve_config_file

    return resolve_config_file(_CONFIG_NAME)


def _str_field(data: Mapping[str, Any], key: str, default: str) -> str:
    raw = data.get(key, default)
    text = str(raw).strip() if raw is not None else ""
    return text or default


def _parse_step_delays(raw: Any) -> Dict[str, float]:
    """解析 ``step_delays``：``{"update_return": 35}`` → float 秒数。"""
    if raw is None:
        return {}
    if not isinstance(raw, Mapping):
        raise ValueError("step_delays 须为对象，如 {\"update_return\": 35}")
    out: Dict[str, float] = {}
    for key, value in raw.items():
        name = str(key).strip()
        if not name:
            continue
        try:
            sec = float(value)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"step_delays[{name!r}] 须为数字秒: {value!r}") from exc
        if sec < 0:
            raise ValueError(f"step_delays[{name!r}] 不能为负: {sec}")
        out[name] = sec
    return out


def load_returned_config(config_path: Path | str | None = None) -> ReturnedConfig:
    """加载配置；文件缺失或字段缺失时回退内置默认。"""
    path = Path(config_path) if config_path else resolve_config_path()
    if path is None or not path.is_file():
        return ReturnedConfig()

    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError(f"returned_config 须为 JSON 对象: {path}")

    return ReturnedConfig(
        workbook_id=_str_field(data, "workbook_id", DEFAULT_WORKBOOK_ID),
        label_share_root=_str_field(data, "label_share_root", DEFAULT_LABEL_SHARE_ROOT),
        register_sheet=_str_field(data, "register_sheet", DEFAULT_REGISTER_SHEET),
        label_sheet=_str_field(data, "label_sheet", DEFAULT_LABEL_SHEET),
        warehouse_code=_str_field(data, "warehouse_code", DEFAULT_WAREHOUSE_CODE),
        sm_code_de=_str_field(data, "sm_code_de", DEFAULT_SM_CODE_DE),
        sm_code_other=_str_field(data, "sm_code_other", DEFAULT_SM_CODE_OTHER),
        operation_desc=_str_field(data, "operation_desc", DEFAULT_OPERATION_DESC),
        step_delays=_parse_step_delays(data.get("step_delays")),
        config_path=path,
    )


def ensure_returned_dir_on_path(caller_file: Path | str) -> Path:
    """把 ``app/hongyu/returned`` 加入 sys.path，便于 ``import runtime_config``。"""
    returned_dir = Path(caller_file).resolve().parent
    returned_s = str(returned_dir)
    if returned_s not in sys.path:
        sys.path.insert(0, returned_s)
    return returned_dir


def init_script(caller_file: Path | str) -> tuple[Path, ReturnedConfig]:
    """脚本入口：returned 目录上 path → bootstrap → 加载 returned_config。"""
    ensure_returned_dir_on_path(caller_file)
    root = bootstrap_project_root(caller_file)
    return root, load_returned_config()
