import json
import os
from typing import Any, Dict


def load_config(config_path: str) -> Dict[str, Any]:
    """加载 JSON 配置文件，并做最基本的路径归一化。"""
    with open(config_path, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    base_dir = os.path.dirname(os.path.abspath(config_path))

    # docx_path / reports_dir 支持相对路径（相对 config 文件所在目录）
    if "docx_path" in cfg and cfg["docx_path"]:
        if not os.path.isabs(cfg["docx_path"]):
            cfg["docx_path"] = os.path.join(base_dir, cfg["docx_path"])

    if "reports_dir" in cfg and cfg["reports_dir"]:
        if not os.path.isabs(cfg["reports_dir"]):
            cfg["reports_dir"] = os.path.join(base_dir, cfg["reports_dir"])

    return cfg


def get_cfg(cfg: Dict[str, Any], key: str, default: Any = None) -> Any:
    """安全获取配置字段"""
    return cfg.get(key, default) if isinstance(cfg, dict) else default


