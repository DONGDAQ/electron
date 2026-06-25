from __future__ import annotations

import json
from pathlib import Path

CONFIG_PATH = Path(__file__).parent.parent / "config" / "save_path_config.json"

_config_cache: dict | None = None
_config_cache_mtime: float = 0


def _load_all() -> dict:
    global _config_cache, _config_cache_mtime
    if not CONFIG_PATH.exists():
        return {}
    try:
        mt = CONFIG_PATH.stat().st_mtime
        if _config_cache is not None and mt == _config_cache_mtime:
            return _config_cache
    except Exception:
        mt = 0

    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            _config_cache = json.load(f)
            _config_cache_mtime = mt or CONFIG_PATH.stat().st_mtime
            return _config_cache
    except Exception:
        return {}


def get_save_path(project_key: str) -> str | None:
    """获取指定项目的默认保存路径"""
    return _load_all().get(project_key)


def set_save_path(project_key: str, save_path: str):
    """设置指定项目的默认保存路径"""
    global _config_cache, _config_cache_mtime
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    configs = dict(_load_all())
    configs[project_key] = save_path
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(configs, f, ensure_ascii=False, indent=2)
    _config_cache = configs
    _config_cache_mtime = CONFIG_PATH.stat().st_mtime
