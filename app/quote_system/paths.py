import json
import os
from pathlib import Path

_CONFIG_DIR = Path(__file__).resolve().parent.parent / "config"
_FACTORY_CONFIG = _CONFIG_DIR / "base_paths.json"       # 出厂默认 — 只读，永不修改
_USER_CONFIG = _CONFIG_DIR / "base_paths.user.json"      # 用户覆盖 — 运行时写入

_config_cache: dict | None = None
_config_cache_mtime: float = 0


def _load_config() -> dict:
    """加载配置：用户覆盖优先，出厂默认兜底。带 mtime 缓存。"""
    global _config_cache, _config_cache_mtime
    current_mtime = 0
    try:
        mtimes = []
        for p in (_FACTORY_CONFIG, _USER_CONFIG):
            if p.exists():
                mtimes.append(p.stat().st_mtime)
        current_mtime = max(mtimes) if mtimes else 0
        if _config_cache is not None and current_mtime == _config_cache_mtime:
            return _config_cache
    except Exception:
        pass

    config = {}
    if _FACTORY_CONFIG.exists():
        try:
            config.update(json.loads(_FACTORY_CONFIG.read_text(encoding="utf-8")))
        except Exception:
            pass
    if _USER_CONFIG.exists():
        try:
            config.update(json.loads(_USER_CONFIG.read_text(encoding="utf-8")))
        except Exception:
            pass

    _config_cache = config
    _config_cache_mtime = current_mtime
    return config


def get_quote_history_dir() -> Path:
    env_val = os.getenv("QUOTE_HISTORY_BASE", "")
    if env_val:
        return Path(env_val)
    config = _load_config()
    if config.get("quote_history_base"):
        return Path(config["quote_history_base"])
    return Path(__file__).resolve().parent.parent / "报价"


def get_settlement_dir() -> Path:
    env_val = os.getenv("SETTLEMENT_BASE", "")
    if env_val:
        return Path(env_val)
    config = _load_config()
    if config.get("settlement_base"):
        return Path(config["settlement_base"])
    return Path(__file__).resolve().parent.parent / "结算"


def save_user_config(data: dict) -> None:
    """保存用户覆盖配置（写入 base_paths.user.json，永不动工厂默认）。"""
    global _config_cache, _config_cache_mtime
    _USER_CONFIG.parent.mkdir(parents=True, exist_ok=True)
    _USER_CONFIG.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    # 更新缓存：合并工厂默认 + 用户覆盖
    config = {}
    if _FACTORY_CONFIG.exists():
        try:
            config.update(json.loads(_FACTORY_CONFIG.read_text(encoding="utf-8")))
        except Exception:
            pass
    config.update(data)
    _config_cache = config
    _config_cache_mtime = _USER_CONFIG.stat().st_mtime
