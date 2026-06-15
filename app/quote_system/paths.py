import json
import os
from pathlib import Path

_CONFIG_DIR = Path(__file__).resolve().parent.parent / "config"
_FACTORY_CONFIG = _CONFIG_DIR / "base_paths.json"       # 出厂默认 — 只读，永不修改
_USER_CONFIG = _CONFIG_DIR / "base_paths.user.json"      # 用户覆盖 — 运行时写入


def _load_config() -> dict:
    """加载配置：用户覆盖优先，出厂默认兜底。"""
    config = {}

    # 1) 先读出厂默认
    if _FACTORY_CONFIG.exists():
        try:
            config.update(json.loads(_FACTORY_CONFIG.read_text(encoding="utf-8")))
        except Exception:
            pass

    # 2) 再用用户覆盖合并（用户值优先）
    if _USER_CONFIG.exists():
        try:
            config.update(json.loads(_USER_CONFIG.read_text(encoding="utf-8")))
        except Exception:
            pass

    return config


def get_quote_history_dir() -> Path:
    env_val = os.getenv("QUOTE_HISTORY_BASE", "")
    if env_val:
        return Path(env_val)
    config = _load_config()
    if config.get("quote_history_base"):
        return Path(config["quote_history_base"])
    return Path(__file__).resolve().parent.parent / "报价单历史"


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
    _USER_CONFIG.parent.mkdir(parents=True, exist_ok=True)
    _USER_CONFIG.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
