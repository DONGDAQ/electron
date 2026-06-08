import json
import os
from pathlib import Path

_CONFIG_DIR = Path(__file__).resolve().parent.parent / "config"
_CONFIG_PATH = _CONFIG_DIR / "base_paths.json"

_smb_registered = False


def _ensure_smb():
    global _smb_registered
    if _smb_registered:
        return
    try:
        from smbclient import register_session
        register_session("192.168.110.111", username="dong_daqian", password="dq46460055")
        _smb_registered = True
    except Exception:
        pass


def _load_config() -> dict:
    if _CONFIG_PATH.exists():
        try:
            return json.loads(_CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def get_quote_history_dir() -> Path:
    env_val = os.getenv("QUOTE_HISTORY_BASE", "")
    if env_val:
        return Path(env_val)
    config = _load_config()
    if config.get("quote_history_base"):
        _ensure_smb()
        return Path(config["quote_history_base"])
    return Path(__file__).resolve().parent.parent / "报价单历史"


def get_settlement_dir() -> Path:
    env_val = os.getenv("SETTLEMENT_BASE", "")
    if env_val:
        return Path(env_val)
    config = _load_config()
    if config.get("settlement_base"):
        _ensure_smb()
        return Path(config["settlement_base"])
    return Path(__file__).resolve().parent.parent / "结算"
