from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path

from .projects import PROJECTS, ProjectConfig


@dataclass
class LanguageConfig:
    project_key: str
    languages: dict[str, float]
    default_languages: list[str]


CONFIG_PATH = Path(__file__).parent.parent / "config" / "language_config.json"

_config_cache: dict | None = None
_config_cache_mtime: float = 0


def ensure_config_dir():
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)


def _load_all_configs() -> dict:
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


def load_language_config(project_key: str) -> LanguageConfig | None:
    ensure_config_dir()
    configs = _load_all_configs()
    if project_key in configs:
        data = configs[project_key]
        return LanguageConfig(
            project_key=project_key,
            languages=data.get("languages", {}),
            default_languages=data.get("default_languages", []),
        )
    return None


def save_language_config(project_key: str, languages: dict[str, float], default_languages: list[str]):
    global _config_cache, _config_cache_mtime
    ensure_config_dir()

    configs = dict(_load_all_configs())

    configs[project_key] = {
        "languages": languages,
        "default_languages": default_languages,
    }

    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(configs, f, ensure_ascii=False, indent=2)
    _config_cache = configs
    _config_cache_mtime = CONFIG_PATH.stat().st_mtime


def get_effective_language_config(project: ProjectConfig) -> tuple[dict[str, float], list[str]]:
    config = load_language_config(project.key)
    
    if config and config.languages:
        languages = config.languages
        default_langs = config.default_languages if config.default_languages else list(project.default_languages)
    else:
        languages = dict(project.prices)
        default_langs = list(project.default_languages)
    
    return languages, default_langs


def reset_to_default(project_key: str):
    global _config_cache, _config_cache_mtime
    ensure_config_dir()

    configs = dict(_load_all_configs())

    if project_key in configs:
        del configs[project_key]
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(configs, f, ensure_ascii=False, indent=2)
        _config_cache = configs
        _config_cache_mtime = CONFIG_PATH.stat().st_mtime


# ======================== 敏感凭证（credentials.json，已 gitignore） ========================
# 读取顺序：credentials.json 配置文件 > 环境变量 > 默认值。
# 凭证不应硬编码在源码中，如需更换请直接编辑 app/config/credentials.json。

CREDENTIALS_PATH = Path(__file__).parent.parent / "config" / "credentials.json"

_credentials_cache: dict | None = None


def get_credentials() -> dict:
    """读取 credentials.json，失败/缺失返回空 dict。带进程内缓存。"""
    global _credentials_cache
    if _credentials_cache is not None:
        return _credentials_cache
    try:
        if CREDENTIALS_PATH.exists():
            with open(CREDENTIALS_PATH, "r", encoding="utf-8") as f:
                _credentials_cache = json.load(f) or {}
                return _credentials_cache
    except Exception:
        pass
    _credentials_cache = {}
    return _credentials_cache


def get_feishu_credentials() -> tuple[str, str]:
    """返回 (app_id, app_secret)。优先级：配置文件 > 环境变量 > 默认值。"""
    cred = get_credentials().get("feishu", {})
    app_id = cred.get("app_id") or os.environ.get("FEISHU_APP_ID", "cli_aa882f3b1abb5bb7")
    app_secret = cred.get("app_secret") or os.environ.get("FEISHU_APP_SECRET", "")
    return app_id, app_secret


def get_smb_credentials() -> tuple[str, str]:
    """返回 (user, password)。优先级：配置文件 > 环境变量 > 默认值。"""
    cred = get_credentials().get("smb", {})
    user = cred.get("user") or os.environ.get("SMB_USER", "dong_daqian")
    password = cred.get("password") or os.environ.get("SMB_PASSWORD", "")
    return user, password


def get_dingtalk_webhook() -> str:
    """返回钉钉 webhook。优先级：配置文件 > 环境变量。"""
    cred = get_credentials().get("dingtalk", {})
    return cred.get("webhook") or os.environ.get("DINGTALK_WEBHOOK", "")
