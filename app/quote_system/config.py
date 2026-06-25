from __future__ import annotations

import json
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
