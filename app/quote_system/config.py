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


def ensure_config_dir():
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)


def load_language_config(project_key: str) -> LanguageConfig | None:
    ensure_config_dir()
    if not CONFIG_PATH.exists():
        return None
    
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            configs = json.load(f)
        
        if project_key in configs:
            data = configs[project_key]
            return LanguageConfig(
                project_key=project_key,
                languages=data.get("languages", {}),
                default_languages=data.get("default_languages", []),
            )
        return None
    except Exception:
        return None


def save_language_config(project_key: str, languages: dict[str, float], default_languages: list[str]):
    ensure_config_dir()
    
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                configs = json.load(f)
        except Exception:
            configs = {}
    else:
        configs = {}
    
    configs[project_key] = {
        "languages": languages,
        "default_languages": default_languages,
    }
    
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(configs, f, ensure_ascii=False, indent=2)


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
    ensure_config_dir()
    
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                configs = json.load(f)
        except Exception:
            configs = {}
        
        if project_key in configs:
            del configs[project_key]
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(configs, f, ensure_ascii=False, indent=2)
