from __future__ import annotations

import json
from pathlib import Path

CONFIG_PATH = Path(__file__).parent.parent / "config" / "save_path_config.json"


def get_save_path(project_key: str) -> str | None:
    """获取指定项目的默认保存路径"""
    if not CONFIG_PATH.exists():
        return None
    
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            configs = json.load(f)
        
        return configs.get(project_key)
    except Exception:
        return None


def set_save_path(project_key: str, save_path: str):
    """设置指定项目的默认保存路径"""
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                configs = json.load(f)
        except Exception:
            configs = {}
    else:
        configs = {}
    
    configs[project_key] = save_path
    
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(configs, f, ensure_ascii=False, indent=2)


def clear_save_path(project_key: str):
    """清除指定项目的默认保存路径"""
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                configs = json.load(f)
            
            if project_key in configs:
                del configs[project_key]
                
                with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                    json.dump(configs, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
