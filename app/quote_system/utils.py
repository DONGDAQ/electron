"""共享工具函数：日期解析、日志保存、API 异常处理装饰器等"""
from __future__ import annotations

import json
import traceback
from datetime import date, datetime
from functools import wraps
from pathlib import Path

from flask import jsonify, Response


def api_handler(fn):
    """
    API 路由异常处理装饰器。
    自动捕获异常、打印 traceback、返回标准错误 JSON。
    适用于：成功时返回 jsonify({"status": "success", ...}) 的路由。
    """
    @wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except Exception as exc:
            traceback.print_exc()
            return jsonify({"status": "error", "message": str(exc)}), 500
    return wrapper


def clean_optional(value: str | None) -> str | None:
    """去除空白，空字符串视为 None"""
    if value is None:
        return None
    value = value.strip()
    return value or None


def parse_date(value: str | None) -> date | None:
    """解析日期字符串，支持 YYYY-MM-DD 和 YYYY/MM/DD 格式"""
    value = clean_optional(value)
    if not value:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"日期格式不正确: {value}，请使用 YYYY-MM-DD")


def save_auto_quote_log(log_dir: Path, project_key: str, output: str, status: str) -> None:
    """保存自动报价执行日志，同时维护 _index.json 摘要索引（最多保留 500 条）"""
    log_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"{project_key}_{ts}.log"
    log_file.write_text(output, encoding="utf-8")

    index_file = log_dir / "_index.json"
    if index_file.exists():
        try:
            index = json.loads(index_file.read_text(encoding="utf-8"))
        except Exception:
            index = []
    else:
        index = []

    index.append({
        "project": project_key,
        "timestamp": datetime.now().isoformat(),
        "status": status,
        "log_file": str(log_file),
        "summary": output.strip().split("\n")[-1] if output.strip() else "",
    })
    if len(index) > 500:
        index = index[-500:]
    index_file.write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")
