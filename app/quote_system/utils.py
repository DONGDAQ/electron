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


def _extract_summary(output: str, status: str) -> str:
    """从脚本输出中提取统一格式的摘要"""
    if status == "error":
        # 错误：取最后一行
        return output.strip().split("\n")[-1][:100] if output.strip() else "执行失败"

    lines = output.strip().split("\n")
    text = output.strip()

    # TK: "完成: X 条已填写, Y 条跳过, Z 条失败"
    import re
    m = re.search(r'完成:\s*(\d+)\s*条已填写.*?(\d+)\s*条跳过.*?(\d+)\s*条失败', text)
    if m:
        return f"处理 {m.group(1)} 条，跳过 {m.group(2)} 条，失败 {m.group(3)} 条"

    # 4399: "行X K/L/M列已写入: ..."
    m = re.search(r'行(\d+)\s*K/L/M列已写入', text)
    if m:
        count = len(re.findall(r'行\d+\s*K/L/M列已写入', text))
        return f"处理 {count} 条"

    # 完美世界/战双飞书: "总计处理 X 条需求"
    m = re.search(r'总计处理\s*(\d+)\s*条', text)
    if m:
        return f"处理 {m.group(1)} 条"

    # bang2: "找到 X 个待报价文档" 或处理行数
    m = re.search(r'找到\s*(\d+)\s*个待报价', text)
    if m:
        processed = len(re.findall(r'已生成报价单', text))
        if processed > 0:
            return f"处理 {processed} 条"
        return f"找到 {m.group(1)} 条待报价"

    # 没找到需求的各种表述
    if '没有找到' in text or '没有待' in text:
        return "没有待处理需求"

    # 通用：取最后一行
    return lines[-1][:100] if lines else ""


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
        "summary": _extract_summary(output, status),
    })
    if len(index) > 500:
        index = index[-500:]
    index_file.write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")
