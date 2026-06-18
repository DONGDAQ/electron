"""共享工具函数：日期解析、日志保存、API 异常处理装饰器等"""
from __future__ import annotations

import io
import json
import sys
import traceback
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from functools import wraps
from pathlib import Path


def configure_output() -> None:
    """重配置 stdout/stderr 编码为 UTF-8"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (OSError, AttributeError):
            pass


def calc_billable(stats, lang: str = "") -> int:
    """计算战双报价字数：85%以上匹配不计费，其余全量计费。"""
    zero_types = {"x-translated / double context", "repetition", "101%",
                  "100%", "95%-99%", "context", "exact"}
    is_en = lang == "英-韩"
    total = 0
    for row in stats.rows:
        t = str(row.type).strip().lower() if row.type else ""
        if t in zero_types or t in ("all",):
            continue
        if is_en:
            total += row.source_non_asian_words or 0
        else:
            total += row.source_asian_characters or 0
    return total


def validate_year_month(year: int, month: int) -> None:
    """校验年月参数，无效时抛出 BadRequest"""
    if not (2020 <= year <= 2099 and 1 <= month <= 12):
        raise BadRequest("年月参数无效")


class BadRequest(ValueError):
    """400 错误，api_handler 捕获后返回 400 而非 500"""
    pass


@contextmanager
def capture_stdout():
    """捕获 stdout 输出的上下文管理器"""
    old_stdout = sys.stdout
    sys.stdout = buffer = io.StringIO()
    try:
        yield buffer
    finally:
        sys.stdout = old_stdout


def excel_serial_to_date(serial, *, as_date: bool = False) -> datetime | date | None:
    """Excel 日期序列号转日期。as_date=True 返回 date，否则返回 datetime。"""
    try:
        num = float(serial)
    except (TypeError, ValueError):
        return None
    base = datetime(1899, 12, 30)
    try:
        result = base + timedelta(days=num)
        return result.date() if as_date else result
    except Exception:
        return None

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
        except BadRequest as exc:
            return jsonify({"status": "error", "message": str(exc)}), 400
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
        filled, skipped, failed = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if filled == 0 and skipped == 0 and failed == 0:
            return "没有待处理需求"
        return f"处理 {filled} 条，跳过 {skipped} 条，失败 {failed} 条"

    # 4399: "行X K/L/M列已写入: ..."
    m = re.search(r'行(\d+)\s*K/L/M列已写入', text)
    if m:
        count = len(re.findall(r'行\d+\s*K/L/M列已写入', text))
        return f"处理 {count} 条"

    # 完美世界/战双飞书: "总计处理 X 条需求"
    m = re.search(r'总计处理\s*(\d+)\s*条', text)
    if m:
        return f"处理 {m.group(1)} 条"

    # 完美世界: "行 X 处理完成!"
    m = re.search(r'行\s*(\d+)\s*处理完成', text)
    if m:
        count = len(re.findall(r'行\s*\d+\s*处理完成', text))
        return f"处理 {count} 条"

    # bang2: "找到 X 个待报价文档" 或处理行数
    m = re.search(r'找到\s*(\d+)\s*个待报价', text)
    if m:
        processed = len(re.findall(r'已生成报价单', text))
        if processed > 0:
            return f"处理 {processed} 条"
        return f"找到 {m.group(1)} 条待报价"

    # 战双: "状态已更新为'已生成'" → 统计更新条数
    m = re.findall(r'状态已更新', text)
    if m:
        return f"更新 {len(m)} 条状态"

    # 没找到需求的各种表述
    if '没有找到' in text or '没有待' in text:
        return "没有待处理需求"

    # 通用：取最后一行
    return lines[-1][:100] if lines else ""


def save_auto_quote_log(log_dir: Path, project_key: str, output: str, status: str, trigger_type: str = "manual") -> None:
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
        "type": trigger_type,
    })
    if len(index) > 500:
        index = index[-500:]
    index_file.write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")
