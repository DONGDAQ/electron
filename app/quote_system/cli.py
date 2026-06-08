from __future__ import annotations

import argparse
import sys
from datetime import date, datetime
from pathlib import Path

from .generator import QuoteRequest, generate_quote
from .projects import PROJECTS, resolve_project


def main(argv: list[str] | None = None) -> int:
    configure_output()
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        if args.command == "projects":
            for project in PROJECTS.values():
                print(f"{project.display_name}: 默认语种 {', '.join(project.default_languages)}")
            return 0
        if args.command == "generate":
            project = resolve_project(args.project)
            request = QuoteRequest(
                project=project,
                html_paths=[Path(path).resolve() for path in args.html],
                languages=args.language,
                quote_date=parse_date(args.quote_date) or date.today(),
                delivery_date=parse_date(args.delivery_date),
                service_content=args.service_content,
                request_name=args.request_name,
                include_extract=args.extract,
                output_path=Path(args.output).resolve() if args.output else None,
            )
            result = generate_quote(Path.cwd(), request)
            print(f"已生成: {result.output_path}")
            for stats in result.stats:
                print(f"- {stats.source_path.name}: {len(stats.rows)} 行统计")
            return 0
    except Exception as exc:
        print(f"错误: {exc}")
        return 1
    parser.print_help()
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="quote-system", description="翻译报价单生成工具")
    subparsers = parser.add_subparsers(dest="command")

    subparsers.add_parser("projects", help="查看已配置项目")

    generate = subparsers.add_parser("generate", help="根据 MemoQ HTML 生成报价单")
    generate.add_argument("--project", required=True, help="项目名，例如 马娘 / BANG2 / 战双")
    generate.add_argument("--html", nargs="+", required=True, help="一个或多个 MemoQ HTML 文件")
    generate.add_argument("--language", action="append", help="语种/服务，可重复。不填则用项目默认语种")
    generate.add_argument("--extract", action="store_true", help="额外增加一行摘字报价")
    generate.add_argument("--quote-date", help="报价日期，格式 YYYY-MM-DD。不填则用今天")
    generate.add_argument("--delivery-date", help="交付日期，格式 YYYY-MM-DD。不填则写未定")
    generate.add_argument("--service-content", help="报价单 C10 服务内容/文件名")
    generate.add_argument("--request-name", help="输出文件名里的需求名，例如 0416需求")
    generate.add_argument("--output", help="指定输出 xlsx 路径；不填则保存到报价单历史对应项目目录")
    return parser


def parse_date(value: str | None) -> date | None:
    if not value:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"日期格式不正确: {value}，请使用 YYYY-MM-DD")


def configure_output() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
