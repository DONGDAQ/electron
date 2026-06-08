"""4399 飞书在线表格报价：从wiki内嵌表格读取需求 → 下载HTML → 回填K/L/M列"""
from __future__ import annotations

import argparse
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from quote_system.feishu_client import FeishuClient
from quote_system.memoq_html import parse_memoq_html

# 4399 表格 token
SPREADSHEET_TOKEN = "YcKKsgMtUhrghlt65ubczJ21n9c"
SHEET_ID = "a63cb8"
DOWNLOAD_DIR = ROOT / "outputs" / "4399_html"

# 列索引（0-based）
COL_REQ_NAME = 1       # B: 需求名
COL_PROJECT = 2        # C: 项目名
COL_HTML_ATTACH = 6    # G: 需求统计文件
COL_ASIAN_CHARS = 10   # K: 全source亚洲字符数
COL_TOTAL_CHARS = 11   # L: 全source字符数
COL_REPEAT_CHARS = 12  # M: 重复字符数


def run_scheduled() -> None:
    """供 auto_quote_scheduled.py 调用的无参数入口。"""
    _configure_output()
    run()


def main(argv: list[str] | None = None) -> int:
    _configure_output()
    parser = argparse.ArgumentParser(prog="auto-quote-4399", description="4399 飞书在线表格报价")
    parser.add_argument("--dry-run", action="store_true", help="只扫描不写入")
    args = parser.parse_args(argv)
    try:
        run(dry_run=args.dry_run)
        return 0
    except KeyboardInterrupt:
        print("\n已停止。")
        return 130
    except Exception as exc:
        print(f"错误: {exc}")
        return 1


def run(dry_run: bool = False) -> None:
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] 4399报价自动化开始{' (dry-run)' if dry_run else ''}")

    client = FeishuClient(
        sheet_id=SHEET_ID,
        spreadsheet_token=SPREADSHEET_TOKEN,
    )

    try:
        rows = client.read_sheet()
    except Exception as e:
        print(f"读取飞书表格失败: {e}")
        return

    pending = find_unprocessed_rows(rows)
    if not pending:
        print("没有找到待处理的需求")
        return

    print(f"找到 {len(pending)} 个待处理需求:")
    for it in pending:
        print(f"  行{it['row_num']}: {it['req_name']} [{it['project']}]")

    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

    for item in pending:
        row_num = item["row_num"]
        try:
            print(f"\n处理行 {row_num}: {item['req_name']}")
            html_path = DOWNLOAD_DIR / item["file_name"]
            print(f"  下载HTML: {item['file_name']}")
            client.download_attachment(item["file_token"], html_path)

            stats = parse_memoq_html(html_path)
            all_row = stats.all_row
            rep_row = stats.repetition_row

            asian_chars = all_row.source_asian_characters or 0
            total_chars = all_row.source_chars or 0
            repeat_chars = rep_row.source_asian_characters or 0

            print(f"  亚洲字符={asian_chars}, 总字符={total_chars}, 重复字符={repeat_chars}")

            if dry_run:
                print(f"  [dry-run] 跳过写入")
            else:
                client.write_cell(row_num, COL_ASIAN_CHARS, asian_chars)
                client.write_cell(row_num, COL_TOTAL_CHARS, total_chars)
                client.write_cell(row_num, COL_REPEAT_CHARS, repeat_chars)
                print(f"  行{row_num} K/L/M列已写入: {asian_chars}/{total_chars}/{repeat_chars}")

                try:
                    from settlement.settlement_tracker import quick_record, add_record
                    add_record(quick_record(
                        project_key="4399",
                        company="4399",
                        req_name=item["req_name"],
                        word_count=asian_chars,
                        total_price=0,
                        quote_file="",
                        language="",
                        source="auto",
                        source_chars=total_chars,
                        billable_words=asian_chars,
                    ))
                except Exception:
                    pass

            html_path.unlink(missing_ok=True)

            time.sleep(0.3)

        except Exception as e:
            print(f"处理行 {row_num} 失败: {e}")
            traceback.print_exc()


def find_unprocessed_rows(rows: list[list]) -> list[dict]:
    """找出有HTML附件但K列为空的待处理行"""
    result = []
    for i, row in enumerate(rows):
        if i == 0:
            continue
        row_num = i + 1

        name = str(row[COL_REQ_NAME]).strip() if len(row) > COL_REQ_NAME else ""
        if not name:
            continue

        k_val = row[COL_ASIAN_CHARS] if len(row) > COL_ASIAN_CHARS else None
        if k_val is not None and str(k_val).strip():
            continue

        g_val = row[COL_HTML_ATTACH] if len(row) > COL_HTML_ATTACH else None
        if not g_val or not isinstance(g_val, list) or len(g_val) == 0:
            continue

        attach = g_val[0]
        file_token = attach.get("fileToken")
        file_name = attach.get("text", "")
        if not file_token:
            continue

        result.append({
            "row_num": row_num,
            "req_name": name,
            "project": str(row[COL_PROJECT]).strip() if len(row) > COL_PROJECT else "",
            "file_token": file_token,
            "file_name": file_name,
        })
    return result


def _configure_output() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


if __name__ == "__main__":
    raise SystemExit(main())
