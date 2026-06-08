"""API-based TK O-Y column fill: download N-column HTML attachments, parse, write back."""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path
from typing import Any

from .feishu_client import FeishuClient
from .memoq_html import is_same_as_above, parse_memoq_html, stats_to_output_values


SPREADSHEET_TOKEN = "KUTQsDzIxhg4SltVkCdceeOfnXb"
SHEET_ID = "d73cd4"
DOWNLOAD_DIR = Path(__file__).resolve().parent.parent / "outputs" / "tk_html"


def run_scheduled() -> None:
    """供 auto_quote_scheduled.py 调用的无参数入口。"""
    configure_output()
    parser = build_parser()
    run(parser.parse_args([]))


def main(argv: list[str] | None = None) -> int:
    configure_output()
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        run(args)
        return 0
    except KeyboardInterrupt:
        print("\n已停止。")
        return 130
    except Exception as exc:
        print(f"错误: {exc}")
        return 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="auto-fill-tk", description="API 自动填表 TK O-Y 列")
    parser.add_argument("--dry-run", action="store_true", help="只扫描不写入")
    parser.add_argument("--tail", type=int, default=0, help="只处理最后 N 行")
    parser.add_argument("--blank-stop", type=int, default=30, help="连续空 N 列停止")
    parser.add_argument("--max-rows", type=int, default=2000)
    parser.add_argument("--download-dir", default=str(DOWNLOAD_DIR))
    return parser


def run(args: argparse.Namespace) -> None:
    download_dir = Path(args.download_dir)
    if download_dir.exists():
        for f in download_dir.iterdir():
            if f.is_file():
                f.unlink()
        print("已清理旧 HTML 文件。")
    download_dir.mkdir(parents=True, exist_ok=True)

    client = FeishuClient(spreadsheet_token=SPREADSHEET_TOKEN, sheet_id=SHEET_ID)
    all_rows = client.read_sheet()
    print(f"读取表格: {len(all_rows)} 行")

    candidates = find_candidates(all_rows, args.tail)
    if not candidates:
        print("没有需要处理的行。")
        return

    print(f"找到 {len(candidates)} 行待处理:")
    for c in candidates:
        print(f"  行{c['row']}: {c['label']}")

    processed = 0
    skipped = 0
    failed = 0
    last_values: list[int] | None = None

    for idx, candidate in enumerate(candidates):
        row_num = candidate["row"]
        print(f"\n[{idx+1}/{len(candidates)}] 处理行 {row_num}...")

        try:
            if candidate["is_same_as_above"]:
                if last_values is None:
                    print("  N列为\"同上\"，但上方没有可复用数据，跳过。")
                    skipped += 1
                    continue
                values = last_values
                print(f"  N列为\"同上\"，复用上一条: {values}")
            else:
                file_token = candidate["file_token"]
                html_path = download_html(client, file_token, candidate["file_name"], download_dir, row_num)
                if html_path is None:
                    failed += 1
                    continue

                stats = parse_memoq_html(html_path)
                values = stats_to_output_values(stats)
                last_values = values
                html_path.unlink(missing_ok=True)
                print(f"  解析完成: O={values[0]}, P={values[1]}, Q={values[2]}...")

            if args.dry_run:
                print(f"  [dry-run] 跳过写入: {values}")
            else:
                write_o_y(client, row_num, values)
                print(f"  已写入 O-Y: {values}")

                try:
                    from settlement.settlement_tracker import quick_record, add_record
                    add_record(quick_record(
                        project_key="tk",
                        company="Bilibili",
                        req_name=candidate.get("label", f"行{row_num}"),
                        word_count=values[0] if values else 0,
                        total_price=0,
                        quote_file="",
                        language="",
                        source="auto",
                        source_chars=values[0] if values else 0,
                        billable_words=values[1] if len(values) > 1 else 0,
                    ))
                except Exception:
                    pass

            processed += 1
            time.sleep(0.3)  # 避免 API 限流

        except Exception as exc:
            failed += 1
            print(f"  失败: {exc}")

    print(f"\n完成: 成功 {processed} 行, 跳过 {skipped} 行, 失败 {failed} 行。")


def find_candidates(all_rows: list[list[Any]], tail: int = 0) -> list[dict]:
    """找出需要处理的行（N列有附件且O列为空）。"""
    candidates = []
    start = max(0, len(all_rows) - tail) if tail > 0 else 0
    blank_streak = 0

    for i in range(start, len(all_rows)):
        row = all_rows[i]
        row_num = i + 1  # 1-based row number
        n_val = row[13] if len(row) > 13 else ""
        o_val = row[14] if len(row) > 14 else ""

        # O列已有值则跳过
        if o_val_is_filled(o_val):
            if isinstance(n_val, list):
                blank_streak = 0
            continue

        # N列为空
        if not n_val or (isinstance(n_val, str) and not n_val.strip()):
            blank_streak += 1
            if blank_streak >= 30:
                break
            continue

        blank_streak = 0

        # N列为"同上"
        if isinstance(n_val, str) and is_same_as_above(n_val):
            candidates.append({
                "row": row_num,
                "is_same_as_above": True,
                "file_token": None,
                "file_name": "",
                "label": "同上",
            })
            continue

        # N列为附件
        if isinstance(n_val, list) and len(n_val) > 0:
            file_info = n_val[0]
            file_token = file_info.get("fileToken")
            file_name = file_info.get("text", "")
            if file_token:
                candidates.append({
                    "row": row_num,
                    "is_same_as_above": False,
                    "file_token": file_token,
                    "file_name": file_name,
                    "label": file_name,
                })
                continue

        # 其他文本，跳过
        print(f"  行{row_num}: N列跳过非附件内容: {str(n_val)[:50]}")

    return candidates


def o_val_is_filled(o_val: Any) -> bool:
    """O列是否已经有值。"""
    if o_val is None:
        return False
    if isinstance(o_val, str) and not o_val.strip():
        return False
    if isinstance(o_val, (int, float)):
        return True
    return bool(o_val)


def download_html(client: FeishuClient, file_token: str, file_name: str, download_dir: Path, row_num: int) -> Path | None:
    """下载附件 HTML 文件。"""
    safe_name = "".join(ch if ch not in '<>:"/\\|?*' else "_" for ch in (file_name or f"row_{row_num}.html"))
    save_path = download_dir / f"{row_num}_{safe_name}"

    # 避免重名
    if save_path.exists():
        stem = save_path.stem
        for idx in range(2, 1000):
            alt = save_path.with_name(f"{stem}_{idx}{save_path.suffix}")
            if not alt.exists():
                save_path = alt
                break

    try:
        client.download_attachment(file_token, save_path)
        print(f"  已下载: {save_path.name}")
        return save_path
    except Exception as exc:
        print(f"  下载失败: {exc}")
        return None


def write_o_y(client: FeishuClient, row_num: int, values: list[int]) -> None:
    """写入 O-Y 列（11列）。"""
    col_letters = ["O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y"]
    range_ref = f"{SHEET_ID}!O{row_num}:Y{row_num}"
    row_values = [[v if v is not None else 0 for v in values]]
    client.write_range(range_ref, row_values)


def configure_output() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


if __name__ == "__main__":
    raise SystemExit(main())
