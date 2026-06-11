"""API-based TK O-Y column fill: lark-cli user identity, direct read/write on original sheet."""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

from .memoq_html import is_same_as_above, parse_memoq_html, stats_to_output_values

ORIGINAL_SHEET_URL = "https://my.feishu.cn/sheets/R8QWsav9Nh77V3tWdTQchuCfnMc"
SHEET_ID = "d73cd4"
DOWNLOAD_DIR = Path(__file__).resolve().parent.parent / "outputs" / "tk_html"

import os
LARK_CLI = os.environ.get("LARK_CLI_PATH", "lark-cli")
# Fallback to known install location
if LARK_CLI == "lark-cli":
    known = Path(os.path.expanduser("~/.workbuddy/binaries/node/cli-connector-packages/lark-cli"))
    if known.exists():
        LARK_CLI = str(known)


def _cli(*args: str, timeout: int = 60) -> subprocess.CompletedProcess:
    """Run lark-cli with common flags. Uses bash on Windows for shell scripts."""
    cmd = ["bash", LARK_CLI, "--format", "json"] + list(args)
    return subprocess.run(
        cmd, capture_output=True, text=True, timeout=timeout,
        env={**os.environ, "LARK_CLI_NO_PROXY": "1"},
    )


def run_scheduled() -> None:
    """Called by auto_quote_scheduled.py."""
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
        print("\nStopped.")
        return 130
    except Exception as exc:
        print(f"Error: {exc}")
        return 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="auto-fill-tk", description="TK auto O-Y fill (lark-cli direct)")
    parser.add_argument("--dry-run", action="store_true", help="Scan only, no write")
    parser.add_argument("--tail", type=int, default=0, help="Process last N rows only")
    parser.add_argument("--blank-stop", type=int, default=30, help="Stop after N consecutive blanks")
    parser.add_argument("--max-rows", type=int, default=2000)
    parser.add_argument("--download-dir", default=str(DOWNLOAD_DIR))
    return parser


def run(args: argparse.Namespace) -> None:
    download_dir = Path(args.download_dir)
    if download_dir.exists():
        for f in download_dir.iterdir():
            if f.is_file():
                f.unlink()
        print("Cleaned old HTML files.")
    download_dir.mkdir(parents=True, exist_ok=True)

    all_rows = read_sheet_with_attachments(download_dir)
    print(f"Read sheet: {len(all_rows)} rows")

    candidates = find_candidates(all_rows, args.tail, args.blank_stop)
    if not candidates:
        print("No rows to process.")
        return

    print(f"Found {len(candidates)} rows to process:")
    for c in candidates:
        label = c["n_attach"]["name"] if c["n_attach"] else c["n_text"]
        print(f"  Row {c['row']}: {label}")

    processed = 0
    skipped = 0
    failed = 0
    last_values: list[int] | None = None

    for idx, candidate in enumerate(candidates):
        row_num = candidate["row"]
        print(f"\n[{idx+1}/{len(candidates)}] Processing row {row_num}...")

        try:
            # 纯文本 + "同上" → 复用上一行
            if not candidate["n_attach"] and is_same_as_above(candidate["n_text"]):
                if last_values is None:
                    print("  N='same as above' but no previous data, skipping.")
                    skipped += 1
                    continue
                values = last_values
                print(f"  N='same as above', reusing: {values}")

            # 有附件 → 下载 + 解析
            elif candidate["n_attach"]:
                attach = candidate["n_attach"]
                html_path = download_html(attach["token"], attach["name"], download_dir, row_num)
                if html_path is None:
                    failed += 1
                    continue

                stats = parse_memoq_html(html_path)
                values = stats_to_output_values(stats)
                last_values = values
                html_path.unlink(missing_ok=True)
                print(f"  Parsed: O={values[0]}, P={values[1]}, Q={values[2]}...")

            # 纯文本非"同上" → 跳过
            else:
                print(f"  N column non-attachment, non-same-as-above: {candidate['n_text'][:50]}, skipping.")
                skipped += 1
                continue

            if args.dry_run:
                print(f"  [dry-run] Skip write: {values}")
            else:
                write_o_y(row_num, values)
                print(f"  Written O-Y: {values}")

                try:
                    from settlement.settlement_tracker import quick_record, add_record
                    add_record(quick_record(
                        project_key="tk",
                        company="Bilibili",
                        req_name=candidate["n_attach"]["name"] if candidate["n_attach"] else candidate["n_text"],
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
            time.sleep(0.3)

        except Exception as exc:
            failed += 1
            print(f"  Failed: {exc}")

    print(f"\nDone: {processed} processed, {skipped} skipped, {failed} failed.")


# ---- Sheet Read via lark-cli ----

def read_sheet_with_attachments(download_dir: Path) -> list[dict]:
    """Read all data rows including attachment info using lark-cli cells-get.

    Returns list of dicts: {row_num, values: {col_letter: ...}, attachments: {col_letter: {...}}}
    """
    # Read the full sheet in chunks (max ~200 rows per call to avoid timeout)
    all_rows: list[dict] = []

    # First, get the sheet dimensions
    result = _cli(
        "sheets", "+workbook-info",
        "--url", ORIGINAL_SHEET_URL,
    )
    info = json.loads(result.stdout)
    if not info.get("ok"):
        raise RuntimeError(f"Failed to get sheet info: {info}")
    sheets = info["data"]["sheets"]
    target = next(s for s in sheets if s["sheet_id"] == SHEET_ID)
    row_count = target["row_count"]
    col_count = target["column_count"]

    # Read in chunks of 100 rows
    chunk_size = 100
    for start_row in range(1, row_count + 1, chunk_size):
        end_row = min(start_row + chunk_size - 1, row_count)
        range_str = f"A{start_row}:O{end_row}"
        print(f"  Reading {range_str}...")

        result = _cli(
            "sheets", "+cells-get",
            "--url", ORIGINAL_SHEET_URL,
            "--sheet-id", SHEET_ID,
            "--range", range_str,
            timeout=120,
        )
        data = json.loads(result.stdout)
        if not data.get("ok"):
            raise RuntimeError(f"Failed to read cells: {data}")

        for rng in data["data"]["ranges"]:
            row_indices = rng["row_indices"]
            col_indices = rng["col_indices"]
            cells = rng["cells"]

            for ri, row_cells in enumerate(cells):
                row_num = row_indices[ri]
                row_data: dict[str, Any] = {"row_num": row_num, "values": {}, "attachments": {}}

                for ci, cell in enumerate(row_cells):
                    col = col_indices[ci]
                    if not cell:
                        row_data["values"][col] = ""
                        continue
                    row_data["values"][col] = cell.get("value", "")

                    # Extract attachment info from rich_text
                    rt = cell.get("rich_text")
                    if rt:
                        for elem in rt:
                            if elem.get("type") == "attachment":
                                row_data["attachments"][col] = {
                                    "token": elem.get("attachment_token", ""),
                                    "name": elem.get("text", ""),
                                    "mime_type": elem.get("mime_type", ""),
                                }
                                break

                all_rows.append(row_data)

    return all_rows


def find_candidates(all_rows: list[dict], tail: int = 0, blank_stop: int = 30) -> list[dict]:
    """Find rows where N has content and O is empty."""
    candidates = []
    start = max(0, len(all_rows) - tail) if tail > 0 else 0
    blank_streak = 0

    for i in range(start, len(all_rows)):
        row = all_rows[i]
        row_num = row["row_num"]
        n_val = str(row["values"].get("N", "") or "").strip()
        o_val = row["values"].get("O", "")
        n_attach = row["attachments"].get("N")

        # O 已填 → 跳过
        if o_val_is_filled(o_val):
            continue

        # N 空 → 累计空白，超过阈值停止扫描
        if not n_attach and not n_val:
            blank_streak += 1
            if blank_streak >= blank_stop:
                break
            continue

        blank_streak = 0
        candidates.append({
            "row": row_num,
            "n_attach": n_attach,
            "n_text": n_val,
        })

    return candidates


def o_val_is_filled(o_val: Any) -> bool:
    if o_val is None:
        return False
    if isinstance(o_val, str) and not o_val.strip():
        return False
    if isinstance(o_val, (int, float)):
        return True
    return bool(o_val)


# ---- Attachment Download via lark-cli ----

def download_html(file_token: str, file_name: str, download_dir: Path, row_num: int) -> Path | None:
    """Download attachment HTML via lark-cli api."""
    safe_name = "".join(ch if ch not in '<>:"/\\|?*' else "_" for ch in (file_name or f"row_{row_num}.html"))
    save_path = download_dir / f"{row_num}_{safe_name}"

    # Avoid overwrite
    if save_path.exists():
        stem = save_path.stem
        for idx in range(2, 1000):
            alt = save_path.with_name(f"{stem}_{idx}{save_path.suffix}")
            if not alt.exists():
                save_path = alt
                break

    try:
        # lark-cli api auto-saves binary responses; returns {saved_path, size_bytes, content_type}
        result = _cli(
            "api", "GET",
            f"/open-apis/drive/v1/medias/{file_token}/download",
            timeout=30,
        )
        data = json.loads(result.stdout)
        saved = Path(data.get("saved_path", ""))
        if not saved.exists() and "saved_path" not in data:
            print(f"  Download failed: {data}")
            return None

        saved = Path(data["saved_path"])
        if saved != save_path:
            saved.rename(save_path)
        print(f"  Downloaded: {save_path.name}")
        return save_path
    except Exception as exc:
        print(f"  Download failed: {exc}")
        return None


# ---- Sheet Write via lark-cli ----

def write_o_y(row_num: int, values: list[int]) -> None:
    """Write O-Y columns using lark-cli cells-set."""
    col_letters = ["O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y"]
    cells = []
    for v in values:
        cells.append({"value": v if v is not None else 0})

    range_str = f"O{row_num}:Y{row_num}"
    cells_json = json.dumps([cells])

    result = _cli(
        "sheets", "+cells-set",
        "--url", ORIGINAL_SHEET_URL,
        "--sheet-id", SHEET_ID,
        "--range", range_str,
        "--cells", cells_json,
        timeout=30,
    )
    data = json.loads(result.stdout)
    if not data.get("ok"):
        raise RuntimeError(f"Write failed: {data}")


def configure_output() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


if __name__ == "__main__":
    raise SystemExit(main())
