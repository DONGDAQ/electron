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
from .utils import configure_output

ORIGINAL_SHEET_URL = "https://my.feishu.cn/sheets/R8QWsav9Nh77V3tWdTQchuCfnMc"
SHEET_ID = "d73cd4"
DOWNLOAD_DIR = Path(__file__).resolve().parent.parent / "outputs" / "tk_html"

import os
import shutil

# 优先使用全局安装的 lark-cli.cmd，其次用 node.exe 调用 run.js
_LARK_CLI_CMD = shutil.which("lark-cli") or shutil.which("lark-cli.cmd")

_WB = Path(os.path.expanduser("~/.workbuddy"))
_LARK_CLI_DIR = _WB / "binaries" / "node" / "cli-connector-packages"
_RUN_JS = _LARK_CLI_DIR / "node_modules" / "@larksuite" / "cli" / "scripts" / "run.js"

# Auto-detect node.exe from managed binaries
_NODE_EXE = None
_node_versions = sorted(
    (_WB / "binaries" / "node" / "versions").glob("*/node.exe"), reverse=True
)
if _node_versions:
    _NODE_EXE = str(_node_versions[0])
elif os.environ.get("LARK_CLI_NODE"):
    _NODE_EXE = os.environ["LARK_CLI_NODE"]


def _cli(*args: str, timeout: int = 60) -> subprocess.CompletedProcess:
    """Run lark-cli via global command or node directly."""
    if _LARK_CLI_CMD:
        cmd = [_LARK_CLI_CMD, "--format", "json"] + list(args)
    elif _RUN_JS.exists() and _NODE_EXE:
        cmd = [_NODE_EXE, str(_RUN_JS), "--format", "json"] + list(args)
    else:
        raise FileNotFoundError("lark-cli not found — install via 'npm install -g @larksuite/cli'")
    result = subprocess.run(
        cmd, capture_output=True, text=True, encoding="utf-8", timeout=timeout,
        env={**os.environ, "LARK_CLI_NO_PROXY": "1"},
    )
    # Safety: lark-cli may return empty output on errors; surface stderr
    if not result.stdout:
        stderr_snippet = (result.stderr or "")[:500]
        raise RuntimeError(
            f"lark-cli returned empty stdout (rc={result.returncode}). "
            f"stderr: {stderr_snippet or '(empty)'}"
        )
    return result


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
        print("\n已停止。")
        return 130
    except Exception as exc:
        print(f"错误: {exc}")
        return 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="auto-fill-tk", description="TK auto O-Y fill (lark-cli direct)")
    parser.add_argument("--dry-run", action="store_true", help="Scan only, no write")
    parser.add_argument("--tail", type=int, default=0, help="Process last N rows only")
    parser.add_argument("--blank-stop", type=int, default=200, help="Stop after N consecutive blanks")
    parser.add_argument("--max-rows", type=int, default=2000)
    parser.add_argument("--download-dir", default=str(DOWNLOAD_DIR))
    return parser


def run(args: argparse.Namespace) -> None:
    download_dir = Path(args.download_dir)
    if download_dir.exists():
        for f in download_dir.iterdir():
            if f.is_file():
                f.unlink()
        print("已清理旧的HTML文件。")
    download_dir.mkdir(parents=True, exist_ok=True)

    all_rows = read_sheet_with_attachments(download_dir)
    print(f"读取飞书表: {len(all_rows)} 行")

    candidates = find_candidates(all_rows, args.tail, args.blank_stop)
    if not candidates:
        print("没有待处理的数据。")
        return

    print(f"找到 {len(candidates)} 行待处理:")
    for c in candidates:
        label = c["n_attach"]["name"] if c["n_attach"] else c["n_text"]
        print(f"  第 {c['row']} 行: {label}")

    processed = 0
    skipped = 0
    failed = 0
    last_values: list[int] | None = None

    for idx, candidate in enumerate(candidates):
        row_num = candidate["row"]
        print(f"\n[{idx+1}/{len(candidates)}] 处理第 {row_num} 行...")

        try:
            # 纯文本 + "同上" → 复用上一行
            if not candidate["n_attach"] and is_same_as_above(candidate["n_text"]):
                if last_values is None:
                    print("  N列标记'同上'但没有上一行数据，跳过。")
                    skipped += 1
                    continue
                values = last_values
                print(f"  N列标记'同上'，复用上一行: {values}")

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
                print(f"  解析完成: O={values[0]}, P={values[1]}, Q={values[2]}...")

            # 纯文本非"同上" → 跳过
            else:
                # 忽略"提前翻译"等特殊标记
                if "提前翻译" in candidate['n_text'] or "无需统计" in candidate['n_text']:
                    continue
                print(f"  N列非附件非同上: {candidate['n_text'][:50]}，跳过。")
                skipped += 1
                continue

            if args.dry_run:
                print(f"  [预览模式] 跳过写入: {values}")
            else:
                write_o_y(row_num, values)
                print(f"  已写入 O-Y: {values}")

            processed += 1
            time.sleep(0.3)

        except Exception as exc:
            failed += 1
            print(f"  失败: {exc}")

    print(f"\n完成: {processed} 条已填写, {skipped} 条跳过, {failed} 条失败.")


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

    # Read N:O columns only (sparse-safe: col_indices is always ['N','O'])
    chunk_size = 50
    for start_row in range(1, row_count + 1, chunk_size):
        end_row = min(start_row + chunk_size - 1, row_count)
        range_str = f"N{start_row}:O{end_row}"
        print(f"  读取范围: {range_str}...")

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

            # 建立列字母→索引的映射，处理 sparse 数据
            col_pos = {col: idx for idx, col in enumerate(col_indices)}

            for ri, row_cells in enumerate(cells):
                row_num = row_indices[ri]
                row_data: dict[str, Any] = {"row_num": row_num, "values": {}, "attachments": {}}

                # 遍历我们关心的列：N 和 O
                for col_letter in ["N", "O"]:
                    ci = col_pos.get(col_letter)
                    if ci is None or ci >= len(row_cells):
                        row_data["values"][col_letter] = ""
                        continue
                    cell = row_cells[ci]
                    if not cell:
                        row_data["values"][col_letter] = ""
                        continue
                    row_data["values"][col_letter] = cell.get("value", "")

                    # Extract attachment info from rich_text
                    rt = cell.get("rich_text")
                    if rt:
                        for elem in rt:
                            if elem.get("type") == "attachment":
                                row_data["attachments"][col_letter] = {
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
            print(f"  下载失败: {data}")
            return None

        saved = Path(data["saved_path"])
        if saved != save_path:
            saved.rename(save_path)
        print(f"  已下载: {save_path.name}")
        return save_path
    except Exception as exc:
        print(f"  下载失败: {exc}")
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


# ---- TK 数据缓存 ----

CACHE_DIR = Path(__file__).resolve().parent.parent / "outputs" / "tk_cache"
CACHE_FILE = CACHE_DIR / "tk_settlement_data.json"


def sync_tk_data() -> dict:
    """从飞书表读取 TK 数据并保存到本地缓存。返回统计数据。"""
    import json as _json
    from datetime import date as _date

    CACHE_DIR.mkdir(parents=True, exist_ok=True)

    today = _date.today()
    current_month = today.month

    # 获取表实际行数（避免硬编码只读前 248 行）
    try:
        info_res = _cli("sheets", "+workbook-info", "--url", ORIGINAL_SHEET_URL)
        info = _json.loads(info_res.stdout)
        sheets = info["data"]["sheets"]
        target = next(s for s in sheets if s["sheet_id"] == SHEET_ID)
        max_row = target.get("row_count", 248)
    except Exception:
        max_row = 248
    # 从第3行开始读到表末尾，分块 50 行
    last_start = max(3, max_row - 1)

    settled = []
    unsettled = []

    # 分批读取
    for start in range(3, last_start + 1, 50):
        end = min(start + 49, max_row)
        if end < start:
            break
        result = _cli(
            "sheets", "+cells-get",
            "--url", ORIGINAL_SHEET_URL,
            "--sheet-id", SHEET_ID,
            "--range", f"B{start}:Z{end}",
            timeout=120,
        )
        data = _json.loads(result.stdout)

        for rng in data["data"]["ranges"]:
            for ri, row_cells in enumerate(rng["cells"]):
                row_num = rng["row_indices"][ri]
                project = row_cells[0].get("value", "") if len(row_cells) > 0 and row_cells[0] else ""
                req_name = row_cells[1].get("value", "") if len(row_cells) > 1 and row_cells[1] else ""
                if not project or not req_name:
                    continue

                deliv = str(row_cells[6].get("value", "")) if len(row_cells) > 6 and row_cells[6] and row_cells[6].get("value") else ""
                amount = 0
                if len(row_cells) > 10 and row_cells[10] and row_cells[10].get("value"):
                    try:
                        amount = float(str(row_cells[10]["value"]).replace(",", ""))
                    except (ValueError, TypeError):
                        pass
                status = str(row_cells[24].get("value", "")) if len(row_cells) > 24 and row_cells[24] and row_cells[24].get("value") else ""

                entry = {
                    "row": row_num,
                    "project": str(project),
                    "req_name": str(req_name)[:50],
                    "deliv": deliv,
                    "amount": amount,
                    "status": status,
                }

                is_current_or_future = False
                if deliv and "月" in deliv:
                    try:
                        month = int(deliv.split("月")[0])
                        is_current_or_future = month >= current_month
                    except (ValueError, IndexError):
                        pass

                if "已请款" in status or "已结算" in status:
                    settled.append(entry)
                elif is_current_or_future:
                    unsettled.append(entry)
                else:
                    unsettled.append(entry)

    result_data = {
        "sync_time": today.isoformat(),
        "settled_count": len(settled),
        "settled_amount": round(sum(e["amount"] for e in settled), 2),
        "unsettled_count": len(unsettled),
        "unsettled_amount": round(sum(e["amount"] for e in unsettled), 2),
        "settled": settled,
        "unsettled": unsettled,
    }

    CACHE_FILE.write_text(_json.dumps(result_data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"TK 数据已同步: 已请款 {len(settled)} 条, 未请款 {len(unsettled)} 条")
    return result_data


def load_tk_cache() -> dict | None:
    """从本地缓存读取 TK 数据。"""
    import json as _json
    if not CACHE_FILE.exists():
        return None
    try:
        return _json.loads(CACHE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return None


if __name__ == "__main__":
    raise SystemExit(main())
