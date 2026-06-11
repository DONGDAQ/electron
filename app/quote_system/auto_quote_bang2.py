"""BANG2自动报价：从飞书表格读取需求 → 下载HTML → 生成报价单"""
from __future__ import annotations

import sys
import traceback
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from quote_system.feishu_client import FeishuClient, excel_date_serial_to_date
from quote_system.projects import resolve_project
from quote_system.generator import QuoteRequest, generate_quote
from quote_system.memoq_html import parse_memoq_html, quote_words
from quote_system.save_path_config import get_save_path
from quote_system.paths import get_quote_history_dir

SPREADSHEET_TOKEN = "Wup0wnUPIiiIr2k8T23c4zASnjd"
SHEET_ID = "mzekJf"

LANG_MAP = {
    "日中": "日翻中",
    "日翻中": "日翻中",
    "日韩": "日翻韩",
    "日翻韩": "日翻韩",
    "日英": "日翻英",
    "日翻英": "日翻英",
}


def find_bang2_rows(rows: list[list]) -> list[dict]:
    """筛选F列为空的行，提取A-E列信息。"""
    result = []
    last_name = None
    for i, row in enumerate(rows):
        if i == 0:
            continue
        row_num = i + 1

        status = str(row[5]).strip() if len(row) > 5 else ""
        if status and status != "None":
            continue

        raw_name = str(row[0]).strip() if len(row) > 0 else ""
        if not raw_name:
            continue

        if raw_name == "同上":
            name = last_name or raw_name
        else:
            name = raw_name
            last_name = raw_name

        lang = str(row[1]).strip() if len(row) > 1 else ""

        date_serial = row[2] if len(row) > 2 else None
        date_str = excel_date_serial_to_date(date_serial) if date_serial else ""

        deliv_serial = row[3] if len(row) > 3 else None
        deliv_str = excel_date_serial_to_date(deliv_serial) if deliv_serial else ""

        e_val = row[4] if len(row) > 4 else None

        file_token = None
        file_name = None
        is_tongshang = False

        if isinstance(e_val, list) and len(e_val) > 0:
            attach = e_val[0]
            file_token = attach.get("fileToken")
            file_name = attach.get("text", "")
        elif isinstance(e_val, str) and e_val.strip() == "同上":
            is_tongshang = True

        result.append({
            "row_num": row_num,
            "req_name": name,
            "lang": lang,
            "date_str": date_str,
            "deliv_str": deliv_str,
            "file_token": file_token,
            "file_name": file_name,
            "is_tongshang": is_tongshang,
        })

    return result


def run() -> None:
    project = resolve_project("bang2")

    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] BANG2报价自动化开始")

    client = FeishuClient(spreadsheet_token=SPREADSHEET_TOKEN, sheet_id=SHEET_ID)

    try:
        rows = client.read_sheet()
    except Exception as e:
        print(f"读取飞书表格失败: {e}")
        return

    all_rows = find_bang2_rows(rows)
    if not all_rows:
        print("没有找到待报价的需求")
        return

    # 按A列(文档名)分组，连续同名行归为一组
    groups: list[list[dict]] = []
    current_group: list[dict] = []
    current_name = None
    for item in all_rows:
        if item["req_name"] != current_name:
            if current_group:
                groups.append(current_group)
            current_group = [item]
            current_name = item["req_name"]
        else:
            current_group.append(item)
    if current_group:
        groups.append(current_group)

    print(f"找到 {len(groups)} 个待报价文档")

    for group in groups:
        req_name = group[0]["req_name"]
        print(f"\n--- [{req_name}]: {len(group)} 行 ---")
        for item in group:
            tag = "同上" if item["is_tongshang"] else (item["file_name"] or "无附件")
            print(f"  行{item['row_num']}: {item['lang']} | {tag}")

    for group in groups:
        try:
            process_group(client, project, group)
        except Exception as e:
            req_name = group[0]["req_name"]
            print(f"处理 [{req_name}] 失败: {e}")
            traceback.print_exc()


def process_group(client: FeishuClient, project, group: list[dict]):
    req_name = group[0]["req_name"]
    print(f"\n=== 处理 [{req_name}] ===")

    html_dir = ROOT / "outputs" / "bang2_html"
    html_dir.mkdir(parents=True, exist_ok=True)

    # 遍历每行，下载有附件的，"同上"行复用上一个文件
    html_paths = []
    per_file_languages = []
    downloaded = []
    first_date_row = None

    for item in group:
        lang = LANG_MAP.get(item["lang"], item["lang"])
        if item["file_token"]:
            html_path = html_dir / item["file_name"]
            print(f"  下载HTML: {item['file_name']}")
            client.download_attachment(item["file_token"], html_path)
            html_paths.append(str(html_path))
            per_file_languages.append([lang])
            downloaded.append(html_path)
            if not first_date_row:
                first_date_row = item
        elif item["is_tongshang"] and html_paths:
            per_file_languages[-1].append(lang)
        else:
            print(f"  跳过行{item['row_num']}: 无附件且不同上")

    if not html_paths:
        print(f"  跳过: 没有找到HTML附件")
        return

    try:
        # 日期
        quote_date = date.today()
        if first_date_row["date_str"]:
            quote_date = datetime.strptime(first_date_row["date_str"], "%Y-%m-%d").date()

        delivery_date = None
        if first_date_row["deliv_str"]:
            delivery_date = datetime.strptime(first_date_row["deliv_str"], "%Y-%m-%d").date()

        # 生成报价单
        save_path = get_save_path("bang2")
        output_path = Path(save_path) if save_path else None

        request = QuoteRequest(
            project=project,
            html_paths=html_paths,
            languages=None,
            quote_date=quote_date,
            delivery_date=delivery_date,
            service_content=req_name,
            request_name=req_name,
            include_extract=False,
            output_path=output_path,
            per_file_languages=per_file_languages,
        )
        result = generate_quote(ROOT, request)
        print(f"  报价单已生成: {result.final_path}")

        # 结算字数：每个文件各自的字数
        stats_list = [parse_memoq_html(p) for p in downloaded]
        billables = [quote_words(s) for s in stats_list]

        # 回写飞书
        file_idx = 0
        billable_idx = 0
        for item in group:
            client.write_cell(item["row_num"], "F", "已报价")
            client.write_cell(item["row_num"], "H", result.final_path.name)
            if not item["is_tongshang"]:
                client.write_cell(item["row_num"], "I", billables[billable_idx])
                print(f"  行{item['row_num']} F→已报价, H→{result.final_path.name}, I→{billables[billable_idx]}")
                billable_idx += 1
            else:
                print(f"  行{item['row_num']} F→已报价, H→{result.final_path.name}, I→跳过(同上)")

    finally:
        for html_path in downloaded:
            if html_path.exists():
                html_path.unlink()
        print(f"  已删除 {len(downloaded)} 个临时HTML文件")


if __name__ == "__main__":
    run()
