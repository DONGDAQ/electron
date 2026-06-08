"""完美世界报价自动化：从飞书表格读取需求 → 下载HTML → 生成报价单 → 上传回飞书"""
from __future__ import annotations

import sys
import traceback
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from quote_system.feishu_client import (
    FeishuClient,
    find_unquoted_rows,
    excel_date_serial_to_date,
    excel_time_serial_to_time,
)
from quote_system.projects import ProjectConfig
from quote_system.generator import QuoteRequest, generate_quote
from quote_system.memoq_html import parse_memoq_html
from quote_system.projects import resolve_project, ProjectConfig
from quote_system.save_path_config import get_save_path
from quote_system.paths import get_quote_history_dir


def run(project_key: str = "huanta"):
    project = resolve_project(project_key)
    if project.generator != "huanta":
        print(f"项目 {project.display_name} 不支持自动报价（需使用幻塔模板）")
        return

    work_dir = get_quote_history_dir() / project.history_dir

    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {project.display_name}报价自动化开始")

    # 检查是否是外部表格（需要API模式）
    if getattr(project, 'spreadsheet_token', None) and project.generator == "huanta":
        client = FeishuClient(
            sheet_id=project.sheet_id,
            spreadsheet_token=project.spreadsheet_token or None,
        )
    else:
        # 正常API模式
        client = FeishuClient(
            sheet_id=project.sheet_id,
            spreadsheet_token=project.spreadsheet_token or None,
        )

    try:
        rows = client.read_sheet()
    except Exception as e:
        print(f"读取飞书表格失败: {e}")
        return

    unquoted = find_unquoted_rows(rows)
    if not unquoted:
        print("没有找到待报价的需求")
        return

    _assign_date_labels(unquoted)

    print(f"找到 {len(unquoted)} 个待报价需求:")
    for item in unquoted:
        print(f"  行{item['row_num']}: 委托日期={item.get('date_label')}, "
              f"需求名={item['req_name']}, TaskNo={item['task_no']}, HTML={item['html_name']}")

    for item in unquoted:
        try:
            process_one(client, project, work_dir, item)
        except Exception as e:
            print(f"处理行 {item['row_num']} 失败: {e}")
            traceback.print_exc()


def _assign_date_labels(items: list[dict]):
    from collections import defaultdict
    groups = defaultdict(list)
    for item in items:
        date_str = item.get("req_date_raw", "")
        groups[date_str].append(item)

    for date_str, group in groups.items():
        if len(group) == 1:
            group[0]["date_label"] = date_str
        else:
            for i, item in enumerate(group, start=1):
                item["date_label"] = f"{date_str}-{i}"


def process_one(client: FeishuClient, project: ProjectConfig, work_dir: Path, item: dict):
    row_num = item["row_num"]
    seq = item["seq"]
    print(f"\n--- 处理行 {row_num} (序号{seq}) ---")

    task_folder = work_dir / str(seq)
    task_folder.mkdir(parents=True, exist_ok=True)

    html_name = item["html_name"]
    html_path = task_folder / html_name
    print(f"  下载HTML: {html_name}")
    client.download_attachment(item["html_token"], html_path)

    stats = parse_memoq_html(html_path)
    all_row = stats.all_row
    print(f"  字数: 亚洲字符={all_row.source_asian_characters}, 总字符={all_row.source_chars}")

    req_date_str = item.get("req_date_str")
    if req_date_str:
        quote_date = datetime.strptime(req_date_str, "%Y-%m-%d").date()
    else:
        quote_date = date.today()
    delivery_date = None
    if item["deliv_date"] is not None:
        date_str = excel_date_serial_to_date(item["deliv_date"])
        if date_str:
            delivery_date = datetime.strptime(date_str, "%Y-%m-%d").date()

    delivery_time = None
    if item["deliv_time"] is not None:
        delivery_time = excel_time_serial_to_time(item["deliv_time"])

    content_type = str(item["content_type"]).strip() if item["content_type"] else "翻译+润色"

    print(f"  生成报价单...")
    request = QuoteRequest(
        project=project,
        html_paths=[html_path],
        languages=[content_type],
        quote_date=quote_date,
        delivery_date=delivery_date,
        service_content=None,
        request_name=item["req_name"],
        include_extract=False,
        output_path=task_folder,
        task_no=item["task_no"],
        delivery_time=delivery_time,
        date_label=item.get("date_label"),
    )
    result = generate_quote(ROOT, request)
    print(f"  报价单已生成: {result.final_path}")

    try:
        from settlement.settlement_tracker import record_from_quote, add_record
        add_record(record_from_quote(request, result.stats, result.final_path, source="auto"))
    except Exception:
        pass

    _copy_to_save_path(project.key, result.final_path)

    billable_words = huan_ta_billable_words(stats)
    price_per_word = project.prices.get(content_type, 0)
    total_amount = round(billable_words * price_per_word, 2)
    print(f"  计费字数: {billable_words}, 单价: {price_per_word}, 总额: {total_amount}")

    # 上传子目录中的副本，删除历史根目录的副本（只保留一份）
    quote_file = result.final_path
    if result.output_path != result.final_path and result.output_path.exists():
        result.output_path.unlink()
    print(f"  上传报价单到飞书云盘...")
    file_token = client.upload_to_drive(quote_file)
    client.set_file_public(file_token)
    print(f"  写入报价单链接到K列...")
    client.write_quote_link(row=row_num, file_token=file_token, file_name=quote_file.name)

    print(f"  回写报价金额: {total_amount}")
    client.write_amount(row=row_num, amount=total_amount)

    print(f"  写入统计: 合计字数={round(billable_words, 1)}, ソースの文字数={all_row.source_chars}")
    client.write_stats(row=row_num, billable_words=billable_words, source_chars=all_row.source_chars)

    print(f"  行 {row_num} 处理完成!")


def huan_ta_billable_words(stats) -> float:
    from quote_system.generator import _build_huanta_match_map

    all_row = stats.all_row
    total_chars = all_row.source_chars or 0
    total_asian = all_row.source_asian_characters or 0

    punctuation_billable = (total_chars - total_asian) * 0.3

    match_weights = {
        "context": 0, "exact": 0, "101": 0, "100": 0.1,
        "95-99": 0.4, "85-94": 1, "75-84": 1, "50-74": 1, "no_match": 1,
    }

    match_map = _build_huanta_match_map(stats.rows)
    weighted = 0.0
    for match_type, weight in match_weights.items():
        matched = match_map.get(match_type)
        if matched and matched.source_asian_characters:
            weighted += matched.source_asian_characters * weight

    return punctuation_billable + weighted


def _copy_to_save_path(project_key: str, source: Path):
    save_path = get_save_path(project_key)
    if not save_path:
        return
    dest_dir = Path(save_path)
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / source.name
    import shutil
    shutil.copy2(str(source), str(dest))
    print(f"  已另存至: {dest}")


if __name__ == "__main__":
    run()
