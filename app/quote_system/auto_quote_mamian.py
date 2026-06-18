"""马娘/HBR 飞书自动报价：从wiki表格读取需求 → 下载HTML → 生成报价单 → 回填飞书"""
from __future__ import annotations

import sys
import traceback
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from quote_system.feishu_client import (
    FeishuClient,
    resolve_wiki_token,
    excel_date_serial_to_date,
)
from quote_system.generator import QuoteRequest, generate_quote
from quote_system.memoq_html import parse_memoq_html, quote_words
from quote_system.projects import resolve_project, ProjectConfig
from quote_system.paths import get_quote_history_dir

WIKI_TOKEN = "WhA6waOPKiSajLkEPunc3Ob3nm8"

COL_LANG = 1       # B: 需求类型（多选）
COL_REQ_DATE = 2   # C: 需求日期
COL_DELIV_DATE = 3 # D: 交付日期
COL_HTML = 4       # E: HTML附件
COL_QUOTE_NAME = 5 # F: 报价单名（回填）
COL_WORDS = 6      # G: 报价字数（回填）


def run(project_key: str = "maniang"):
    project = resolve_project(project_key)

    work_dir = get_quote_history_dir() / project.history_dir
    print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {project.display_name}自动报价开始")

    actual_token = resolve_wiki_token(WIKI_TOKEN)
    client = FeishuClient(
        sheet_id=project.sheet_id,
        spreadsheet_token=actual_token,
    )

    try:
        rows = client.read_sheet()
    except Exception as e:
        print(f"读取飞书表格失败: {e}")
        return

    pending = find_pending_rows(rows)
    if not pending:
        print("没有找到待报价的需求")
        return

    print(f"找到 {len(pending)} 个待报价需求:")
    for it in pending:
        print(f"  行{it['row_num']}: {it['languages']}")

    for item in pending:
        try:
            process_one(client, project, work_dir, item)
        except Exception as e:
            print(f"处理行 {item['row_num']} 失败: {e}")
            traceback.print_exc()


def find_pending_rows(rows: list[list]) -> list[dict]:
    """找到G列为空且E列有HTML附件的行"""
    result = []
    for i, row in enumerate(rows):
        if i == 0:
            continue
        row_num = i + 1

        words_val = row[COL_WORDS] if len(row) > COL_WORDS else None
        if words_val is not None and str(words_val).strip():
            continue

        html_col = row[COL_HTML] if len(row) > COL_HTML else None
        if not html_col or not isinstance(html_col, list) or len(html_col) == 0:
            continue

        attach = html_col[0]
        if not isinstance(attach, dict):
            continue
        file_token = attach.get("fileToken")
        file_name = attach.get("text", "")
        if not file_token:
            continue

        lang_col = row[COL_LANG] if len(row) > COL_LANG else None
        languages = _parse_languages(lang_col)

        req_date = row[COL_REQ_DATE] if len(row) > COL_REQ_DATE else None
        deliv_date = row[COL_DELIV_DATE] if len(row) > COL_DELIV_DATE else None

        result.append({
            "row_num": row_num,
            "languages": languages,
            "req_date": req_date,
            "deliv_date": deliv_date,
            "file_token": file_token,
            "file_name": file_name,
        })
    return result


def _parse_languages(val) -> list[str]:
    """解析B列多选值"""
    if isinstance(val, list):
        result = []
        for v in val:
            if v:
                result.extend(s.strip() for s in str(v).split(",") if s.strip())
        return result
    if isinstance(val, str) and val.strip():
        return [s.strip() for s in val.split(",") if s.strip()]
    return ["日翻中"]


TEMP_DIR = ROOT / "outputs" / "mamian_html"


def process_one(client: FeishuClient, project: ProjectConfig, work_dir: Path, item: dict):
    row_num = item["row_num"]
    print(f"\n--- 处理行 {row_num} ---")

    TEMP_DIR.mkdir(parents=True, exist_ok=True)

    html_path = TEMP_DIR / item["file_name"]
    print(f"  下载HTML: {item['file_name']}")
    client.download_attachment(item["file_token"], html_path)

    stats = parse_memoq_html(html_path)
    all_row = stats.all_row
    print(f"  字数: 亚洲字符={all_row.source_asian_characters}, 总字符={all_row.source_chars}")

    req_date = _parse_date(item["req_date"])
    deliv_date = _parse_date(item["deliv_date"])

    languages = item["languages"]
    print(f"  生成报价单: {languages}")
    request = QuoteRequest(
        project=project,
        html_paths=[html_path],
        languages=languages,
        quote_date=req_date or date.today(),
        delivery_date=deliv_date,
        service_content=None,
        request_name=None,
        include_extract=False,
        output_path=work_dir,
    )
    result = generate_quote(ROOT, request)
    print(f"  报价单已生成: {result.final_path}")

    html_path.unlink(missing_ok=True)

    billable = _calc_billable_words(stats, languages)
    print(f"  计费字数: {billable}")

    quote_file = result.final_path
    if result.output_path != result.final_path and result.output_path.exists():
        result.output_path.unlink()
    print(f"  上传报价单到飞书云盘...")
    file_token = client.upload_to_drive(quote_file)
    client.set_file_public(file_token)

    print(f"  回填F列（报价单链接）...")
    _write_quote_link(client, row_num, file_token, quote_file.name)

    print(f"  回填G列（报价字数）...")
    client.write_cell(row_num, COL_WORDS, round(billable, 1))

    print(f"  行 {row_num} 处理完成!")


def _parse_date(val) -> date | None:
    if val is None:
        return None
    date_str = excel_date_serial_to_date(val)
    if date_str:
        try:
            return datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            pass
    return None


def _calc_billable_words(stats, languages: list[str]) -> float:
    qw = float(quote_words(stats))
    total_chars = float(stats.all_row.source_chars or 0)

    has_non_extract = any(lang != "摘字" for lang in languages)
    has_extract = any(lang == "摘字" for lang in languages)

    if has_non_extract and has_extract:
        return qw + total_chars
    if has_extract:
        return total_chars
    return qw


def _write_quote_link(client: FeishuClient, row: int, file_token: str, file_name: str):
    from quote_system.feishu_client import TENANT_DOMAIN, _col_letter
    file_url = f"https://{TENANT_DOMAIN}/file/{file_token}"
    formula = f'=HYPERLINK("{file_url}", "{file_name}")'
    body = {
        "valueRange": {
            "range": f"{client.sheet_id}!{_col_letter(COL_QUOTE_NAME)}{row}:{_col_letter(COL_QUOTE_NAME)}{row}",
            "values": [[{"type": "formula", "text": formula}]],
        }
    }
    result = client._api(
        "PUT",
        f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{client.spreadsheet_token}/values",
        body,
    )
    if result.get("code") != 0:
        raise RuntimeError(f"写入报价单链接失败: {result}")



if __name__ == "__main__":
    project_key = sys.argv[1] if len(sys.argv) > 1 else "maniang"
    run(project_key)
