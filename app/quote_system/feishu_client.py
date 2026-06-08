from __future__ import annotations

import json
import os
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any

# 飞书应用凭证
APP_ID = os.getenv("FEISHU_APP_ID", "cli_aa882f3b1abb5bb7")
APP_SECRET = os.getenv("FEISHU_APP_SECRET", "qnfvPU0Dx1GHhUQbxGodfd7ItIQ6JHGn")

# 默认保留旧幻塔表，具体项目可单独指定 spreadsheet_token 和 sheet_id。
SPREADSHEET_TOKEN = os.getenv("FEISHU_SPREADSHEET_TOKEN", "Z6vowfyZZi8DySk1ixncexdmnle")
DEFAULT_SHEET_ID = os.getenv("FEISHU_SHEET_ID", "ZXBogz")
TENANT_DOMAIN = os.getenv("FEISHU_TENANT_DOMAIN", "my.feishu.cn")

# 列索引（0-based）
COL_SEQ = 0       # A列: 序号
COL_PROJECT = 1   # B列: 项目名称
COL_CONTENT = 2   # C列: 需求内容(语种)
COL_REQ_NAME = 3  # D列: 文本(需求名)
COL_TASK_NO = 4   # E列: Task No(需求单号)
COL_REQ_DATE = 5  # F列: 委托日期
COL_DELIV_DATE = 6  # G列: 交付日期
COL_DELIV_TIME = 7  # H列: 交付时间
COL_AMOUNT = 8    # I列: 报价金额
COL_STATUS = 9    # J列: 交付状态
COL_QUOTE = 10    # K列: 报价单附件
COL_MQXLZ = 11    # L列: mqxlz交付文件
COL_REMARK = 12   # M列: 备注
COL_BILLABLE = 13  # N列: 合计字数
COL_SOURCE_CHARS = 14  # O列: ソースの文字数


class FeishuClient:
    def __init__(
        self,
        sheet_id: str = DEFAULT_SHEET_ID,
        spreadsheet_token: str | None = SPREADSHEET_TOKEN,
    ):
        self._token: str | None = None
        self._token_expire: float = 0
        self.sheet_id = sheet_id
        self.spreadsheet_token = spreadsheet_token or SPREADSHEET_TOKEN

    def _ensure_token(self):
        if self._token and time.time() < self._token_expire - 60:
            return
        url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
        data = json.dumps({"app_id": APP_ID, "app_secret": APP_SECRET}).encode()
        req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
        resp = urllib.request.urlopen(req)
        result = json.loads(resp.read())
        if result.get("code") != 0:
            raise RuntimeError(f"飞书认证失败: {result}")
        self._token = result["tenant_access_token"]
        self._token_expire = time.time() + result.get("expire", 7200)

    def _api(self, method: str, url: str, body: dict | None = None) -> dict:
        self._ensure_token()
        data = json.dumps(body).encode() if body else None
        req = urllib.request.Request(url, data=data, headers={
            "Authorization": f"Bearer {self._token}",
            "Content-Type": "application/json; charset=utf-8",
        })
        req.method = method
        try:
            resp = urllib.request.urlopen(req)
            response_body = resp.read().decode()
            # 检查响应是否为HTML（错误页面）
            if response_body.startswith('<!doctype') or response_body.startswith('<!DOCTYPE'):
                print(f"飞书API返回HTML响应: {response_body[:500]}...")
                raise RuntimeError(f"飞书API返回错误页面，可能是表格权限或Token错误。响应内容: {response_body[:200]}...")
            return json.loads(response_body)
        except urllib.error.HTTPError as e:
            error_body = e.read().decode()
            if error_body.startswith('<!doctype') or error_body.startswith('<!DOCTYPE'):
                raise RuntimeError(f"飞书API返回错误页面 (HTTP {e.code})，可能是表格权限或Token错误")
            try:
                error_json = json.loads(error_body)
            except json.JSONDecodeError:
                error_json = {}
            if error_json.get("code") == 91403:
                raise RuntimeError(
                    f"飞书API错误 {e.code}: 没有权限访问该表格。"
                    "请把飞书应用添加为表格协作者，或确认应用已获得 Sheets/Drive 读写权限。"
                )
            raise RuntimeError(f"飞书API错误 {e.code}: {error_body}")
        except json.JSONDecodeError as e:
            response_body = resp.read().decode() if 'resp' in locals() else "No response body"
            raise RuntimeError(f"飞书API返回非JSON响应: {response_body[:200]}...")

    def read_sheet(self, sheet_id: str | None = None, range_str: str | None = None) -> list[list[Any]]:
        if sheet_id is None:
            sheet_id = self.sheet_id
        """读取飞书表格，返回二维数组"""

        # 检查 sheet_id 是否包含 !（如 "sheet_id!range" 格式）
        if '!' in sheet_id:
            parts = sheet_id.split('!', 1)
            actual_sheet_id = parts[0]
            actual_range = parts[1]
            url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{self.spreadsheet_token}/values/{actual_sheet_id}!{actual_range}"
        elif range_str:
            url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{self.spreadsheet_token}/values/{sheet_id}!{range_str}"
        else:
            url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{self.spreadsheet_token}/values/{sheet_id}"
        result = self._api("GET", url)
        if result.get("code") != 0:
            raise RuntimeError(f"读取表格失败: {result}")
        return result["data"]["valueRange"]["values"]

    def read_range(self, range_str: str, sheet_id: str | None = None) -> list[list[Any]]:
        """读取指定范围，如 A1:O20。"""
        return self.read_sheet(sheet_id=sheet_id, range_str=range_str)

    def read_cell(self, row: int, col: int | str, sheet_id: str | None = None) -> Any:
        """读取单个单元格的值。col 可传 0-based 数字或 A/B/C 列名。"""
        col_letter = _col_letter(col) if isinstance(col, int) else col.upper()
        range_str = f"{col_letter}{row}:{col_letter}{row}"
        result = self.read_range(range_str, sheet_id)
        return result[0][0] if result and result[0] else None

    def write_cell(self, row: int, col: int | str, value: Any, sheet_id: str | None = None):
        """写入一个单元格。col 可传 0-based 数字或 A/B/C 列名。"""
        target_sheet_id = sheet_id or self.sheet_id
        col_name = _col_letter(col) if isinstance(col, int) else col.upper()
        self.write_range(f"{target_sheet_id}!{col_name}{row}:{col_name}{row}", [[value]])

    def write_range(self, range_ref: str, values: list[list[Any]]):
        """写入二维数组到指定范围，range_ref 形如 d73cd4!A1:B2。"""
        body = {
            "valueRange": {
                "range": range_ref,
                "values": values,
            }
        }
        result = self._api(
            "PUT",
            f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{self.spreadsheet_token}/values",
            body,
        )
        if result.get("code") != 0:
            raise RuntimeError(f"写入范围失败: {result}")

    @classmethod
    def from_url(cls, sheet_url: str) -> "FeishuClient":
        """从飞书表格链接创建客户端。"""
        parsed = urllib.parse.urlparse(sheet_url)
        parts = [part for part in parsed.path.split("/") if part]
        try:
            spreadsheet_token = parts[parts.index("sheets") + 1]
        except (ValueError, IndexError) as exc:
            raise ValueError(f"无法从链接解析飞书表格 token: {sheet_url}") from exc
        query = urllib.parse.parse_qs(parsed.query)
        sheet_id = query.get("sheet", [DEFAULT_SHEET_ID])[0]
        return cls(sheet_id=sheet_id, spreadsheet_token=spreadsheet_token)

    def download_attachment(self, file_token: str, save_path: Path) -> Path:
        """下载飞书附件到本地"""
        self._ensure_token()
        url = f"https://open.feishu.cn/open-apis/drive/v1/medias/{file_token}/download"
        req = urllib.request.Request(url, headers={
            "Authorization": f"Bearer {self._token}",
        })
        resp = urllib.request.urlopen(req)
        save_path.parent.mkdir(parents=True, exist_ok=True)
        save_path.write_bytes(resp.read())
        return save_path

    def upload_to_drive(self, file_path: Path) -> str:
        """上传文件到飞书云盘，返回 file_token"""
        self._ensure_token()
        file_size = file_path.stat().st_size
        file_name = file_path.name

        url = "https://open.feishu.cn/open-apis/drive/v1/files/upload_all"
        boundary = "----FormBoundary" + os.urandom(8).hex()
        body = self._build_upload_body(boundary, file_path, file_name, file_size)

        req = urllib.request.Request(url, data=body, headers={
            "Authorization": f"Bearer {self._token}",
            "Content-Type": f"multipart/form-data; boundary={boundary}",
        })
        resp = urllib.request.urlopen(req)
        result = json.loads(resp.read())
        if result.get("code") != 0:
            raise RuntimeError(f"上传文件失败: {result}")
        return result["data"]["file_token"]

    def _build_upload_body(self, boundary: str, file_path: Path, file_name: str,
                           file_size: int) -> bytes:
        parts: list[bytes] = []
        parts.append(f"--{boundary}\r\n".encode())
        parts.append(b'Content-Disposition: form-data; name="file_name"\r\n\r\n')
        parts.append(f"{file_name}\r\n".encode())
        parts.append(f"--{boundary}\r\n".encode())
        parts.append(b'Content-Disposition: form-data; name="parent_type"\r\n\r\n')
        parts.append(b"explorer\r\n")
        parts.append(f"--{boundary}\r\n".encode())
        parts.append(b'Content-Disposition: form-data; name="size"\r\n\r\n')
        parts.append(f"{file_size}\r\n".encode())
        parts.append(f"--{boundary}\r\n".encode())
        parts.append(f'Content-Disposition: form-data; name="file"; filename="{file_name}"\r\n'.encode())
        parts.append(b"Content-Type: application/octet-stream\r\n\r\n")
        parts.append(file_path.read_bytes())
        parts.append(f"\r\n--{boundary}--\r\n".encode())
        return b"".join(parts)

    def set_file_public(self, file_token: str):
        """设置文件为互联网公开可访问（公司内外均可通过链接查看）"""
        url = f"https://open.feishu.cn/open-apis/drive/v1/permissions/{file_token}/public?type=file"
        body = {
            "link_share_entity": "anyone_readable",
            "security_entity": "anyone_can_view",
            "comment_entity": "anyone_can_view",
            "share_entity": "anyone",
            "invite_external": True,
        }
        result = self._api("PATCH", url, body)
        if result.get("code") != 0:
            raise RuntimeError(f"设置公开权限失败: {result}")

    def write_quote_link(self, row: int, file_token: str, file_name: str):
        """在K列写入报价单超链接公式"""
        file_url = f"https://{TENANT_DOMAIN}/file/{file_token}"
        formula = f'=HYPERLINK("{file_url}", "{file_name}")'
        self._write_cell_formula(row, COL_QUOTE, formula)

    def write_amount(self, row: int, amount: float):
        """在I列写入报价金额并居中"""
        self._write_cell_value(row, COL_AMOUNT, amount)
        self._set_cell_center(row, COL_AMOUNT)

    def write_stats(self, row: int, billable_words: float, source_chars: int):
        """在N列写入合计字数，O列写入ソースの文字数"""
        self._write_cell_value(row, COL_BILLABLE, round(billable_words, 1))
        self._write_cell_value(row, COL_SOURCE_CHARS, source_chars)

    def _set_cell_center(self, row: int, col: int):
        """设置单元格水平居中"""
        body = {
            "appendStyle": {
                "range": f"{self.sheet_id}!{_col_letter(col)}{row}:{_col_letter(col)}{row}",
                "style": {
                    "horizontalAlignment": 1,
                    "verticalAlignment": 1,
                },
            }
        }
        result = self._api("PUT",
            f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{self.spreadsheet_token}/style",
            body)
        if result.get("code") != 0:
            raise RuntimeError(f"设置单元格样式失败: {result}")

    def set_row_background(self, row: int, color: str):
        """设置整行背景色，color 为 hex 如 '#D9D9D9'"""
        body = {
            "appendStyle": {
                "range": f"{self.sheet_id}!A{row}:Z{row}",
                "style": {
                    "backColor": color,
                },
            }
        }
        result = self._api("PUT",
            f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{self.spreadsheet_token}/style",
            body)
        if result.get("code") != 0:
            raise RuntimeError(f"设置行背景色失败: {result}")

    def _write_cell_value(self, row: int, col: int, value):
        body = {
            "valueRange": {
                "range": f"{self.sheet_id}!{_col_letter(col)}{row}:{_col_letter(col)}{row}",
                "values": [[value]],
            }
        }
        result = self._api("PUT",
            f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{self.spreadsheet_token}/values",
            body)
        if result.get("code") != 0:
            raise RuntimeError(f"写入单元格失败: {result}")

    def _write_cell_formula(self, row: int, col: int, formula: str):
        body = {
            "valueRange": {
                "range": f"{self.sheet_id}!{_col_letter(col)}{row}:{_col_letter(col)}{row}",
                "values": [[{
                    "type": "formula",
                    "text": formula,
                }]],
            }
        }
        result = self._api("PUT",
            f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{self.spreadsheet_token}/values",
            body)
        if result.get("code") != 0:
            raise RuntimeError(f"写入公式失败: {result}")


def _col_letter(index: int) -> str:
    if index < 0:
        raise ValueError(f"列索引不能为负数: {index}")
    letters = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def find_unquoted_rows(rows: list[list[Any]]) -> list[dict]:
    """找到I列(报价金额)为空但K列有HTML附件的待报价行"""
    result = []
    for i, row in enumerate(rows):
        if i == 0:
            continue
        row_num = i + 1

        amount = row[COL_AMOUNT] if len(row) > COL_AMOUNT else None
        if amount is not None and str(amount).strip():
            continue

        quote_col = row[COL_QUOTE] if len(row) > COL_QUOTE else None

        # K列为附件类型
        if quote_col and isinstance(quote_col, list) and len(quote_col) > 0:
            attachment = quote_col[0]
            if isinstance(attachment, dict):
                file_token = attachment.get("fileToken")
                mime_type = attachment.get("mimeType", "")
                file_name = attachment.get("text", "")
                if file_token and ("html" in mime_type.lower() or file_name.lower().endswith(".html")):
                    result.append(_build_item(row_num, row, file_token, file_name))
                    continue

        # K列为公式类型（HYPERLINK，之前已处理过但可能需要重新处理）
        if quote_col and isinstance(quote_col, list) and len(quote_col) > 0:
            first = quote_col[0]
            if isinstance(first, dict) and first.get("type") == "formula":
                continue

        # K列为纯文本URL
        if quote_col and isinstance(quote_col, str) and quote_col.strip():
            continue

    return result


def _build_item(row_num: int, row: list, file_token: str, file_name: str) -> dict:
    req_date_serial = row[COL_REQ_DATE] if len(row) > COL_REQ_DATE else None
    req_date_str = excel_date_serial_to_date(req_date_serial)
    return {
        "row_num": row_num,
        "seq": row[COL_SEQ] if len(row) > COL_SEQ else None,
        "project_name": row[COL_PROJECT] if len(row) > COL_PROJECT else None,
        "content_type": row[COL_CONTENT] if len(row) > COL_CONTENT else None,
        "req_name": row[COL_REQ_NAME] if len(row) > COL_REQ_NAME else None,
        "task_no": row[COL_TASK_NO] if len(row) > COL_TASK_NO else None,
        "req_date": req_date_serial,
        "req_date_str": req_date_str,
        "req_date_raw": req_date_str.replace("-", "") if req_date_str else "",
        "deliv_date": row[COL_DELIV_DATE] if len(row) > COL_DELIV_DATE else None,
        "deliv_time": row[COL_DELIV_TIME] if len(row) > COL_DELIV_TIME else None,
        "html_token": file_token,
        "html_name": file_name,
    }


def excel_date_serial_to_date(serial) -> str | None:
    """将Excel日期序列号转换为日期字符串 YYYY-MM-DD"""
    from datetime import datetime, timedelta
    try:
        num = float(serial)
    except (TypeError, ValueError):
        return None
    base = datetime(1899, 12, 30)
    try:
        dt = base + timedelta(days=num)
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return None


def excel_time_serial_to_time(serial) -> str | None:
    """将Excel时间序列号转换为时间字符串 HH:MM"""
    try:
        num = float(serial)
    except (TypeError, ValueError):
        return None
    hours = int(num * 24)
    minutes = int((num * 24 - hours) * 60)
    return f"{hours:02d}:{minutes:02d}"
