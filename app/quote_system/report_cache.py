"""报告数据缓存：提前读取所有项目数据，保存到 JSON 供报告页直接读取"""
import json
import os
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl

OUTPUTS_DIR = Path(r"D:\baojia\electron\outputs")
CACHE_FILE = OUTPUTS_DIR / "cache" / "report_dashboard.json"


def _win32com_refresh(fpath: Path):
    import win32com.client
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb_com = excel.Workbooks.Open(str(fpath))
        wb_com.Save()
        wb_com.Close()
    except Exception:
        pass
    finally:
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass


def _read_quote_amount(fpath: Path) -> float:
    from quote_system.generator import is_formula_cached
    cached = is_formula_cached(fpath)
    try:
        wb = openpyxl.load_workbook(str(fpath), data_only=True)
        ws = wb.active
        if "战双" in fpath.name and "版更" in str(fpath.parent):
            val = ws.cell(18, 7).value
            wb.close()
            if val is not None and float(val) > 0:
                return round(float(val), 2)
            if cached:
                return 0.0
            _win32com_refresh(fpath)
            wb2 = openpyxl.load_workbook(str(fpath), data_only=True)
            val2 = wb2.cell(18, 7).value
            wb2.close()
            return round(float(val2), 2) if val2 else 0.0
        if "战双" in fpath.name and "发行" in str(fpath.parent):
            val = ws.cell(52, 8).value
            wb.close()
            if val is not None and float(val) > 0:
                return round(float(val), 2)
            if cached:
                return 0.0
            _win32com_refresh(fpath)
            wb2 = openpyxl.load_workbook(str(fpath), data_only=True)
            val2 = wb2.cell(52, 8).value
            wb2.close()
            return round(float(val2), 2) if val2 else 0.0
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            has_total = any(
                c.value and isinstance(c.value, str) and "合计" in c.value
                for c in row
            )
            if has_total:
                for cell in reversed(row):
                    if cell.value is not None and isinstance(cell.value, (int, float)):
                        wb.close()
                        return round(float(cell.value), 2)
        wb.close()
        if cached:
            return 0.0
        _win32com_refresh(fpath)
        wb2 = openpyxl.load_workbook(str(fpath), data_only=True)
        ws2 = wb2.active
        for row in ws2.iter_rows(min_row=1, max_row=ws2.max_row):
            has_total = any(
                c.value and isinstance(c.value, str) and "合计" in c.value
                for c in row
            )
            if has_total:
                for cell in reversed(row):
                    if cell.value is not None and isinstance(cell.value, (int, float)):
                        wb2.close()
                        return round(float(cell.value), 2)
        wb2.close()
    except Exception:
        pass
    return 0.0


def _calc_quote_amount(fpath: Path) -> float:
    try:
        wb = openpyxl.load_workbook(str(fpath))
        quote_ws = wb.active
        title = quote_ws.title
        if title == "需求列表":
            sheet1 = wb["Sheet1"] if "Sheet1" in wb.sheetnames else None
            if not sheet1:
                wb.close()
                return 0.0
            blocks = []
            cur_asian = cur_chars = cur_rep = 0
            cur_start = 0
            for row in sheet1.iter_rows(min_row=1, max_row=sheet1.max_row):
                a_val = row[0].value
                if not a_val:
                    continue
                a_str = str(a_val).strip()
                if a_str == "All":
                    if cur_start:
                        eff = cur_asian - cur_rep
                        punc = cur_chars - cur_asian
                        blocks.append((cur_start, round(eff + punc * 0.3)))
                    cur_start = row[0].row
                    cur_asian = row[4].value if isinstance(row[4].value, (int, float)) else 0
                    cur_chars = row[5].value if isinstance(row[5].value, (int, float)) else 0
                    cur_rep = 0
                elif a_str == "Repetition":
                    cur_rep = row[4].value if isinstance(row[4].value, (int, float)) else 0
            if cur_start:
                eff = cur_asian - cur_rep
                punc = cur_chars - cur_asian
                blocks.append((cur_start, round(eff + punc * 0.3)))
            import re as _re
            total = 0.0
            for row in quote_ws.iter_rows(min_row=2, max_row=quote_ws.max_row):
                f_val = row[5].value if len(row) > 5 else None
                g_val = row[6].value if len(row) > 6 else None
                if not (g_val and isinstance(g_val, (int, float)) and g_val > 0):
                    continue
                if not isinstance(f_val, str) or "=Sheet1!J" not in f_val:
                    continue
                m = _re.search(r"J(\d+)", f_val)
                if not m:
                    continue
                j_row = int(m.group(1))
                for b_start, b_words in blocks:
                    if j_row >= b_start and j_row <= b_start + 15:
                        total += b_words * g_val * 1.06
                        break
            wb.close()
            return round(total, 2)
        unit_price = 0.0
        for row in quote_ws.iter_rows(min_row=9, max_row=25):
            for c in row:
                if (c.column >= 6 and c.value is not None
                        and isinstance(c.value, (int, float))
                        and 0 < c.value < 2):
                    unit_price = float(c.value)
                    break
            if unit_price:
                break
        if not unit_price:
            wb.close()
            return 0.0
        stats_ws = None
        for name in ("具体数据", "【字数统计】", "字数统计"):
            if name in wb.sheetnames:
                stats_ws = wb[name]
                break
        if not stats_ws:
            wb.close()
            return 0.0
        all_asian = all_chars = rep_asian = 0
        for row in stats_ws.iter_rows(min_row=1, max_row=min(stats_ws.max_row, 40)):
            a_val = row[0].value
            if not a_val:
                continue
            a_str = str(a_val).strip()
            if a_str in ("全部", "All", "すべて"):
                e_val = row[4].value if len(row) > 4 else None
                f_val = row[5].value if len(row) > 5 else None
                if isinstance(e_val, (int, float)):
                    all_asian = e_val
                if isinstance(f_val, (int, float)):
                    all_chars = f_val
            elif a_str in ("重复", "Repetition", "繰り返し"):
                e_val = row[4].value if len(row) > 4 else None
                if isinstance(e_val, (int, float)):
                    rep_asian = e_val
        wb.close()
        if not all_asian:
            return 0.0
        effective = all_asian - rep_asian
        punctuation = all_chars - all_asian
        words = round(effective + punctuation * 0.3)
        return round(words * unit_price * 1.06, 2)
    except Exception:
        return 0.0


def sync_report_cache() -> dict:
    """读取所有项目数据，汇总后保存到缓存 JSON。"""
    from quote_system.paths import get_quote_history_dir

    today = date.today()
    month_start = today.replace(day=1)
    history_dir = get_quote_history_dir()

    # 上月1号
    if month_start.month == 1:
        last_month_start = month_start.replace(year=month_start.year - 1, month=12)
    else:
        last_month_start = month_start.replace(month=month_start.month - 1)

    _skip_companies = {"完美世界", "4399"}
    file_meta = []
    if history_dir.exists():
        for f in history_dir.rglob("*.xlsx"):
            try:
                rel = f.relative_to(history_dir)
                if rel.parts and rel.parts[0] in _skip_companies:
                    continue
                if any("已结算" in part for part in rel.parts):
                    continue
                mtime_date = date.fromtimestamp(f.stat().st_mtime)
                amount = _read_quote_amount(f)
                file_meta.append((rel.parts, mtime_date, amount))
            except Exception:
                pass

    month_file_count = 0
    month_total_amount = 0.0
    unsettled_count = 0
    unsettled_amount = 0.0
    trend = defaultdict(float)
    company_summary = defaultdict(lambda: {"amount": 0.0, "count": 0, "projects": {}})

    for parts, mtime_date, amount in file_meta:
        if mtime_date >= month_start:
            month_file_count += 1
        company = parts[0] if parts else "未知"
        unsettled_count += 1
        unsettled_amount += amount
        # 本地文件用修改日期近似交付日期
        if mtime_date >= last_month_start:
            month_total_amount += amount
        if (today - mtime_date).days <= 30:
            trend[mtime_date.strftime("%Y-%m-%d")] += amount
        project = parts[1] if len(parts) > 1 else company
        company_summary[company]["amount"] += amount
        company_summary[company]["count"] += 1
        if project not in company_summary[company]["projects"]:
            company_summary[company]["projects"][project] = 0
        company_summary[company]["projects"][project] += 1

    trend_list = []
    for i in range(29, -1, -1):
        d = today - timedelta(days=i)
        key = d.strftime("%Y-%m-%d")
        trend_list.append({"date": key, "amount": round(trend.get(key, 0), 2)})

    # TK (USD → CNY, 汇率6.8)
    try:
        from quote_system.auto_fill_tk import load_tk_cache
        tk_data = load_tk_cache()
        if tk_data:
            tk_rate = 6.8
            tk_count = tk_data.get("unsettled_count", 0)
            tk_amount_cny = round(tk_data.get("unsettled_amount", 0) * tk_rate, 2)
            unsettled_count += tk_count
            unsettled_amount += tk_amount_cny
            # TK 交付日期在上月/本月的金额
            tk_month_amount = 0
            for e in tk_data.get("unsettled", []):
                deliv_str = e.get("deliv", "")
                try:
                    month = int(deliv_str.split("月")[0])
                    if month >= last_month_start.month and month <= today.month:
                        tk_month_amount += round(e.get("amount", 0) * tk_rate, 2)
                except (ValueError, IndexError):
                    pass
            month_total_amount += tk_month_amount
            company_summary["Bilibili"]["amount"] += tk_amount_cny
            company_summary["Bilibili"]["count"] += tk_count
            if "TK" not in company_summary["Bilibili"]["projects"]:
                company_summary["Bilibili"]["projects"]["TK"] = 0
            company_summary["Bilibili"]["projects"]["TK"] += tk_count
    except Exception:
        pass

    # 完美世界
    try:
        from quote_system.feishu_client import FeishuClient
        from settlement._perfect_world_engine import PERFECT_WORLD_PROFILES
        pw_projects = [("幻塔", "huanta"), ("异环游戏内", "yihuan_nei"), ("异环发行", "yihuan_faxing")]
        base = datetime(1899, 12, 30)
        for pw_name, pw_key in pw_projects:
            profile = PERFECT_WORLD_PROFILES[pw_key]
            client = FeishuClient(profile.sheet_id)
            data = client.read_sheet(profile.sheet_id, 'A1:O')
            proj_count = 0
            proj_amount = 0
            for i, row in enumerate(data):
                if i == 0:
                    continue
                g = row[6] if len(row) > 6 else None
                if g is None:
                    continue
                try:
                    g_num = float(g)
                except (TypeError, ValueError):
                    continue
                deliv = base + timedelta(days=g_num)
                if deliv.date() < last_month_start:
                    continue
                j = str(row[9] or '').strip() if len(row) > 9 else ''
                if j == '已请款':
                    continue
                proj_count += 1
                try:
                    amt = float(row[8]) if len(row) > 8 and row[8] is not None else 0
                except (ValueError, TypeError):
                    amt = 0
                proj_amount += amt
                # 交付日期在本月之前的计入本月总额
                if deliv.date() >= month_start or deliv.date() >= last_month_start:
                    month_total_amount += amt
            if proj_count:
                label = f"完美世界-{pw_name}"
                company_summary[label]["amount"] += proj_amount
                company_summary[label]["count"] += proj_count
                if pw_name not in company_summary[label]["projects"]:
                    company_summary[label]["projects"][pw_name] = 0
                company_summary[label]["projects"][pw_name] += proj_count
            unsettled_count += proj_count
            unsettled_amount += proj_amount
    except Exception:
        pass

    # 4399
    try:
        from quote_system.feishu_client import FeishuClient
        from settlement.generate_settlement_4399 import (
            SPREADSHEET_TOKEN_4399, SHEET_ID_4399,
            SPREADSHEET_TOKEN_BOQI, SHEET_ID_BOQI,
        )
        base4 = datetime(1899, 12, 30)
        count_4399 = 0
        amount_4399 = 0
        c4399 = FeishuClient(spreadsheet_token=SPREADSHEET_TOKEN_4399, sheet_id=SHEET_ID_4399)
        rows4_raw = c4399.read_sheet()
        rows4_fmt = c4399.read_sheet(render='FormattedValue')
        for i, row in enumerate(rows4_raw):
            if i == 0:
                continue
            dv = row[4] if len(row) > 4 else None
            if dv is None:
                continue
            try:
                ds = float(dv)
            except (TypeError, ValueError):
                continue
            dd = base4 + timedelta(days=ds)
            if dd.date() < last_month_start:
                continue
            o = str(row[14] or '').strip() if len(row) > 14 else ''
            if o == '已请款':
                continue
            count_4399 += 1
            try:
                fmt_row = rows4_fmt[i] if i < len(rows4_fmt) else row
                j_val = float(fmt_row[9]) if len(fmt_row) > 9 and fmt_row[9] is not None else 0
                amount_4399 += j_val
                if dd.date() >= last_month_start:
                    month_total_amount += j_val
            except (ValueError, TypeError):
                pass
        cbo = FeishuClient(spreadsheet_token=SPREADSHEET_TOKEN_BOQI, sheet_id=SHEET_ID_BOQI)
        rowsb_raw = cbo.read_sheet()
        rowsb_fmt = cbo.read_sheet(render='FormattedValue')
        for i, row in enumerate(rowsb_raw):
            if i == 0:
                continue
            dv = row[3] if len(row) > 3 else None
            if dv is None:
                continue
            try:
                ds = float(dv)
            except (TypeError, ValueError):
                continue
            dd = base4 + timedelta(days=ds)
            if dd.date() < last_month_start:
                continue
            g = str(row[6] or '').strip() if len(row) > 6 else ''
            if g != '已交付':
                continue
            count_4399 += 1
            try:
                fmt_row = rowsb_fmt[i] if i < len(rowsb_fmt) else row
                amt = float(fmt_row[5]) if len(fmt_row) > 5 and fmt_row[5] is not None else 0
                amount_4399 += amt
                if dd.date() >= last_month_start:
                    month_total_amount += amt
            except (ValueError, TypeError):
                pass
        if count_4399:
            company_summary["4399"]["amount"] += amount_4399
            company_summary["4399"]["count"] += count_4399
            if "4399" not in company_summary["4399"]["projects"]:
                company_summary["4399"]["projects"]["4399"] = 0
            company_summary["4399"]["projects"]["4399"] += count_4399
        unsettled_count += count_4399
        unsettled_amount += amount_4399
    except Exception:
        pass

    result = {
        "sync_time": datetime.now().isoformat(),
        "month_file_count": month_file_count,
        "month_total_amount": round(month_total_amount, 2),
        "unsettled_count": unsettled_count,
        "unsettled_amount": round(unsettled_amount, 2),
        "trend": trend_list,
        "company_summary": {k: v for k, v in company_summary.items()},
    }

    CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    CACHE_FILE.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"报告缓存已同步: 未结算 {unsettled_count} 条, ¥{unsettled_amount:.2f}")
    return result


if __name__ == "__main__":
    sync_report_cache()
