"""MD（电心）月度结算单生成"""
import sys
import os
import shutil
from datetime import datetime, timedelta
from pathlib import Path
import re

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from quote_system.feishu_client import FeishuClient
from quote_system.paths import get_settlement_dir
from settlement._com_utils import com_excel, xlsx_to_pdf

SPREADSHEET_TOKEN = "FOF5wRPcfil8nvkLEnCc9qrXn4b"
SHEET_ID = "RCwC8C"
UNIT_PRICE = 0.424  # 含税单价(CNY)

if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys._MEIPASS)
    TEMPLATE_DIR = BASE_DIR / "模板" / "结算模板"  # PyInstaller 打包后模板在 _MEIPASS/模板/结算模板
else:
    BASE_DIR = Path(__file__).resolve().parent.parent.parent
    TEMPLATE_DIR = BASE_DIR / "app" / "模板" / "结算模板"  # 开发环境模板在 app/模板/结算模板


def read_feishu_data(year, month):
    """读取飞书表，筛选指定月份的已交付记录"""
    client = FeishuClient(spreadsheet_token=SPREADSHEET_TOKEN, sheet_id=SHEET_ID)
    data = client.read_sheet()

    month_start = datetime(year, month, 1)
    if month == 12:
        month_end = datetime(year + 1, 1, 1)
    else:
        month_end = datetime(year, month + 1, 1)
    base = datetime(1899, 12, 30)
    serial_start = (month_start - base).days
    serial_end = (month_end - base).days

    records = []
    for i, row in enumerate(data):
        if i == 0:
            continue
        if len(row) < 9:
            continue

        # 列8: 交付状态，只取"已交付"
        status = str(row[8] or "").strip()
        if status != "已交付":
            continue

        # 列5: 交付日期（Excel序列号）
        deliv_val = row[5] if len(row) > 5 else None
        if deliv_val is None:
            continue
        try:
            deliv_serial = float(deliv_val)
            if not (serial_start <= deliv_serial < serial_end):
                continue
        except (TypeError, ValueError):
            continue

        # 列7: 结算字数
        word_count = row[7] if len(row) > 7 else None
        # 列1: 需求名
        req_name = str(row[1] or "")
        # 列4: 需求日期
        req_date = row[4] if len(row) > 4 else None

        records.append({
            'row_index': i + 1,
            'req_name': req_name,
            'req_date': req_date,
            'deliv_date': deliv_serial,
            'word_count': float(word_count) if word_count is not None else 0,
        })

    return records


def _fill_template(records, year, month, template_filename, output_filename, output_dir):
    """通用的模板填充函数（使用 COM Excel，保留印章图片等格式）。
    
    模板结构：
      第3行 C3: 结算单(中译韩)
      第7行 C7=结算单号, E7=请款日期, F7=日期
      第9行 表头: C=翻译内容, D=字数, E=含税单价, F=金额, G=交付日
      第10-29行为数据行（最多20行）
      第30行: 合计
      第37行起: 联系信息
    """
    template = TEMPLATE_DIR / template_filename
    output_path = output_dir / output_filename

    if output_path.exists():
        try:
            os.remove(str(output_path))
        except PermissionError:
            ts = datetime.now().strftime('%H%M%S')
            stem = output_filename.rsplit('.', 1)[0]
            ext = output_filename.rsplit('.', 1)[1]
            output_filename = f"{stem}_{ts}.{ext}"
            output_path = output_dir / output_filename

    shutil.copy2(str(template), str(output_path))

    # ---- 按需求名排序 ----
    def _sort_key(rec):
        name = rec['req_name']
        m = re.search(r'(\d+(?:\.\d+)?)', name)
        if m:
            return (0, float(m.group(1)))
        return (1, name)
    records.sort(key=_sort_key)

    # ---- 用 COM Excel 填充数据（保留图片和格式）----
    with com_excel() as excel:
        wb = excel.Workbooks.Open(str(output_path.resolve()))
        ws = wb.ActiveSheet

        # 请款日期
        today = datetime.now()
        invoice_date_str = today.strftime('%Y%m%d')

        # C7: 结算单号（请款日期）
        ws.Range('C7').Value = f'结算单号：{invoice_date_str}'
        # E7: 请款日期
        ws.Range('E7').Value = f'请款日期：{today.strftime("%Y-%m-%d")}'
        # F7: 日期
        ws.Range('F7').Value = today

        # ---- 清除模板示例数据（第10-29行数据区，C-G列）----
        for row in range(10, 30):
            for col in range(3, 8):  # C=3, G=7
                ws.Cells(row, col).Value = None

        # ---- 填充数据（第10行开始）----
        data_start_row = 10
        for i, rec in enumerate(records):
            r = data_start_row + i
            if r > 29:
                break
            # C列: 翻译内容（需求名）
            ws.Cells(r, 3).Value = rec['req_name']
            # D列: 字数
            ws.Cells(r, 4).Value = rec['word_count'] if rec['word_count'] else 0
            # E列: 含税单价
            ws.Cells(r, 5).Value = UNIT_PRICE
            ws.Cells(r, 5).NumberFormat = '0.000'
            # F列: 金额公式 =D*E
            ws.Cells(r, 6).Formula = f'=D{r}*E{r}'
            # G列: 交付日（格式: YYYY/M/D）
            base = datetime(1899, 12, 30)
            deliv_dt = base + timedelta(days=int(rec['deliv_date']))
            ws.Cells(r, 7).Value = deliv_dt
            ws.Cells(r, 7).NumberFormat = 'yyyy/m/d'

        # ---- 隐藏无数据的行（第10+N 至 29行）----
        last_data_row = data_start_row + len(records) - 1
        for row in range(last_data_row + 1, 30):
            ws.Rows(row).Hidden = True

        wb.Save()
        wb.Close()

    print(f'[OK] 结算单已保存: {output_path}')
    return output_path


def generate_settlement_confirm(records, year, month, output_dir):
    """生成确认版结算单（Excel）"""
    template_file = f"结算单_MD-KR_{year}年{month}月份确认.xlsx"
    output_file = f"结算单_MD-KR_{year}年{month}月份确认.xlsx"
    return _fill_template(records, year, month, template_file, output_file, output_dir)


def generate_settlement_stamp(records, year, month, output_dir):
    """生成盖章版结算单（只保留PDF）"""
    template_file = f"结算单_MD-KR_{year}年{month}月份盖章版.xlsx"
    output_file = f"结算单_MD-KR_{year}年{month}月份盖章版.xlsx"

    excel_path = _fill_template(records, year, month, template_file, output_file, output_dir)

    # ---- Excel 转 PDF（复用项目已有工具）----
    pdf_path = excel_path.with_suffix('.pdf')
    xlsx_to_pdf(excel_path, pdf_path)
    # 删除 Excel，只保留 PDF
    try:
        os.remove(str(excel_path))
    except Exception:
        pass
    print(f'[OK] 盖章版PDF已生成: {pdf_path}')
    return pdf_path, pdf_path


def generate_invoice(records, year, month, output_dir):
    """生成請求書（Excel → PDF），使用 COM Excel 保留印章图片格式。
    
    請求書模板结构：
      G3: 请款日期（YYYYMMDD）
      C21: 总字数
      D21: =C21（请求字数）
      E21: 单价
      F21: =C21*E21（金额公式）
      H22: =F21（总额）
    """
    month_str = f"{year:04d}{month:02d}"
    template_file = f"請求書_{month_str}月【GamerHouse】.xlsx"
    output_file = f"請求書_{month_str}月【GamerHouse】.xlsx"

    template = TEMPLATE_DIR / template_file
    output_path = output_dir / output_file

    if output_path.exists():
        try:
            os.remove(str(output_path))
        except PermissionError:
            ts = datetime.now().strftime('%H%M%S')
            output_file = f"請求書_{month_str}月【GamerHouse】_{ts}.xlsx"
            output_path = output_dir / output_file

    shutil.copy2(str(template), str(output_path))

    total_words = sum(r['word_count'] for r in records)
    today = datetime.now()

    # ---- 用 COM Excel 填充（保留印章图片）----
    with com_excel() as excel:
        wb = excel.Workbooks.Open(str(output_path.resolve()))
        ws = wb.ActiveSheet

        # G3: 请款日期（YYYYMMDD）
        ws.Range('G3').Value = today.strftime('%Y%m%d')
        # A21: 结算标题（随月份变化）
        ws.Range('A21').Value = f'MD {month}月份结算'
        # H1: 日期
        ws.Range('H1').Value = today
        # C21: 总字数
        ws.Range('C21').Value = total_words
        # D21: 请求字数（=C21）
        ws.Range('D21').Formula = '=C21'
        # E21: 单价
        ws.Range('E21').Value = UNIT_PRICE
        ws.Range('E21').NumberFormat = '0.000'
        # F21: 金额公式
        ws.Range('F21').Formula = '=C21*E21'
        # H22: 总额
        ws.Range('H22').Formula = '=F21'

        wb.Save()
        wb.Close()

    print(f'[OK] 請求書Excel已保存: {output_path}')

    # 转 PDF
    pdf_path = output_path.with_suffix('.pdf')
    xlsx_to_pdf(output_path, pdf_path)
    # 删除 Excel，只保留 PDF
    try:
        os.remove(str(output_path))
    except Exception:
        pass
    print(f'[OK] 請求書PDF已生成: {pdf_path}')
    return pdf_path, pdf_path


def main():
    import argparse
    parser = argparse.ArgumentParser(description='MD月度结算单生成')
    parser.add_argument('--year', type=int, default=2026)
    parser.add_argument('--month', type=int, default=6)
    parser.add_argument('--output', type=str, default=None)
    args = parser.parse_args()

    year, month = args.year, args.month
    output_dir = Path(args.output) if args.output else get_settlement_dir() / f"{year}年{month}月"
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f'读取飞书数据: {year}年{month}月...')
    records = read_feishu_data(year, month)
    print(f'找到 {len(records)} 条交付记录')

    if not records:
        print('没有找到匹配的交付记录，退出。')
        return

    # 生成确认版Excel
    confirm_path = generate_settlement_confirm(records, year, month, output_dir)
    print(f'确认版Excel已生成: {confirm_path}')

    # 生成盖章版PDF
    _, stamp_pdf = generate_settlement_stamp(records, year, month, output_dir)
    print(f'盖章版PDF已生成: {stamp_pdf}')

    print('\n完成！')


if __name__ == '__main__':
    main()
