"""COM 对象上下文管理器，消除重复的 CoInitialize/Dispatch/Cleanup 模式"""
from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path


@contextmanager
def com_excel():
    """Excel COM 上下文管理器。退出时自动 Close/Quit/CoUninitialize。"""
    import pythoncom
    pythoncom.CoInitialize()
    import win32com.client
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        yield excel
    finally:
        try:
            for wb in excel.Workbooks:
                try:
                    wb.Close(False)
                except Exception:
                    pass
        except Exception:
            pass
        try:
            excel.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


@contextmanager
def com_word():
    """Word COM 上下文管理器。退出时自动 Close/Quit/CoUninitialize。"""
    import pythoncom
    pythoncom.CoInitialize()
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        yield word
    finally:
        try:
            for doc in word.Documents:
                try:
                    doc.Close(False)
                except Exception:
                    pass
        except Exception:
            pass
        try:
            word.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def recalc_excel(file_path: Path) -> None:
    """用 COM 打开 xlsx 并保存（触发公式重算）"""
    with com_excel() as excel:
        wb = excel.Workbooks.Open(str(file_path.resolve()))
        wb.Save()
        wb.Close()


def xlsx_to_pdf(xlsx_path: Path, pdf_path: Path) -> None:
    """用 COM 将 xlsx 另存为 PDF"""
    with com_excel() as excel:
        wb = excel.Workbooks.Open(str(xlsx_path.resolve()))
        wb.SaveAs(str(pdf_path.resolve()), FileFormat=57)
        wb.Close()


def docx_to_pdf(docx_path: Path, pdf_path: Path) -> None:
    """用 COM 将 docx 另存为 PDF"""
    with com_word() as word:
        doc = word.Documents.Open(str(docx_path.resolve()))
        doc.SaveAs(str(pdf_path.resolve()), FileFormat=17)
        doc.Close()
