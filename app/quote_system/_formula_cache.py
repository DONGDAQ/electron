"""报价单公式缓存工具

生成报价单后，openpyxl 只写公式字符串、不执行计算，
导致 G48/I48 等合计单元格的 Excel 缓存为 None。
本模块在用 Excel COM 打开文件、强制重算、保存，
将计算结果写入文件缓存区，后续 data_only 读取即可拿到数值。
"""

from __future__ import annotations

from pathlib import Path

from settlement._com_utils import com_excel


def ensure_quote_formulas_cached(out_path: Path) -> None:
    """生成报价单后用 Excel COM 重算公式并缓存到文件

    - out_path: 报价单文件路径（Path）
    - 失败时不抛异常，避免阻塞主流程
    - report_cache.py 读取时发现缓存为 None 会再次触发本函数
    """
    try:
        with com_excel() as excel:
            wb = excel.Workbooks.Open(str(out_path.resolve()))
            excel.CalculateUntilAsyncQueriesDone()
            wb.Save()
            wb.Close()
    except Exception:
        # 缓存失败不阻塞主流程，report_cache.py 读取时会自动重算
        pass
