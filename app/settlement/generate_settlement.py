#!/usr/bin/env python3
"""幻塔月度结算单生成脚本（薄封装 → _perfect_world_engine）"""
import sys
import os
from pathlib import Path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from settlement._perfect_world_engine import (
    SettlementProfile,
    PERFECT_WORLD_PROFILES,
    run_settlement,
    _excel_serial_to_date,
    _format_date_ymd,
    _get_unit_price,
    _fill_acceptance_row,
    TEMPLATE_DIR,
    BASE_DIR,
)

_PROFILE = PERFECT_WORLD_PROFILES["huanta"]


# ---- 向后兼容封装（保持旧的 2/4 参数调用签名）----

def read_feishu_data(year, month):
    """从飞书读取指定月份的交付数据（幻塔）"""
    from settlement._perfect_world_engine import read_feishu_data as _read
    return _read(year, month, _PROFILE)


def generate_settlement_excel(records, year, month, output_dir):
    """生成结算单Excel文件（幻塔）"""
    from settlement._perfect_world_engine import generate_settlement_excel as _gen
    return _gen(records, year, month, output_dir, _PROFILE)


def generate_acceptance_docx(records, year, month, output_dir):
    """生成验收单Word文件（幻塔）"""
    from settlement._perfect_world_engine import generate_acceptance_docx as _gen
    return _gen(records, year, month, output_dir, _PROFILE)


def main():
    import argparse
    parser = argparse.ArgumentParser(description='幻塔月度结算单生成')
    parser.add_argument('--year', type=int, default=2026, help='年份')
    parser.add_argument('--month', type=int, default=4, help='月份')
    parser.add_argument('--output', type=str, default=None, help='输出目录')
    args = parser.parse_args()
    output_dir = Path(args.output) if args.output else None
    run_settlement("huanta", args.year, args.month, output_dir)


if __name__ == '__main__':
    main()
