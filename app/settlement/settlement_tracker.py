"""结算记录追踪：报价生成时自动记录，结算时读取未结算条目。"""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

from quote_system.generator import QuoteRequest
from quote_system.memoq_html import MemoqStats, quote_words

TRACKER_DIR = Path(r"D:\baojia\electron\outputs") / "settlement_tracker"


@dataclass
class SettlementRecord:
    project_key: str
    company: str
    req_name: str
    word_count: int
    total_price: float
    delivery_date: str | None
    quote_date: str
    quote_file: str
    language: str
    settled: bool = False
    settlement_month: str | None = None
    source: str = "manual"
    source_chars: int = 0
    billable_words: int = 0


def _ensure_dir(project_key: str) -> Path:
    p = TRACKER_DIR / project_key
    p.mkdir(parents=True, exist_ok=True)
    return p


def _records_path(project_key: str) -> Path:
    return _ensure_dir(project_key) / "records.json"


def load_records(project_key: str) -> list[dict]:
    path = _records_path(project_key)
    if not path.exists():
        return []
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def save_records(project_key: str, records: list[dict]) -> None:
    path = _records_path(project_key)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)


def add_record(record: SettlementRecord) -> None:
    records = load_records(record.project_key)
    records.append(asdict(record))
    save_records(record.project_key, records)


def mark_settled(project_key: str, settlement_month: str, *, record_ids: list[int] | None = None) -> None:
    """标记指定记录为已结算。"""
    records = load_records(project_key)
    for i, rec in enumerate(records):
        if record_ids is None or i in record_ids:
            rec["settled"] = True
            rec["settlement_month"] = settlement_month
    save_records(project_key, records)


def get_unsettled(project_key: str, target_year: int, target_month: int) -> list[dict]:
    """获取指定月份已交付但未结算的记录。"""
    records = load_records(project_key)
    result = []
    for rec in records:
        if rec.get("settled"):
            continue
        dd = rec.get("delivery_date")
        if not dd:
            continue
        try:
            dt = datetime.strptime(dd, "%Y-%m-%d").date()
        except (ValueError, TypeError):
            continue
        if dt.year == target_year and dt.month == target_month:
            result.append(rec)
    result.sort(key=lambda r: r.get("delivery_date", ""))
    return result


def record_from_quote(request: QuoteRequest, stats_list: list[MemoqStats], output_path: Path, source: str = "manual") -> SettlementRecord:
    """从报价请求和统计信息创建结算记录。"""
    project = request.project

    total_words = 0
    total_price_tax = 0.0
    total_source_chars = 0
    total_billable = 0

    for i, stats in enumerate(stats_list):
        if request.per_file_languages:
            languages = request.per_file_languages[i]
        else:
            languages = request.languages or list(project.default_languages)

        total_source_chars += stats.all_row.source_chars or 0

        for lang in languages:
            if lang == "摘字":
                words = stats.all_row.source_chars or 0
            else:
                words = quote_words(stats)
                total_billable += words

            price_key = lang
            unit_price = 0
            if request.price_overrides and price_key in request.price_overrides:
                unit_price = request.price_overrides[price_key]
            else:
                unit_price = project.prices.get(price_key, 0)

            total_words += words
            total_price_tax += words * unit_price

    if request.delivery_date and isinstance(request.delivery_date, (date, datetime)):
        deliv_str = request.delivery_date.strftime("%Y-%m-%d")
    elif request.delivery_date:
        deliv_str = str(request.delivery_date)
    else:
        deliv_str = None

    return SettlementRecord(
        project_key=project.key,
        company=project.company,
        req_name=output_path.stem,
        word_count=total_words,
        total_price=round(total_price_tax, 2),
        delivery_date=deliv_str,
        quote_date=request.quote_date.strftime("%Y-%m-%d"),
        quote_file=output_path.name,
        language=" / ".join(request.languages or project.default_languages),
        source=source,
        source_chars=total_source_chars,
        billable_words=total_billable,
    )


def quick_record(
    project_key: str,
    company: str,
    req_name: str,
    word_count: int,
    total_price: float,
    quote_file: str,
    language: str,
    delivery_date: str | None = None,
    source: str = "auto",
    source_chars: int = 0,
    billable_words: int = 0,
) -> SettlementRecord:
    return SettlementRecord(
        project_key=project_key,
        company=company,
        req_name=req_name,
        word_count=word_count,
        total_price=total_price,
        delivery_date=delivery_date,
        quote_date=date.today().strftime("%Y-%m-%d"),
        quote_file=quote_file,
        language=language,
        source=source,
        source_chars=source_chars,
        billable_words=billable_words,
    )


def archive_settled(months: int = 3) -> int:
    """将超过 N 个月的已结算记录移入归档目录。返回归档条数。"""
    if not TRACKER_DIR.exists():
        return 0
    cutoff = datetime.now()
    archived = 0
    for proj_dir in TRACKER_DIR.iterdir():
        if not proj_dir.is_dir() or proj_dir.name == "archives":
            continue
        records_path = proj_dir / "records.json"
        if not records_path.exists():
            continue
        try:
            records = json.loads(records_path.read_text(encoding="utf-8"))
        except Exception:
            continue
        keep = []
        to_archive = []
        for rec in records:
            if not rec.get("settled"):
                keep.append(rec)
                continue
            sm = rec.get("settlement_month", "")
            try:
                dt = datetime.strptime(sm, "%Y年%m月")
                age_months = (cutoff.year - dt.year) * 12 + (cutoff.month - dt.month)
                if age_months >= months:
                    to_archive.append(rec)
                else:
                    keep.append(rec)
            except (ValueError, TypeError):
                keep.append(rec)
        if to_archive:
            archive_dir = TRACKER_DIR / "archives" / proj_dir.name
            archive_dir.mkdir(parents=True, exist_ok=True)
            archive_path = archive_dir / "records.json"
            existing = []
            if archive_path.exists():
                try:
                    existing = json.loads(archive_path.read_text(encoding="utf-8"))
                except Exception:
                    pass
            existing.extend(to_archive)
            archive_path.write_text(json.dumps(existing, ensure_ascii=False, indent=2), encoding="utf-8")
            records_path.write_text(json.dumps(keep, ensure_ascii=False, indent=2), encoding="utf-8")
            archived += len(to_archive)
    return archived
