from __future__ import annotations

from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path
from typing import Iterable


@dataclass(frozen=True)
class MemoqRow:
    type: str
    segments: int | None = None
    source_words: int | None = None
    source_non_asian_words: int | None = None
    source_asian_characters: int | None = None
    source_chars: int | None = None
    source_tags: int | None = None
    percent: str | None = None

    def as_a_to_h(self) -> list[object]:
        return [
            self.type,
            self.segments,
            self.source_words,
            self.source_non_asian_words,
            self.source_asian_characters,
            self.source_chars,
            self.source_tags,
            self.percent,
        ]

    def as_b_to_h(self) -> list[object]:
        return [
            self.segments,
            self.source_words,
            self.source_non_asian_words,
            self.source_asian_characters,
            self.source_chars,
            self.source_tags,
            self.percent,
        ]


@dataclass(frozen=True)
class MemoqStats:
    source_path: Path
    title: str
    rows: list[MemoqRow]

    @property
    def all_row(self) -> MemoqRow:
        return self.find_row({"all", "すべて", "全部"})

    @property
    def repetition_row(self) -> MemoqRow:
        return self.find_row({"repetition", "繰り返し", "重复"})

    @property
    def no_match_row(self) -> MemoqRow:
        return self.find_row({"no match", "一致しない", "无匹配"})

    def find_row(self, names: set[str]) -> MemoqRow:
        normalized = {normalize_label(name) for name in names}
        for row in self.rows:
            if normalize_label(row.type) in normalized:
                return row
        raise ValueError(f"{self.source_path.name} 里找不到统计行: {', '.join(sorted(names))}")


class _TableCollector(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.tables: list[list[list[str]]] = []
        self._table: list[list[str]] | None = None
        self._row: list[str] | None = None
        self._cell: list[str] | None = None
        self._title_parts: list[str] = []
        self._in_title = False
        self._heading_parts: list[str] = []
        self._headings: list[str] = []
        self._in_heading = False

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag == "title":
            self._in_title = True
        elif tag in {"h1", "h2"}:
            self._in_heading = True
            self._heading_parts = []
        elif tag == "table":
            self._table = []
        elif tag == "tr" and self._table is not None:
            self._row = []
        elif tag in {"td", "th"} and self._row is not None:
            self._cell = []

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag == "title":
            self._in_title = False
        elif tag in {"h1", "h2"} and self._in_heading:
            heading = " ".join("".join(self._heading_parts).split())
            if heading:
                self._headings.append(heading)
            self._in_heading = False
        elif tag in {"td", "th"} and self._cell is not None and self._row is not None:
            text = " ".join("".join(self._cell).split())
            self._row.append(text)
            self._cell = None
        elif tag == "tr" and self._row is not None and self._table is not None:
            if any(cell.strip() for cell in self._row):
                self._table.append(self._row)
            self._row = None
        elif tag == "table" and self._table is not None:
            self.tables.append(self._table)
            self._table = None

    def handle_data(self, data: str) -> None:
        if self._cell is not None:
            self._cell.append(data)
        if self._in_title:
            self._title_parts.append(data)
        if self._in_heading:
            self._heading_parts.append(data)

    @property
    def title(self) -> str:
        title = " ".join("".join(self._title_parts).split())
        if title:
            return title
        return self._headings[0] if self._headings else ""


HEADER_ALIASES = {
    "type": {"type", "类型", "種類"},
    "segments": {"segments", "句段数", "句段", "セグメント数"},
    "source_words": {"source words", "原文字数", "原文字（词）", "ソースの単語数"},
    "source_non_asian_words": {
        "source non-asian words",
        "非亚洲字数",
        "原文非亚洲字（词）",
        "ソースがアジア言語以外の単語数",
    },
    "source_asian_characters": {
        "source asian characters",
        "亚洲字符",
        "亚洲字符原文",
        "ソースがアジア言語の文字数",
    },
    "source_chars": {"source chars", "原文字符", "ソースの文字数"},
    "source_tags": {"source tags", "标签数", "原文标签", "ソースのタグ"},
    "percent": {"percent", "%", "比例", "割合"},
}

EXPECTED_ORDER = [
    "type",
    "segments",
    "source_words",
    "source_non_asian_words",
    "source_asian_characters",
    "source_chars",
    "source_tags",
    "percent",
]


def parse_memoq_html(path: str | Path) -> MemoqStats:
    html_path = Path(path)
    raw = html_path.read_bytes()
    text = decode_html(raw)
    parser = _TableCollector()
    parser.feed(text)

    table, header_index = pick_stats_table(parser.tables)
    headers = map_headers(table[header_index])
    data_rows = table[header_index + 1 :]
    rows = [row for row in (build_row(raw_row, headers) for raw_row in data_rows) if row is not None]
    if not rows:
        raise ValueError(f"{html_path.name} 没有解析到 MemoQ 统计数据")

    title = parser.title or html_path.stem
    return MemoqStats(source_path=html_path, title=title, rows=rows)


def decode_html(raw: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-16", "gb18030", "shift_jis", "cp932"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def pick_stats_table(tables: Iterable[list[list[str]]]) -> tuple[list[list[str]], int]:
    for table in tables:
        for index, row in enumerate(table):
            normalized = {normalize_label(cell) for cell in row}
            if has_header(normalized, "source_asian_characters") and has_header(normalized, "source_chars"):
                return table, index
            if has_header(normalized, "segments") and has_header(normalized, "percent"):
                return table, index
    raise ValueError("HTML 中没有找到 MemoQ 字数统计表")


def has_header(normalized_headers: set[str], key: str) -> bool:
    return bool(normalized_headers & {normalize_label(alias) for alias in HEADER_ALIASES[key]})


def map_headers(header_row: list[str]) -> dict[str, int]:
    result: dict[str, int] = {}
    for idx, header in enumerate(header_row):
        normalized = normalize_label(header)
        for key, aliases in HEADER_ALIASES.items():
            if normalized in {normalize_label(alias) for alias in aliases}:
                result[key] = idx
    # Some MemoQ exports omit the Type header but still keep the standard column order.
    for idx, key in enumerate(EXPECTED_ORDER):
        result.setdefault(key, idx)
    return result


def build_row(raw_row: list[str], headers: dict[str, int]) -> MemoqRow | None:
    def value(key: str) -> str | None:
        idx = headers[key]
        if idx >= len(raw_row):
            return None
        cell = raw_row[idx].strip()
        return cell if cell else None

    row_type = value("type")
    if not row_type:
        return None

    return MemoqRow(
        type=row_type,
        segments=parse_int(value("segments")),
        source_words=parse_int(value("source_words")),
        source_non_asian_words=parse_int(value("source_non_asian_words")),
        source_asian_characters=parse_int(value("source_asian_characters")),
        source_chars=parse_int(value("source_chars")),
        source_tags=parse_int(value("source_tags")),
        percent=value("percent"),
    )


def parse_int(value: str | None) -> int | None:
    if value is None:
        return None
    cleaned = value.replace(",", "").replace("，", "").replace("%", "").strip()
    if cleaned in {"", "-"}:
        return None
    try:
        return int(float(cleaned))
    except ValueError:
        return None


def normalize_label(value: str) -> str:
    return " ".join(value.replace("　", " ").strip().lower().split())


def is_same_as_above(value: str) -> bool:
    normalized = normalize_label(value)
    return normalized in {"同上", "同 上", "same as above", "same", "ditto"}


def source_asian_for(stats: MemoqStats, labels: set[str]) -> int:
    normalized = {normalize_label(label) for label in labels}
    for row in stats.rows:
        if normalize_label(row.type) in normalized:
            return row.source_asian_characters or 0
    return 0


def stats_to_output_values(stats: MemoqStats) -> list[int]:
    all_row = stats.all_row
    return [
        all_row.source_chars or 0,
        all_row.source_asian_characters or 0,
        source_asian_for(stats, {
            "context",
            "コンテキスト",
            "上下文",
            "クロスファイル",
            "クロス翻訳",
            "x 翻译/双重上下文",
            "x翻译/双重上下文",
            "双重上下文",
            "双重",
        }),
        source_asian_for(stats, {"repetition", "繰り返し", "重复", "完全一致", "完全匹配", "exact"}),
        source_asian_for(stats, {"101%", "101"}),
        source_asian_for(stats, {"100%", "100"}),
        source_asian_for(stats, {"95%-99%", "95% - 99%", "95-99"}),
        source_asian_for(stats, {"85%-94%", "85% - 94%", "85-94"}),
        source_asian_for(stats, {"75%-84%", "75% - 84%", "75-84"}),
        source_asian_for(stats, {"50%-74%", "50% - 74%", "50-74"}),
        source_asian_for(stats, {"no match", "一致しない", "无匹配", "0%-49%", "0% - 49%", "0-49"}),
    ]


def quote_words(stats: MemoqStats) -> int:
    all_row = stats.all_row
    repetition_row = stats.repetition_row
    if all_row.source_asian_characters is None or all_row.source_chars is None:
        raise ValueError(f"{stats.source_path.name} 的 All 行缺少 E/F 列字数")
    repeated = repetition_row.source_asian_characters or 0
    effective_without_punctuation = all_row.source_asian_characters - repeated
    punctuation = all_row.source_chars - all_row.source_asian_characters
    return round(effective_without_punctuation + punctuation * 0.3)
