import pytest
from quote_system.memoq_html import (
    MemoqRow,
    MemoqStats,
    parse_memoq_html,
    quote_words,
    stats_to_output_values,
    decode_html,
    normalize_label,
    parse_int,
    is_same_as_above,
    pick_stats_table,
    map_headers,
    build_row,
    HEADER_ALIASES,
    EXPECTED_ORDER,
)


# ========== 数据类单元测试 ==========

class TestMemoqRow:
    def test_should_create_row_with_all_fields(self):
        row = MemoqRow(
            type="All",
            segments=100,
            source_words=5000,
            source_non_asian_words=4800,
            source_asian_characters=4500,
            source_chars=5000,
            source_tags=20,
            percent="100%",
        )
        assert row.type == "All"
        assert row.segments == 100
        assert row.source_chars == 5000

    def test_should_create_row_with_none_fields(self):
        row = MemoqRow(type="No Match", segments=None, source_chars=None)
        assert row.type == "No Match"
        assert row.segments is None
        assert row.source_chars is None

    def test_as_a_to_h_should_return_correct_order(self):
        row = MemoqRow(
            type="All", segments=10, source_words=100,
            source_non_asian_words=90, source_asian_characters=80,
            source_chars=85, source_tags=5, percent="100%"
        )
        result = row.as_a_to_h()
        assert result == ["All", 10, 100, 90, 80, 85, 5, "100%"]

    def test_as_b_to_h_should_exclude_type(self):
        row = MemoqRow(
            type="All", segments=10, source_words=100,
            source_non_asian_words=90, source_asian_characters=80,
            source_chars=85, source_tags=5, percent="100%"
        )
        result = row.as_b_to_h()
        assert result == [10, 100, 90, 80, 85, 5, "100%"]


# ========== parse_int 测试 ==========

class TestParseInt:
    def test_should_parse_plain_int(self):
        assert parse_int("123") == 123

    def test_should_parse_float_string(self):
        assert parse_int("123.45") == 123

    def test_should_parse_comma_separated(self):
        assert parse_int("1,234") == 1234

    def test_should_parse_chinese_comma(self):
        assert parse_int("1，234") == 1234

    def test_should_parse_with_percent(self):
        assert parse_int("50%") == 50

    def test_should_return_none_for_none(self):
        assert parse_int(None) is None

    def test_should_return_none_for_empty(self):
        assert parse_int("") is None

    def test_should_return_none_for_dash(self):
        assert parse_int("-") is None

    def test_should_return_none_for_non_numeric(self):
        assert parse_int("abc") is None


# ========== normalize_label 测试 ==========

class TestNormalizeLabel:
    def test_should_lowercase(self):
        assert normalize_label("ALL") == "all"

    def test_should_collapse_whitespace(self):
        assert normalize_label("  all   row  ") == "all row"

    def test_should_handle_fullwidth_space(self):
        assert normalize_label("all\u3000row") == "all row"


# ========== is_same_as_above 测试 ==========

class TestIsSameAsAbove:
    def test_should_match_chinese(self):
        assert is_same_as_above("同上") is True

    def test_should_match_english(self):
        assert is_same_as_above("Same as above") is True

    def test_should_match_ditto(self):
        assert is_same_as_above("ditto") is True

    def test_should_not_match_other(self):
        assert is_same_as_above("random text") is False


# ========== decode_html 测试 ==========

class TestDecodeHtml:
    def test_should_decode_utf8(self):
        raw = "<html>测试</html>".encode("utf-8")
        assert decode_html(raw) == "<html>测试</html>"

    def test_should_decode_utf8_sig(self):
        raw = b"\xef\xbb\xbf<html>test</html>"
        assert decode_html(raw) == "<html>test</html>"

    def test_should_decode_utf16(self):
        raw = "<html>test</html>".encode("utf-16")
        assert "test" in decode_html(raw)

    def test_should_fallback_to_replace(self):
        raw = b"\xff\xfe\x00\x01\x02"
        result = decode_html(raw)
        assert len(result) > 0


# ========== map_headers 测试 ==========

class TestMapHeaders:
    def test_should_map_standard_headers(self):
        header_row = ["Type", "Segments", "Source Words", "Percent"]
        result = map_headers(header_row)
        assert result["type"] == 0
        assert result["segments"] == 1
        assert result["source_words"] == 2
        assert result["percent"] == 3

    def test_should_map_japanese_headers(self):
        header_row = ["種類", "セグメント数", "ソースの単語数", "割合"]
        result = map_headers(header_row)
        assert result["type"] == 0
        assert result["segments"] == 1

    def test_should_map_chinese_headers(self):
        header_row = ["类型", "句段数", "原文字数", "比例"]
        result = map_headers(header_row)
        assert result["type"] == 0
        assert result["segments"] == 1

    def test_should_fill_missing_with_expected_order(self):
        header_row = ["Type"]
        result = map_headers(header_row)
        assert result["type"] == 0
        assert result["segments"] == 1  # 从 EXPECTED_ORDER 回退


# ========== build_row 测试 ==========

class TestBuildRow:
    def test_should_build_valid_row(self):
        headers = map_headers(["Type", "Segments", "Source Words", "Percent", "Source Chars"])
        raw = ["All", "100", "5000", "100%", "4800"]
        row = build_row(raw, headers)
        assert row is not None
        assert row.type == "All"
        assert row.segments == 100
        assert row.source_chars == 4800

    def test_should_return_none_for_empty_type(self):
        headers = map_headers(["Type", "Segments"])
        raw = ["", "100"]
        row = build_row(raw, headers)
        assert row is None

    def test_should_handle_missing_columns(self):
        headers = map_headers(["Type", "Segments"])
        raw = ["All", "100"]
        row = build_row(raw, headers)
        assert row is not None
        assert row.type == "All"
        assert row.segments == 100


# ========== MemoqStats 测试 ==========

class TestMemoqStats:
    def _make_stats(self, rows_data):
        from pathlib import Path
        rows = [MemoqRow(type=r[0], segments=r[1], source_asian_characters=r[2], source_chars=r[3])
                for r in rows_data]
        return MemoqStats(source_path=Path(__file__), title="test", rows=rows)

    def test_all_row_should_find_all(self):
        stats = self._make_stats([
            ("All", 100, 4500, 5000),
            ("No Match", 30, 1000, 1200),
        ])
        assert stats.all_row.type == "All"

    def test_all_row_should_match_japanese(self):
        stats = self._make_stats([
            ("すべて", 100, 4500, 5000),
        ])
        assert stats.all_row.type == "すべて"

    def test_all_row_should_match_chinese(self):
        stats = self._make_stats([
            ("全部", 100, 4500, 5000),
        ])
        assert stats.all_row.type == "全部"

    def test_all_row_should_raise_if_not_found(self):
        stats = self._make_stats([
            ("No Match", 30, 1000, 1200),
        ])
        with pytest.raises(ValueError):
            stats.all_row

    def test_repetition_row_should_find(self):
        stats = self._make_stats([
            ("All", 100, 4500, 5000),
            ("Repetition", 20, 200, 250),
        ])
        assert stats.repetition_row.type == "Repetition"

    def test_no_match_row_should_find(self):
        stats = self._make_stats([
            ("All", 100, 4500, 5000),
            ("No Match", 30, 1000, 1200),
        ])
        assert stats.no_match_row.type == "No Match"


# ========== quote_words 测试 ==========

class TestQuoteWords:
    def test_should_calculate_standard_case(self):
        from pathlib import Path
        rows = [
            MemoqRow(type="All", source_asian_characters=4500, source_chars=5000),
            MemoqRow(type="Repetition", source_asian_characters=500),
        ]
        stats = MemoqStats(source_path=Path(__file__), title="test", rows=rows)
        # effective = 4500 - 500 = 4000, punctuation = 5000 - 4500 = 500
        # result = 4000 + 500 * 0.3 = 4000 + 150 = 4150
        assert quote_words(stats) == 4150

    def test_should_calculate_without_repetition(self):
        from pathlib import Path
        rows = [
            MemoqRow(type="All", source_asian_characters=3000, source_chars=3500),
            MemoqRow(type="Repetition", source_asian_characters=0),
        ]
        stats = MemoqStats(source_path=Path(__file__), title="test", rows=rows)
        # effective = 3000, punctuation = 500
        # result = 3000 + 500 * 0.3 = 3150
        assert quote_words(stats) == 3150

    def test_should_raise_when_missing_columns(self):
        from pathlib import Path
        rows = [
            MemoqRow(type="All", source_asian_characters=None, source_chars=None),
        ]
        stats = MemoqStats(source_path=Path(__file__), title="test", rows=rows)
        with pytest.raises(ValueError):
            quote_words(stats)


# ========== stats_to_output_values 测试 ==========

class TestStatsToOutputValues:
    def test_should_return_correct_order(self):
        from pathlib import Path
        rows = [
            MemoqRow(type="All", segments=100, source_words=5000,
                     source_non_asian_words=4800, source_asian_characters=4500,
                     source_chars=5000, source_tags=20, percent="100%"),
            MemoqRow(type="Context", source_asian_characters=100),
            MemoqRow(type="Repetition", source_asian_characters=200),
            MemoqRow(type="101%", source_asian_characters=300),
            MemoqRow(type="100%", source_asian_characters=400),
            MemoqRow(type="95% - 99%", source_asian_characters=500),
            MemoqRow(type="85% - 94%", source_asian_characters=600),
            MemoqRow(type="75% - 84%", source_asian_characters=700),
            MemoqRow(type="50% - 74%", source_asian_characters=800),
            MemoqRow(type="No Match", source_asian_characters=900),
        ]
        stats = MemoqStats(source_path=Path(__file__), title="test", rows=rows)
        result = stats_to_output_values(stats)
        assert result == [5000, 4500, 100, 200, 300, 400, 500, 600, 700, 800, 900]


# ========== HEADER_ALIASES 完整性检查 ==========

class TestHeaderAliases:
    def test_should_cover_all_expected_order_keys(self):
        for key in EXPECTED_ORDER:
            assert key in HEADER_ALIASES, f"Missing HEADER_ALIASES entry for: {key}"
