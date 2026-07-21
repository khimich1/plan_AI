import pytest

from core.plate_line_parser import BARE_PLATE_LINE_RE, build_lwh_mm_load_warning, match_bare_plate_line, parse_line


USER_BARE_ORDER = """
71-12-8  3
71-9,20-8   1
69-12-8  13
69-6,65-10   4
69-7,20-10   2
69-12-12,5  7
65,6-12-12,5  2
65-12-12,5  4
65-12-8  5
64,2-12-8  1
55-12-8  2
55-12-10  1
55-7,2-8  1
55-6,65-8   1
34,2-12-8  1
32,5-12-8  5
32,5-6,65-8  1
31,5-12-8  11
31,5-10,8-8  3
20,3-12-8  7
19,1-12-8  2
64,2-12-12,5  4
64,2-9,20-8   1
71-12-12  3
70,6-12-10  1
""".strip()


@pytest.mark.parametrize(
    "line,length_m,width_m,load_code,qty",
    [
        ("71-12-8  3", 7.1, 1.2, 8.0, 3),
        ("71-9,20-8   1", 7.1, 0.92, 8.0, 1),
        ("65,6-12-12,5  2", 6.56, 1.2, 12.5, 2),
        ("32,5-6,65-8  1", 3.25, 0.665, 8.0, 1),
        ("31,5-10,8-8  3", 3.15, 1.08, 8.0, 3),
        ("70,6-12-10  1", 7.06, 1.2, 10.0, 1),
    ],
)
def test_parse_line_bare_mark_key_cases(line, length_m, width_m, load_code, qty):
    result = parse_line(line)
    assert result.parsed is True
    assert result.stage == "bare_lwd"
    assert result.length_m == length_m
    assert result.width_m == width_m
    assert result.load_code == load_code
    assert result.qty == qty


def test_parse_line_bare_mark_full_user_order():
    for raw in USER_BARE_ORDER.splitlines():
        result = parse_line(raw.strip())
        assert result.parsed is True, raw
        assert result.stage == "bare_lwd", raw
        assert result.load_code is not None and result.load_code > 0, raw
        assert result.qty >= 1, raw


def test_parse_line_regression_pb_format():
    result = parse_line("ПБ 78-12-8п 2")
    assert result.parsed is True
    assert result.stage == "tolerant_pbpk"
    assert result.length_m == 7.8
    assert result.width_m == 1.2
    assert result.load_code == 8.0
    assert result.qty == 2


def test_parse_line_regression_wxl_format():
    result = parse_line("0,32x6,63 - 4")
    assert result.parsed is True
    assert result.stage == "strict_wxl"
    assert result.width_m == 0.32
    assert result.length_m == 6.63
    assert result.qty == 4


@pytest.mark.parametrize(
    "line,length_m,width_m,load_code,qty",
    [
        ("3880x1200x220 7", 3.88, 1.2, 8.0, 7),
        ("3880x710x220 1", 3.88, 0.71, 8.0, 1),
        ("4880x1200x220 5", 4.88, 1.2, 8.0, 5),
        ("4880x720x220", 4.88, 0.72, 8.0, 1),
    ],
)
def test_parse_line_lwh_mm_format(line, length_m, width_m, load_code, qty):
    result = parse_line(line)
    assert result.parsed is True
    assert result.stage == "strict_lwh_mm"
    assert result.length_m == length_m
    assert result.width_m == width_m
    assert result.load_code == load_code
    assert result.qty == qty
    assert result.load_assumed is True
    assert result.length_dm_raw


def test_build_lwh_mm_load_warning_lists_source_lines():
    warning = build_lwh_mm_load_warning(["3880x1200x220 7", "4880x720x220"])
    assert "3880x1200x220 7" in warning
    assert "4880x720x220" in warning
    assert "Проверьте нагрузку" in warning


def test_parse_line_regression_tolerant_pb_with_dot():
    result = parse_line("ПБ.19,6-12-10 7")
    assert result.parsed is True
    assert result.stage == "tolerant_pbpk"
    assert result.length_m == 1.96
    assert result.width_m == 1.2
    assert result.load_code == 10.0
    assert result.qty == 7


@pytest.mark.parametrize("line", ["Непонятный текст"])
def test_parse_line_negative_cases(line):
    result = parse_line(line)
    assert result.parsed is False


def test_parse_line_date_like_string_rejected_by_validation():
    from core.plate_validation import validate_plate_values

    result = parse_line("2025-05-25")
    assert result.parsed is True
    assert result.stage == "bare_lwd"
    validation = validate_plate_values(result.width_m, result.length_m, result.qty)
    assert validation.ok is False


def test_parse_line_bare_unrealistic_dimensions_still_matches_pattern():
    result = parse_line("1-2-3")
    assert result.parsed is True
    assert result.stage == "bare_lwd"


def test_match_bare_plate_line_helper():
    matched = match_bare_plate_line("71-12-8 3")
    assert matched == ("71", "12", "8", 3)
    assert match_bare_plate_line("ПБ 78-12-8п 2") is None


def test_bare_plate_line_re_exported():
    assert BARE_PLATE_LINE_RE.match("71-12-8 3") is not None


def test_get_wide_plate_lines_bare_format():
    from core.plate_text_normalizer import get_wide_plate_lines

    result = get_wide_plate_lines("71-15-8 2")
    assert len(result) == 1
    assert result[0] == ("71-15-8 2", 2)
