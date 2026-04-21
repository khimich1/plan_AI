import core.config_and_data as cfg
from core.ocr_gpt import parse_gpt_response
from core.plate_line_parser import parse_line
from core.plate_text_normalizer import normalize_plate_prefixes


def test_normalize_prefix_with_dot_and_comma():
    assert normalize_plate_prefixes("ПБ.19,6-12-10 7") == "ПБ 19,6-12-10 7"
    assert normalize_plate_prefixes("ПБ,19,6-12-10 7") == "ПБ 19,6-12-10 7"
    assert normalize_plate_prefixes("ПБ19,6-12-10 7") == "ПБ 19,6-12-10 7"
    assert normalize_plate_prefixes("ПВ.19,6-12-10 7") == "ПБ 19,6-12-10 7"


def test_parse_line_tolerant_pbpk():
    result = parse_line("ПБ.19,6-12-10 7")
    assert result.parsed is True
    assert result.stage == "tolerant_pbpk"
    assert result.length_m == 1.96
    assert result.width_m == 1.2
    assert result.qty == 7
    assert result.load_code == 10.0


def test_parse_gpt_response_new_contract():
    response = """
    [
      {
        "raw_name": "ПБ.19,6-12-10",
        "normalized_candidate": "ПБ 19,6-12-10",
        "qty": "7",
        "confidence": 0.8,
        "issues": ["prefix_separator_dot"]
      }
    ]
    """
    parsed = parse_gpt_response(response)
    assert len(parsed) == 1
    assert parsed[0]["raw_name"] == "ПБ.19,6-12-10"
    assert parsed[0]["normalized_candidate"] == "ПБ 19,6-12-10"
    assert parsed[0]["qty"] == 7
    assert parsed[0]["confidence"] == 0.8
    assert parsed[0]["issues"] == ["prefix_separator_dot"]


def test_set_plate_lists_diagnostics_for_unparsed():
    unparsed, _, _ = cfg.set_plate_lists_from_text("непонятная строка")
    assert len(unparsed) == 1
    diagnostics = cfg.get_last_parse_diagnostics()
    assert diagnostics
    assert diagnostics[0]["validation_status"] == "failed"
    assert diagnostics[0]["reason_code"] in {"pattern_not_matched", "empty_line"}
