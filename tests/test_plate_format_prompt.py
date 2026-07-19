from core.plate_format_prompt import build_plate_parser_system_prompt


def test_plate_parser_system_prompt_contains_all_formats():
    prompt = build_plate_parser_system_prompt()
    assert "tolerant_pbpk" in prompt
    assert "strict_wxl" in prompt
    assert "strict_lwh_mm" in prompt
    assert "bare_lwd" in prompt
    assert "ПБ 78-12-8п 2" in prompt
    assert "0,32x6,63 - 4" in prompt
    assert "3880x1200x220 7" in prompt
    assert "71-12-8 3" in prompt
    assert "normalized_candidate" in prompt
    assert "JSON" in prompt
