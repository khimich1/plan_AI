import pytest
from core.parsers.text_parser import parse_order_from_text, length_dm_to_m, normalize_dimension
from core.models.plate import PlateOrder
from core.exceptions import PlateParseError

def test_length_dm_to_m():
    assert length_dm_to_m("38") == 3.78
    assert length_dm_to_m("38,0") == 3.8
    assert length_dm_to_m("75.5") == 7.55
    assert length_dm_to_m("59,8") == 5.98

def test_normalize_dimension():
    assert normalize_dimension("1,2") == 1.2
    assert normalize_dimension("900") == 9.0 # 900/100
    assert normalize_dimension("32") == 0.32
    assert normalize_dimension("86") == 0.86

def test_parse_order_from_text_wxl_format():
    text = "7.2x1.2 5шт"
    order = parse_order_from_text(text)
    
    assert len(order.items) == 1
    assert order.items[0].length_m == 7.2
    assert order.items[0].width_m == 1.2
    assert order.items[0].quantity == 5
    assert order.items[0].width_category == "1.2"
    assert len(order.unparsed_lines) == 0

def test_parse_order_from_text_pb_format():
    text = "Плиты ПБ 33,1-3.2-8п 3"
    order = parse_order_from_text(text)
    
    assert len(order.items) == 1
    assert order.items[0].length_m == 3.31
    assert order.items[0].width_m == 0.32
    assert order.items[0].quantity == 3
    assert order.items[0].load_code == 8.0
    assert order.items[0].raw_length_dm == "33,1"
    assert order.items[0].width_category == "0.32"

def test_parse_order_1_5m_split():
    text = "ПБ 50-15-8п 2шт"
    order = parse_order_from_text(text)
    
    assert len(order.items) == 2
    
    # Плита 1.2
    assert order.items[0].width_m == 1.2
    assert order.items[0].width_category == "1.2"
    assert order.items[0].quantity == 2
    
    # Плита 0.3 (категория 0.32)
    assert order.items[1].width_m == 0.3
    assert order.items[1].width_category == "0.32"
    assert order.items[1].quantity == 2

def test_parse_order_empty():
    with pytest.raises(PlateParseError):
        parse_order_from_text("")
    
    with pytest.raises(PlateParseError):
        parse_order_from_text("   \n   ")

def test_parse_order_invalid_line():
    text = "Плиты ПБ 33,1-3.2-8п 3\nНепонятный текст"
    order = parse_order_from_text(text)
    
    assert len(order.items) == 1
    assert len(order.unparsed_lines) == 1
    assert order.unparsed_lines[0] == "Непонятный текст"
