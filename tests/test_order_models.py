import pytest

try:
    from core.models.plate import PlateItem, PlateOrder
except ImportError as exc:
    pytest.skip(
        f"Архивные тесты моделей плит: нет core.models.plate (OPT-010): {exc}",
        allow_module_level=True,
    )

def test_plate_order_to_dict():
    order = PlateOrder()
    order.items.append(PlateItem(length_m=5.5, width_m=1.2, quantity=2, load_code=8.0, raw_length_dm="55", width_category="1.2"))
    order.unparsed_lines.append("invalid line")
    
    data = order.to_dict()
    
    assert "items" in data
    assert "unparsed_lines" in data
    assert len(data["items"]) == 1
    assert data["items"][0]["length_m"] == 5.5
    assert data["items"][0]["width_m"] == 1.2
    assert data["items"][0]["quantity"] == 2
    assert data["items"][0]["load_code"] == 8.0
    assert data["items"][0]["raw_length_dm"] == "55"
    assert data["items"][0]["width_category"] == "1.2"
    assert data["unparsed_lines"] == ["invalid line"]

def test_plate_order_from_dict():
    data = {
        "items": [
            {
                "length_m": 7.2,
                "width_m": 1.0,
                "quantity": 3,
                "load_code": 10.0,
                "raw_length_dm": "72",
                "width_category": "1.0"
            }
        ],
        "unparsed_lines": ["bad format"]
    }
    
    order = PlateOrder.from_dict(data)
    
    assert len(order.items) == 1
    assert isinstance(order.items[0], PlateItem)
    assert order.items[0].length_m == 7.2
    assert order.items[0].width_m == 1.0
    assert order.items[0].quantity == 3
    assert order.items[0].load_code == 10.0
    assert order.items[0].raw_length_dm == "72"
    assert order.items[0].width_category == "1.0"
    assert order.unparsed_lines == ["bad format"]

def test_plate_order_properties():
    order = PlateOrder()
    order.items.append(PlateItem(length_m=5.0, width_m=1.2, quantity=2, width_category="1.2"))
    order.items.append(PlateItem(length_m=3.0, width_m=1.08, quantity=4, width_category="1.08"))
    
    assert order.total_plates == 6
    assert order.scrap_strips_0_12_m_total == 12.0  # 3.0 * 4 = 12.0
