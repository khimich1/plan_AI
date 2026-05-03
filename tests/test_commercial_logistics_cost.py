from core.commercial_offer_xlsx import calculate_total_cost


def test_calculate_total_cost_applies_discount_only_to_products() -> None:
    order_data = [
        {
            "name": "ПБ 59-12-8п",
            "qty": 1,
            "unit_price": 122.0,
        }
    ]

    totals = calculate_total_cost(order_data, discount_percent=50, logistics_cost=100.0)

    assert totals["total_qty"] == 1
    # Цены с НДС: плиты 122*0.5=61 + логистика 100 = 161; НДС 22% только от плит: 61*0.22=13.42
    assert totals["total_with_vat"] == 161.0
    assert totals["vat_amount"] == 13.42
    assert totals["subtotal"] == 147.58


def test_calculate_total_cost_ignores_negative_logistics_cost() -> None:
    order_data = [
        {
            "name": "ПБ 59-12-8п",
            "qty": 1,
            "unit_price": 122.0,
        }
    ]

    totals = calculate_total_cost(order_data, discount_percent=0, logistics_cost=-500.0)

    assert totals["total_with_vat"] == 122.0
    assert totals["vat_amount"] == 26.84
    assert totals["subtotal"] == 95.16
