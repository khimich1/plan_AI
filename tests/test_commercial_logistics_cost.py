from core.cargo_delivery_pricing import (
    CARGO_DELIVERY_TRUCK_CAPACITY_KG,
    cargo_delivery_trips_count,
    delivery_service_charge_rub,
    total_order_cargo_weight_kg,
)
from core.commercial_offer import calculate_total_cost as calculate_total_cost_pdf
from core.commercial_offer_xlsx import calculate_total_cost as calculate_total_cost_xlsx
from core.kp_plate_weight import resolve_kp_line_weight_kg


def test_total_order_cargo_weight_kg_empty_and_sum_matches_lines() -> None:
    assert total_order_cargo_weight_kg([]) == 0.0
    order_data = [
        {"name": "A", "qty": 2, "length_m": 1.0, "width_m": 1.0},
        {"name": "B", "qty": 1, "length_m": 1.0, "width_m": 0.6},
    ]
    expected = sum(resolve_kp_line_weight_kg(item)[1] for item in order_data)
    assert total_order_cargo_weight_kg(order_data) == expected


def test_cargo_delivery_trips_count() -> None:
    assert cargo_delivery_trips_count(0) == 0
    assert cargo_delivery_trips_count(1) == 1
    assert cargo_delivery_trips_count(CARGO_DELIVERY_TRUCK_CAPACITY_KG) == 1
    assert cargo_delivery_trips_count(CARGO_DELIVERY_TRUCK_CAPACITY_KG + 0.01) == 2
    # Ровно два полных рейса — без «лишнего» третьего рейса.
    assert cargo_delivery_trips_count(2 * CARGO_DELIVERY_TRUCK_CAPACITY_KG) == 2
    # Дробный вес: порог второго рейса по ceil, не по целым килограммам.
    assert cargo_delivery_trips_count(CARGO_DELIVERY_TRUCK_CAPACITY_KG + 0.001) == 2
    assert cargo_delivery_trips_count(-50.0) == 0


def test_delivery_service_charge_rub() -> None:
    assert delivery_service_charge_rub(100.0, 1000.0) == 100.0
    assert delivery_service_charge_rub(100.0, CARGO_DELIVERY_TRUCK_CAPACITY_KG + 1.0) == 200.0
    assert delivery_service_charge_rub(0.0, 50000.0) == 0.0
    assert delivery_service_charge_rub(500.0, 0.0) == 0.0
    assert delivery_service_charge_rub(50.55, CARGO_DELIVERY_TRUCK_CAPACITY_KG + 1.0) == 101.10


def test_calculate_total_cost_single_trip_when_under_truck_capacity() -> None:
    # 65 × ~283.3 кг < 18600 ⇒ один рейс при logistics_cost=100.
    order_data = [
        {
            "name": "ПБ",
            "qty": 65,
            "unit_price": 10.0,
            "length_m": 1.0,
            "width_m": 1.0,
        }
    ]
    totals = calculate_total_cost_xlsx(order_data, discount_percent=0, logistics_cost=100.0)
    assert totals["total_with_vat"] == 750.0  # 650 плиты + 100 доставка
    assert totals["vat_amount"] == 143.0
    assert totals["subtotal"] == 607.0


def test_calculate_total_cost_applies_discount_only_to_products() -> None:
    order_data = [
        {
            "name": "ПБ 59-12-8п",
            "qty": 1,
            "unit_price": 122.0,
            "length_m": 1.0,
            "width_m": 1.0,
        }
    ]

    totals = calculate_total_cost_xlsx(order_data, discount_percent=50, logistics_cost=100.0)

    assert totals["total_qty"] == 1
    # Цены с НДС: плиты 122*0.5=61 + доставка 100×1 рейс = 161; НДС 22% только от плит: 61*0.22=13.42
    assert totals["total_with_vat"] == 161.0
    assert totals["vat_amount"] == 13.42
    assert totals["subtotal"] == 147.58


def test_calculate_total_cost_scales_delivery_by_cargo_trips() -> None:
    # Вес строки ≈283.3 кг × 66 > 18600 кг ⇒ 2 рейса; доставка 100×2.
    order_data = [
        {
            "name": "ПБ",
            "qty": 66,
            "unit_price": 10.0,
            "length_m": 1.0,
            "width_m": 1.0,
        }
    ]
    totals = calculate_total_cost_xlsx(order_data, discount_percent=0, logistics_cost=100.0)
    assert totals["total_with_vat"] == 860.0  # 660 плиты + 200 доставка
    assert totals["vat_amount"] == 145.2
    assert totals["subtotal"] == 714.8


def test_calculate_total_cost_ignores_negative_logistics_cost() -> None:
    order_data = [
        {
            "name": "ПБ 59-12-8п",
            "qty": 1,
            "unit_price": 122.0,
        }
    ]

    totals = calculate_total_cost_xlsx(order_data, discount_percent=0, logistics_cost=-500.0)

    assert totals["total_with_vat"] == 122.0
    assert totals["vat_amount"] == 26.84
    assert totals["subtotal"] == 95.16


def test_pdf_and_xlsx_calculate_total_cost_agree_for_sample_orders() -> None:
    """Два модуля дублируют calculate_total_cost — итоги по доставке не должны расходиться."""
    fixtures = [
        (
            [
                {"name": "ПБ", "qty": 2, "unit_price": 10.0, "length_m": 1.2, "width_m": 1.0},
            ],
            10,
            1500.5,
        ),
        (
            [
                {"name": "ПБ 59-12-8п", "qty": 3, "unit_price": 100.0, "length_m": 1.0, "width_m": 1.0},
            ],
            0,
            0.0,
        ),
    ]
    for order_data, discount, logistics in fixtures:
        assert calculate_total_cost_pdf(order_data, discount, logistics_cost=logistics) == (
            calculate_total_cost_xlsx(order_data, discount, logistics_cost=logistics)
        )
