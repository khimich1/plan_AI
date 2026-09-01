from core.cargo_delivery_pricing import (
    CARGO_DELIVERY_TRUCK_CAPACITY_KG,
    cargo_delivery_trips_count,
    delivery_service_charge_rub,
    total_order_cargo_weight_kg,
)
from core.commercial_offer import calculate_total_cost as calculate_total_cost_pdf
from core.commercial_offer_xlsx import calculate_total_cost as calculate_total_cost_xlsx
from core.commercial_pricing import calculate_total_cost as calculate_total_cost_core
from core.kp_plate_weight import resolve_kp_line_weight_kg
from core.pile_catalog import PileCatalogEntry, upsert_pile_catalog
from tests.helpers import kp_db_fixtures as fx


def test_total_order_cargo_weight_kg_empty_and_sum_matches_lines() -> None:
    assert total_order_cargo_weight_kg([]) == 0.0
    order_data = [
        {"name": "A", "qty": 2, "length_m": 1.0, "width_m": 1.0},
        {"name": "B", "qty": 1, "length_m": 1.0, "width_m": 0.6},
    ]
    expected = sum(resolve_kp_line_weight_kg(item)[1] for item in order_data)
    assert total_order_cargo_weight_kg(order_data) == expected
    # Backward compatible: explicit None == no filter (all lines).
    assert total_order_cargo_weight_kg(order_data, product_types=None) == expected


def test_total_order_cargo_weight_kg_plates_filter_empty_is_zero() -> None:
    assert total_order_cargo_weight_kg([], product_types={"plates"}) == 0.0


def test_total_order_cargo_weight_kg_plates_filter_all_non_plates_is_zero() -> None:
    order_data = [
        {
            "name": "С30.15-3",
            "product_type": "piles",
            "qty": 10,
            "length_m": 3.0,
            "width_m": 0.3,
        },
        {
            "name": "ЛС-12",
            "product_type": "steps",
            "qty": 5,
            "length_m": 1.2,
            "width_m": 0.3,
        },
    ]
    # Without filter non-plates still have formula weight; with plates filter → 0.
    assert total_order_cargo_weight_kg(order_data) > 0.0
    assert total_order_cargo_weight_kg(order_data, product_types={"plates"}) == 0.0


def test_total_order_cargo_weight_kg_plates_filter_mixed_sums_plates_only() -> None:
    plate = {
        "name": "ПБ 59-12-8п",
        "product_type": "plates",
        "qty": 2,
        "length_m": 1.0,
        "width_m": 1.0,
    }
    pile = {
        "name": "С30.15-3",
        "product_type": "piles",
        "qty": 10,
        "length_m": 3.0,
        "width_m": 0.3,
    }
    order_data = [plate, pile]
    plates_only = resolve_kp_line_weight_kg(plate)[1]
    all_lines = plates_only + resolve_kp_line_weight_kg(pile)[1]

    assert total_order_cargo_weight_kg(order_data) == all_lines
    assert total_order_cargo_weight_kg(order_data, product_types={"plates"}) == plates_only
    assert plates_only > 0.0
    assert all_lines > plates_only


def test_total_order_cargo_weight_kg_missing_product_type_treated_as_plates() -> None:
    """Legacy mono lines without product_type count as plates when filtering."""
    legacy_plate = {"name": "ПБ", "qty": 1, "length_m": 1.0, "width_m": 1.0}
    explicit_pile = {
        "name": "С30.15-3",
        "product_type": "piles",
        "qty": 4,
        "length_m": 3.0,
        "width_m": 0.3,
    }
    order_data = [legacy_plate, explicit_pile]
    expected_plates = resolve_kp_line_weight_kg(legacy_plate)[1]

    assert total_order_cargo_weight_kg(order_data, product_types={"plates"}) == expected_plates
    assert expected_plates > 0.0


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
        (
            [
                {
                    "name": "ПБ",
                    "product_type": "plates",
                    "qty": 65,
                    "unit_price": 10.0,
                    "length_m": 1.0,
                    "width_m": 1.0,
                },
                {
                    "name": "С30.15-3",
                    "product_type": "piles",
                    "qty": 10,
                    "unit_price": 50.0,
                    "length_m": 3.0,
                    "width_m": 0.3,
                },
            ],
            0,
            100.0,
        ),
    ]
    for order_data, discount, logistics in fixtures:
        assert calculate_total_cost_pdf(order_data, discount, logistics_cost=logistics) == (
            calculate_total_cost_xlsx(order_data, discount, logistics_cost=logistics)
        )


def test_calculate_total_cost_plates_only_delivery_unchanged() -> None:
    """MNA-201: моно-плиты — доставка как сегодня (вес всех plate-строк)."""
    order_data = [
        {
            "name": "ПБ",
            "product_type": "plates",
            "qty": 65,
            "unit_price": 10.0,
            "length_m": 1.0,
            "width_m": 1.0,
        }
    ]
    trip = 100.0
    plates_kg = total_order_cargo_weight_kg(order_data, product_types={"plates"})
    expected_delivery = delivery_service_charge_rub(trip, plates_kg)
    assert cargo_delivery_trips_count(plates_kg) == 1

    totals = calculate_total_cost_xlsx(order_data, discount_percent=0, logistics_cost=trip)

    assert totals["total_with_vat"] == 650.0 + expected_delivery
    assert totals["vat_amount"] == 143.0
    assert totals["subtotal"] == round(650.0 + expected_delivery - 143.0, 2)


def test_calculate_total_cost_mixed_delivery_from_plates_kg_only() -> None:
    """MNA-201: mixed plates+piles — рейсы только от веса ПБ (сваи не увеличивают доставку)."""
    plate = {
        "name": "ПБ",
        "product_type": "plates",
        "qty": 65,
        "unit_price": 10.0,
        "length_m": 1.0,
        "width_m": 1.0,
    }
    pile = {
        "name": "С30.15-3",
        "product_type": "piles",
        "qty": 10,
        "unit_price": 50.0,
        "length_m": 3.0,
        "width_m": 0.3,
    }
    order_data = [plate, pile]

    plates_kg = total_order_cargo_weight_kg(order_data, product_types={"plates"})
    all_kg = total_order_cargo_weight_kg(order_data)
    assert plates_kg > 0.0
    assert all_kg > plates_kg
    # Sanity: без plates-фильтра сваи толкают заказ на лишний рейс.
    assert cargo_delivery_trips_count(plates_kg) == 1
    assert cargo_delivery_trips_count(all_kg) == 2

    trip = 100.0
    expected_delivery = delivery_service_charge_rub(trip, plates_kg)
    products = 65 * 10.0 + 10 * 50.0  # 1150

    totals = calculate_total_cost_xlsx(order_data, discount_percent=0, logistics_cost=trip)

    assert expected_delivery == 100.0
    assert totals["total_with_vat"] == products + expected_delivery
    assert totals["vat_amount"] == round(products * 0.22, 2)
    assert totals["subtotal"] == round(products + expected_delivery - totals["vat_amount"], 2)
    # Не суммировать вес свай в доставку.
    assert totals["total_with_vat"] != products + delivery_service_charge_rub(trip, all_kg)


def test_calculate_total_cost_piles_only_delivery_zero_despite_logistics_cost() -> None:
    """MNA-201: только сваи → доставка 0 даже при logistics_cost > 0."""
    order_data = [
        {
            "name": "С30.15-3",
            "product_type": "piles",
            "qty": 10,
            "unit_price": 50.0,
            "length_m": 3.0,
            "width_m": 0.3,
        }
    ]
    trip = 100.0
    assert total_order_cargo_weight_kg(order_data) > 0.0
    assert total_order_cargo_weight_kg(order_data, product_types={"plates"}) == 0.0

    totals = calculate_total_cost_xlsx(order_data, discount_percent=0, logistics_cost=trip)

    products = 500.0
    assert totals["total_with_vat"] == products
    assert totals["vat_amount"] == round(products * 0.22, 2)
    assert totals["subtotal"] == round(products - totals["vat_amount"], 2)


def _seed_mini_pile_catalog(tmp_path) -> str:
    db_path = fx.make_iso_db(tmp_path)
    upsert_pile_catalog(
        db_path,
        [
            PileCatalogEntry("С140.40", 14.0, 400, 2.26, 5650.0, 3),
            PileCatalogEntry("С60.30", 6.0, 300, 0.55, 1380.0, 14),
        ],
    )
    return db_path


def test_calculate_total_cost_piles_ready_uses_pile_logistics_cost(tmp_path) -> None:
    """Сваи с нормой: доставка = pile_logistics_cost × рейсы; logistics_cost плит не берём."""
    db_path = _seed_mini_pile_catalog(tmp_path)
    order_data = [
        {
            "name": "С60.30",
            "mark": "С60.30",
            "product_type": "piles",
            "qty": 14,
            "unit_price": 50.0,
        }
    ]
    totals = calculate_total_cost_core(
        order_data,
        discount_percent=0,
        logistics_cost=999.0,
        db_path=db_path,
        require_all_priced=False,
        pile_logistics_cost=100.0,
        pile_catalog_db_path=db_path,
    )
    assert totals["pile_delivery_ready"] is True
    assert totals["pile_trips"] == 1
    assert totals["pile_delivery_total"] == 100.0
    assert totals["plate_delivery_total"] == 0.0
    assert totals["total_with_vat"] == 800.0  # 14*50 + 100 pile delivery
    assert CARGO_DELIVERY_TRUCK_CAPACITY_KG == 18600


def test_calculate_total_cost_pending_pile_delivery_zero(tmp_path) -> None:
    db_path = _seed_mini_pile_catalog(tmp_path)
    order_data = [
        {
            "name": "C18-40T8",
            "mark": "C18-40T8",
            "product_type": "bridge_piles",
            "qty": 49,
            "unit_price": 10.0,
        }
    ]
    totals = calculate_total_cost_core(
        order_data,
        0,
        logistics_cost=0,
        db_path=db_path,
        require_all_priced=False,
        pile_logistics_cost=1000.0,
        pile_catalog_db_path=db_path,
    )
    assert totals["pile_delivery_ready"] is False
    assert totals["pile_delivery_total"] == 0.0
    assert totals["pile_trips"] == 0
    assert totals["total_with_vat"] == 490.0


def test_calculate_total_cost_mixed_sums_two_deliveries(tmp_path) -> None:
    db_path = _seed_mini_pile_catalog(tmp_path)
    plate = {
        "name": "ПБ",
        "product_type": "plates",
        "qty": 65,
        "unit_price": 10.0,
        "length_m": 1.0,
        "width_m": 1.0,
    }
    pile = {
        "name": "C14-40T4",
        "mark": "C14-40T4",
        "product_type": "bridge_piles",
        "qty": 3,
        "unit_price": 100.0,
    }
    order_data = [plate, pile]
    plate_kg = total_order_cargo_weight_kg(order_data, product_types={"plates"})
    plate_delivery = delivery_service_charge_rub(100.0, plate_kg)
    totals = calculate_total_cost_core(
        order_data,
        0,
        logistics_cost=100.0,
        db_path=db_path,
        require_all_priced=False,
        pile_logistics_cost=200.0,
        pile_catalog_db_path=db_path,
    )
    assert cargo_delivery_trips_count(plate_kg) == 1
    assert totals["pile_delivery_ready"] is True
    assert totals["pile_trips"] == 1  # 3 шт / pcs 3
    assert totals["plate_delivery_total"] == plate_delivery
    assert totals["pile_delivery_total"] == 200.0
    products = 65 * 10.0 + 3 * 100.0
    assert totals["total_with_vat"] == products + plate_delivery + 200.0
