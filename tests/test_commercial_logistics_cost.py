from core.cargo_delivery_pricing import (
    CARGO_DELIVERY_TRUCK_CAPACITY_KG,
    cargo_delivery_trips_count,
    delivery_service_charge_rub,
    total_order_cargo_weight_kg,
)
from core.commercial_pricing import kp_delivery_export_lines
from core.trip_counts import TripBucketInput, dismiss_trip_notices, resolve_trip_counts
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
    assert totals["vat_amount"] == 135.24
    assert totals["subtotal"] == 614.76


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
    # Цены с НДС: плиты 122*0.5=61 + доставка 100×1 рейс = 161.
    # Без kp_id НДС по строкам × 22/122, включая доставку.
    assert totals["total_with_vat"] == 161.0
    assert totals["vat_amount"] == 29.03
    assert totals["subtotal"] == 131.97


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
    assert totals["vat_amount"] == 155.09
    assert totals["subtotal"] == 704.91


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
    assert totals["vat_amount"] == 22.0
    assert totals["subtotal"] == 100.0


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
    assert totals["vat_amount"] == 135.24
    assert totals["subtotal"] == round(650.0 + expected_delivery - totals["vat_amount"], 2)


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
    assert totals["vat_amount"] == 225.4
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
    assert totals["vat_amount"] == 90.16
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


def _bucket(present: bool, ready: bool, cargo_kg: float, formula: int) -> TripBucketInput:
    return TripBucketInput(present=present, ready=ready, cargo_kg=cargo_kg, formula=formula)


def _idle_bucket() -> TripBucketInput:
    return _bucket(False, True, 0.0, 0)


def test_resolve_plate_215220_kg_is_12_trips_without_notice() -> None:
    resolved = resolve_trip_counts(
        None,
        plate=_bucket(True, True, 215_220.0, 12),
        pile=_idle_bucket(),
        fbs_lm=_idle_bucket(),
    )
    assert cargo_delivery_trips_count(215_220.0) == 12
    assert resolved.accepted["plate"] == 12
    assert resolved.formula["plate"] == 12
    assert resolved.notice_for("plate") is None
    assert resolved.stored["plate"] == {
        "accepted": 12,
        "basis_kg": 215_220.0,
        "formula": 12,
    }
    assert resolved.stored["notices"] == []


def test_resolve_explicit_zero_is_kept_and_garbage_json_is_missing_key() -> None:
    resolved = resolve_trip_counts(
        "not-json",
        plate=_bucket(True, True, 215_220.0, 12),
        pile=_idle_bucket(),
        fbs_lm=_idle_bucket(),
        overrides={"plate": 0},
    )
    assert resolved.accepted["plate"] == 0
    assert resolved.notice_for("plate") is None
    assert resolved.stored["plate"]["accepted"] == 0
    assert resolved.stored["plate"]["formula"] == 12


def test_resolve_mass_step_replaces_manual_even_when_formula_integer_is_unchanged() -> None:
    stored = {
        "plate": {"accepted": 10, "basis_kg": 215_220.0, "formula": 12},
        "notices": [],
    }
    resolved = resolve_trip_counts(
        stored,
        plate=_bucket(True, True, 216_000.0, 12),
        pile=_idle_bucket(),
        fbs_lm=_idle_bucket(),
        overrides={"plate": 7},
    )
    assert resolved.accepted["plate"] == 12
    assert resolved.notice_for("plate") == {"previous": 10, "next": 12}
    assert resolved.stored["plate"]["basis_kg"] == 216_000.0
    assert resolved.stored["plate"]["accepted"] == 12


def test_resolve_same_mass_keeps_manual_and_epsilon_does_not_reset() -> None:
    stored = {
        "plate": {"accepted": 10, "basis_kg": 215_220.0, "formula": 12},
        "notices": [{"bucket": "plate", "previous": 4, "next": 10}],
    }
    same = resolve_trip_counts(
        stored,
        plate=_bucket(True, True, 215_220.0, 12),
        pile=_idle_bucket(),
        fbs_lm=_idle_bucket(),
    )
    assert same.accepted["plate"] == 10
    assert same.notice_for("plate") == {"previous": 4, "next": 10}

    tiny = resolve_trip_counts(
        stored,
        plate=_bucket(True, True, 215_220.01, 12),
        pile=_idle_bucket(),
        fbs_lm=_idle_bucket(),
    )
    assert tiny.accepted["plate"] == 10
    assert tiny.stored["plate"]["basis_kg"] == 215_220.0


def test_resolve_formula_growth_replaces_plate_notice_only() -> None:
    stored = {
        "plate": {"accepted": 12, "basis_kg": 215_220.0, "formula": 12},
        "fbs_lm": {"accepted": 9, "basis_kg": 15_000.0, "formula": 1},
        "notices": [{"bucket": "pile", "previous": 1, "next": 2}],
    }
    resolved = resolve_trip_counts(
        stored,
        plate=_bucket(True, True, 215_220.0 + 18_600.0, 13),
        pile=_idle_bucket(),
        fbs_lm=_bucket(True, True, 15_000.0, 1),
        overrides={"fbs_lm": 4},
    )
    assert resolved.accepted["plate"] == 13
    assert resolved.notice_for("plate") == {"previous": 12, "next": 13}
    assert resolved.accepted["fbs_lm"] == 1
    assert resolved.notice_for("fbs_lm") is None
    assert resolved.stored["fbs_lm"]["accepted"] == 1
    # Чужой котёл не в заказе — его старое напоминание не трогаем.
    assert any(row["bucket"] == "pile" for row in resolved.stored["notices"])


def test_resolve_pile_accepted_is_ignored_and_key_stays() -> None:
    stored = {
        "pile": {"accepted": 125, "basis_kg": 40_000.0, "formula": 2},
        "notices": [{"bucket": "pile", "previous": 125, "next": 2}],
    }
    resolved = resolve_trip_counts(
        stored,
        plate=_idle_bucket(),
        pile=_bucket(True, True, 40_000.0, 2),
        fbs_lm=_idle_bucket(),
        overrides={"pile": 9},
    )
    assert resolved.accepted["pile"] == 2
    assert resolved.formula["pile"] == 2
    assert resolved.notice_for("pile") is None
    assert "pile" in resolved.stored
    assert resolved.stored["pile"]["accepted"] == 2
    assert resolved.stored["notices"] == []


def test_resolve_archive_flag_keeps_plate_57_when_formula_would_be_49() -> None:
    stored = {
        "plate": {"accepted": 57, "basis_kg": 1_060_200.0, "formula": 57},
        "notices": [],
    }
    kept = resolve_trip_counts(
        stored,
        plate=_bucket(True, True, 899_947.5, 49),
        pile=_idle_bucket(),
        fbs_lm=_idle_bucket(),
        reset_plate_on_mass_change=False,
    )
    assert kept.accepted["plate"] == 57
    assert kept.formula["plate"] == 49
    assert kept.notice_for("plate") is None
    assert kept.stored["plate"]["accepted"] == 57
    assert kept.stored["plate"]["basis_kg"] == 1_060_200.0

    reset = resolve_trip_counts(
        stored,
        plate=_bucket(True, True, 899_947.5, 49),
        pile=_idle_bucket(),
        fbs_lm=_idle_bucket(),
    )
    assert reset.accepted["plate"] == 49
    assert reset.notice_for("plate") == {"previous": 57, "next": 49}
    assert reset.stored["plate"]["accepted"] == 49
    assert reset.stored["plate"]["basis_kg"] == 899_947.5


def test_resolve_plate_mass_reset_does_not_notice_fbs() -> None:
    stored = {
        "plate": {"accepted": 10, "basis_kg": 100.0, "formula": 12},
        "fbs_lm": {"accepted": 2, "basis_kg": 100.0, "formula": 1},
        "notices": [],
    }
    resolved = resolve_trip_counts(
        stored,
        plate=_bucket(True, True, 215_220.0, 12),
        pile=_idle_bucket(),
        fbs_lm=_bucket(True, True, 15_000.0, 1),
    )
    assert resolved.notice_for("plate") == {"previous": 10, "next": 12}
    assert resolved.accepted["fbs_lm"] == 1
    assert resolved.notice_for("fbs_lm") is None
    assert resolved.stored["fbs_lm"]["accepted"] == 1
    dismissed = dismiss_trip_notices(resolved.stored, ["plate"])
    assert dismissed["plate"]["accepted"] == 12
    assert dismissed["fbs_lm"]["accepted"] == 1
    assert dismissed["notices"] == []


def test_resolve_not_ready_drops_key_so_next_ready_is_first_fill() -> None:
    stored = {
        "pile": {"accepted": 4, "basis_kg": 40_000.0, "formula": 2},
        "notices": [{"bucket": "pile", "previous": 4, "next": 2}],
    }
    pending = resolve_trip_counts(
        stored,
        plate=_idle_bucket(),
        pile=_bucket(True, False, 40_000.0, 2),
        fbs_lm=_idle_bucket(),
    )
    assert pending.accepted["pile"] == 0
    assert "pile" not in pending.stored
    assert pending.stored["notices"] == []

    ready = resolve_trip_counts(
        pending.stored,
        plate=_idle_bucket(),
        pile=_bucket(True, True, 50_000.0, 3),
        fbs_lm=_idle_bucket(),
    )
    assert ready.accepted["pile"] == 3
    assert ready.notice_for("pile") is None


def _plate_order() -> list[dict]:
    return [
        {
            "name": "ПБ",
            "product_type": "plates",
            "qty": 1,
            "unit_price": 100.0,
            "length_m": 6.0,
            "width_m": 1.2,
        }
    ]


def test_calculate_total_cost_manual_plate_trips_and_tariff_does_not_reset(tmp_path) -> None:
    db_path = fx.make_iso_db(tmp_path)
    order = _plate_order()
    kg = total_order_cargo_weight_kg(order, product_types={"plates"})
    formula = cargo_delivery_trips_count(kg)
    stored = {
        "plate": {"accepted": 10, "basis_kg": kg, "formula": formula},
        "notices": [],
    }
    manual = calculate_total_cost_core(
        order,
        0,
        logistics_cost=2_000.0,
        db_path=db_path,
        require_all_priced=False,
        trip_counts=stored,
    )
    assert manual["plate_trips"] == 10
    assert manual["plate_trips_formula"] == formula
    assert manual["plate_trips_notice"] is None
    assert manual["plate_delivery_total"] == 20_000.0
    lines = kp_delivery_export_lines(
        manual, plate_trip_cost=2_000.0, pile_trip_cost=0.0
    )
    assert lines[0]["trips"] == 10
    assert lines[0]["amount"] == 20_000.0

    retariff = calculate_total_cost_core(
        order,
        0,
        logistics_cost=3_000.0,
        db_path=db_path,
        require_all_priced=False,
        trip_counts=manual["trip_counts"],
    )
    assert retariff["plate_trips"] == 10
    assert retariff["plate_delivery_total"] == 30_000.0
    assert retariff["plate_trips_notice"] is None

    zero = calculate_total_cost_core(
        order,
        0,
        logistics_cost=2_000.0,
        db_path=db_path,
        require_all_priced=False,
        trip_counts={
            "plate": {"accepted": 0, "basis_kg": kg, "formula": formula},
            "notices": [],
        },
    )
    assert zero["plate_delivery_total"] == 0.0
    assert kp_delivery_export_lines(zero, plate_trip_cost=2_000.0, pile_trip_cost=0.0) == []


def test_calculate_total_cost_without_trip_counts_matches_formula(tmp_path) -> None:
    db_path = fx.make_iso_db(tmp_path)
    order = _plate_order()
    kg = total_order_cargo_weight_kg(order, product_types={"plates"})
    formula = cargo_delivery_trips_count(kg)
    bare = calculate_total_cost_core(
        order,
        0,
        logistics_cost=2_000.0,
        db_path=db_path,
        require_all_priced=False,
    )
    charged = delivery_service_charge_rub(2_000.0, kg)
    assert bare["plate_trips"] == formula
    assert bare["plate_trips_formula"] == formula
    assert bare["plate_trips_notice"] is None
    assert bare["plate_delivery_total"] == charged
    assert bare["trip_counts"]["notices"] == []
    assert bare["trip_counts"]["plate"]["accepted"] == formula


def test_calculate_total_cost_pending_pile_ignores_accepted_count(tmp_path) -> None:
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
        pile_logistics_cost=1_000.0,
        pile_catalog_db_path=db_path,
        trip_counts={
            "pile": {"accepted": 5, "basis_kg": 1.0, "formula": 1},
            "notices": [{"bucket": "pile", "previous": 5, "next": 1}],
        },
    )
    assert totals["pile_delivery_ready"] is False
    assert totals["pile_trips"] == 0
    assert totals["pile_delivery_total"] == 0.0
    assert totals["total_with_vat"] == 490.0
    assert "pile" not in totals["trip_counts"]
    assert totals["trip_counts"]["notices"] == []


def test_calculate_total_cost_short_pile_ignores_stored_accepted(tmp_path) -> None:
    """Сваи: trip_counts.pile.accepted не подменяет формулу и не ставит плашку."""
    db_path = _seed_mini_pile_catalog(tmp_path)
    order_data = [
        {
            "name": "С60.30",
            "mark": "С60.30",
            "product_type": "piles",
            "qty": 28,
            "unit_price": 50.0,
        }
    ]
    probed = calculate_total_cost_core(
        order_data,
        0,
        logistics_cost=999.0,
        db_path=db_path,
        require_all_priced=False,
        pile_logistics_cost=1_000.0,
        pile_catalog_db_path=db_path,
    )
    basis = probed["trip_counts"]["pile"]["basis_kg"]
    formula = probed["pile_trips_formula"]
    assert formula == 2
    manual = calculate_total_cost_core(
        order_data,
        0,
        logistics_cost=999.0,
        db_path=db_path,
        require_all_priced=False,
        pile_logistics_cost=1_000.0,
        pile_catalog_db_path=db_path,
        trip_counts={
            "pile": {"accepted": 125, "basis_kg": basis, "formula": formula},
            "notices": [],
        },
        accepted_trip_overrides={"pile": 9},
    )
    assert manual["pile_trips"] == 2
    assert manual["pile_delivery_total"] == 2_000.0
    assert manual["pile_trips_notice"] is None
    assert manual["total_with_vat"] == 28 * 50.0 + 2_000.0
    assert manual["trip_counts"]["pile"]["accepted"] == 2

    shifted = calculate_total_cost_core(
        order_data,
        0,
        logistics_cost=999.0,
        db_path=db_path,
        require_all_priced=False,
        pile_logistics_cost=2_000.0,
        pile_catalog_db_path=db_path,
        trip_counts={
            "pile": {"accepted": 125, "basis_kg": basis + 1.0, "formula": formula},
            "notices": [],
        },
    )
    assert shifted["pile_trips"] == formula
    assert shifted["pile_trips_notice"] is None
    assert shifted["pile_delivery_total"] == round(2_000.0 * formula, 2)
    assert "pile" in shifted["trip_counts"]


def _plate_order_formula_49() -> list[dict]:
    """12 м × 1,2 м × 219 шт → ceil(кг / 18600) = 49."""
    return [
        {
            "name": "ПБ",
            "product_type": "plates",
            "qty": 219,
            "unit_price": 100.0,
            "length_m": 12.0,
            "width_m": 1.2,
        }
    ]


def test_calculate_total_cost_same_mass_keeps_plate_57(tmp_path) -> None:
    db_path = fx.make_iso_db(tmp_path)
    order = _plate_order_formula_49()
    kg = total_order_cargo_weight_kg(order, product_types={"plates"})
    formula = cargo_delivery_trips_count(kg)
    assert formula == 49
    kept = calculate_total_cost_core(
        order,
        0,
        logistics_cost=2_000.0,
        db_path=db_path,
        require_all_priced=False,
        trip_counts={
            "plate": {"accepted": 57, "basis_kg": kg, "formula": formula},
            "notices": [],
        },
    )
    assert kept["plate_trips"] == 57
    assert kept["plate_trips_formula"] == 49
    assert kept["plate_trips_notice"] is None
    assert kept["plate_delivery_total"] == 114_000.0


def test_calculate_total_cost_archive_flag_keeps_plate_57_when_mass_changes(tmp_path) -> None:
    db_path = fx.make_iso_db(tmp_path)
    order = _plate_order_formula_49()
    kg = total_order_cargo_weight_kg(order, product_types={"plates"})
    assert cargo_delivery_trips_count(kg) == 49
    stored = {
        "plate": {"accepted": 57, "basis_kg": 1_060_200.0, "formula": 57},
        "pile": {"accepted": 125, "basis_kg": 1.0, "formula": 2},
        "notices": [],
    }
    kept = calculate_total_cost_core(
        order,
        0,
        logistics_cost=1_000.0,
        db_path=db_path,
        require_all_priced=False,
        trip_counts=stored,
        reset_plate_on_mass_change=False,
    )
    assert kept["plate_trips"] == 57
    assert kept["plate_trips_formula"] == 49
    assert kept["plate_trips_notice"] is None
    assert kept["trip_counts"]["plate"]["accepted"] == 57
    assert kept["trip_counts"]["pile"]["accepted"] == 125

    reset = calculate_total_cost_core(
        order,
        0,
        logistics_cost=1_000.0,
        db_path=db_path,
        require_all_priced=False,
        trip_counts=stored,
    )
    assert reset["plate_trips"] == 49
    assert reset["plate_trips_notice"] == {"previous": 57, "next": 49}
    assert reset["plate_delivery_total"] == 49_000.0
