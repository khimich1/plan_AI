from __future__ import annotations

import pytest

from app.services.commercial_calculation_service import (
    ERR_EMPTY_PLATES,
    ERR_NO_CLIENT,
    ERR_NO_DELIVERY,
    ERR_NO_MANAGER,
    ERR_NO_PAYMENT,
    ERR_WIDE_PLATES,
    CommercialCalculationService,
)
from core.cargo_delivery_pricing import (
    cargo_delivery_trips_count,
    delivery_service_charge_rub,
    total_order_cargo_weight_kg,
)

_SERVICE = CommercialCalculationService()

_SAMPLE_ORDER = [{"name": "n", "qty": 1, "length_m": 1, "width_m": 1, "unit_price": 1}]


def _ready_metadata(**overrides: object) -> dict:
    base = {
        "manager_id": 1,
        "client_name": "ООО Тест",
        "conditions_mode": "standard",
        "wide_plate_lines": [],
        "wide_plates_resolved": True,
    }
    base.update(overrides)
    return base


def test_validate_empty_plates() -> None:
    errors = _SERVICE.validate_calculate_prerequisites(
        order_data=[],
        metadata=_ready_metadata(),
    )
    assert errors == [ERR_EMPTY_PLATES]
    assert _SERVICE.meta_ready_for_calculate(_ready_metadata())
    assert not _SERVICE.wide_lines_blocking(_ready_metadata())


def test_validate_wide_plates_unresolved() -> None:
    metadata = _ready_metadata(
        wide_plate_lines=[{"id": "w1", "line": "X", "qty": 1}],
        wide_plates_resolved=False,
    )
    errors = _SERVICE.validate_calculate_prerequisites(
        order_data=_SAMPLE_ORDER,
        metadata=metadata,
    )
    assert errors == [ERR_WIDE_PLATES]
    assert _SERVICE.wide_lines_blocking(metadata)
    assert _SERVICE.meta_ready_for_calculate(metadata)


def test_validate_missing_manager() -> None:
    metadata = _ready_metadata(manager_id=None)
    errors = _SERVICE.validate_calculate_prerequisites(
        order_data=_SAMPLE_ORDER,
        metadata=metadata,
    )
    assert errors == [ERR_NO_MANAGER]
    assert not _SERVICE.meta_ready_for_calculate(metadata)
    assert not _SERVICE.wide_lines_blocking(metadata)


def test_validate_client_terms_custom_mode() -> None:
    metadata = _ready_metadata(
        conditions_mode="custom",
        delivery_conditions="",
        payment_conditions="",
    )
    errors = _SERVICE.validate_calculate_prerequisites(
        order_data=_SAMPLE_ORDER,
        metadata=metadata,
    )
    assert errors == [ERR_NO_DELIVERY, ERR_NO_PAYMENT]
    assert not _SERVICE.meta_ready_for_calculate(metadata)


def test_validate_ready_returns_no_errors() -> None:
    metadata = _ready_metadata()
    errors = _SERVICE.validate_calculate_prerequisites(
        order_data=_SAMPLE_ORDER,
        metadata=metadata,
    )
    assert errors == []
    assert _SERVICE.meta_ready_for_calculate(metadata)
    assert not _SERVICE.wide_lines_blocking(metadata)


def test_enforce_calculate_prerequisites_raises_first_error() -> None:
    with pytest.raises(ValueError, match=ERR_EMPTY_PLATES):
        _SERVICE.enforce_calculate_prerequisites(
            order_data=[],
            metadata=_ready_metadata(),
        )


def test_compute_totals_plates_only_delivery_unchanged() -> None:
    """MNA-201: plates-only через calculation service — доставка как сегодня."""
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

    totals = _SERVICE.compute_totals(
        order_data,
        discount_percent=0,
        logistics_cost=trip,
    )

    assert totals["total_with_vat"] == 650.0 + expected_delivery
    assert totals["vat_amount"] == 143.0


def test_compute_totals_mixed_delivery_from_plates_kg_only() -> None:
    """MNA-201: mixed — compute_totals считает рейсы только по весу plates."""
    order_data = [
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
    ]
    trip = 100.0
    plates_kg = total_order_cargo_weight_kg(order_data, product_types={"plates"})
    all_kg = total_order_cargo_weight_kg(order_data)
    assert cargo_delivery_trips_count(plates_kg) == 1
    assert cargo_delivery_trips_count(all_kg) == 2

    expected_delivery = delivery_service_charge_rub(trip, plates_kg)
    products = 1150.0

    totals = _SERVICE.compute_totals(
        order_data,
        discount_percent=0,
        logistics_cost=trip,
    )

    assert expected_delivery == 100.0
    assert totals["total_with_vat"] == products + expected_delivery
    assert totals["vat_amount"] == round(products * 0.22, 2)


def test_compute_totals_piles_only_delivery_zero_despite_logistics_cost() -> None:
    """MNA-201: piles-only — delivery 0 при logistics_cost > 0."""
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
    assert total_order_cargo_weight_kg(order_data, product_types={"plates"}) == 0.0

    totals = _SERVICE.compute_totals(
        order_data,
        discount_percent=0,
        logistics_cost=100.0,
    )

    assert totals["total_with_vat"] == 500.0
    assert totals["vat_amount"] == 110.0


# --- MNA-202: mixed discount + calculate validation (per-line / has-any-plates) ---


def _mixed_plates_piles_order() -> list[dict]:
    return [
        {
            "name": "ПБ",
            "product_type": "plates",
            "qty": 2,
            "unit_price": 1000.0,
            "length_m": 1.0,
            "width_m": 1.0,
        },
        {
            "name": "С30.15-3",
            "product_type": "piles",
            "product_kind": "pile",
            "mark": "С30.15-3",
            "concrete_grade": "B25",
            "qty": 4,
            "unit_price": 500.0,
            "length_m": 3.0,
            "width_m": 0.3,
        },
    ]


def test_order_has_plates_true_for_mixed_and_plates_only() -> None:
    """MNA-202: helper detects plates by line product_type, not metadata cycle type."""
    mixed = _mixed_plates_piles_order()
    plates_only = [mixed[0]]
    piles_only = [mixed[1]]

    assert _SERVICE.order_has_plates(mixed) is True
    assert _SERVICE.order_has_plates(plates_only) is True
    assert _SERVICE.order_has_plates(piles_only) is False
    assert _SERVICE.order_has_plates([]) is False


def test_order_has_plates_treats_missing_product_type_as_plates() -> None:
    """MNA-202: legacy lines without product_type count as plates."""
    legacy = [{"name": "ПБ", "qty": 1, "unit_price": 1.0, "length_m": 1.0, "width_m": 1.0}]
    assert _SERVICE.order_has_plates(legacy) is True


def test_validate_mixed_cycle_piles_blocks_unresolved_wide_plates() -> None:
    """MNA-202: cycle product_type=piles must not skip wide-plate gate when plates exist.

    Today ``is_pile_draft(metadata)`` short-circuits ``_wide_plate_errors`` and incorrectly
    allows calculate on mixed drafts with unresolved wide plates.
    """
    metadata = _ready_metadata(
        product_type="piles",
        wide_plate_lines=[{"id": "w1", "line": "ПБ 59-15-8п 2", "qty": 2}],
        wide_plates_resolved=False,
    )
    order_data = _mixed_plates_piles_order()
    errors = _SERVICE.validate_calculate_prerequisites(
        order_data=order_data,
        metadata=metadata,
    )
    assert errors == [ERR_WIDE_PLATES]
    # wide_lines_blocking must consider order lines, not only metadata.product_type
    assert _SERVICE.wide_lines_blocking(metadata, order_data=order_data)


def test_validate_piles_only_skips_wide_plates_despite_stale_meta_lines() -> None:
    """MNA-202: mono piles still skip wide-plate gate (no plate lines in order)."""
    metadata = _ready_metadata(
        product_type="piles",
        wide_plate_lines=[{"id": "w1", "line": "X", "qty": 1}],
        wide_plates_resolved=False,
    )
    piles_only = [
        {
            "name": "С30.15-3",
            "product_type": "piles",
            "product_kind": "pile",
            "mark": "С30.15-3",
            "qty": 1,
            "unit_price": 1.0,
            "concrete_grade": "B25",
        }
    ]
    errors = _SERVICE.validate_calculate_prerequisites(
        order_data=piles_only,
        metadata=metadata,
    )
    assert ERR_WIDE_PLATES not in errors
    assert errors == []
    assert not _SERVICE.wide_lines_blocking(metadata, order_data=piles_only)


def test_validate_mixed_cycle_piles_ready_allows_calculate() -> None:
    """MNA-202: mixed plates+piles with cycle=piles and resolved meta is calculable."""
    metadata = _ready_metadata(product_type="piles")
    order_data = _mixed_plates_piles_order()
    assert _SERVICE.is_pile_draft(metadata)
    assert _SERVICE.order_has_plates(order_data)

    errors = _SERVICE.validate_calculate_prerequisites(
        order_data=order_data,
        metadata=metadata,
    )
    assert errors == []
    assert not _SERVICE.wide_lines_blocking(metadata, order_data=order_data)


def test_compute_totals_mixed_discount_applies_to_plates_and_piles() -> None:
    """MNA-202: one discount_percent reduces product total across all line types."""
    order_data = _mixed_plates_piles_order()
    # plates 2*1000 + piles 4*500 = 4000 list; 10% → 3600; delivery 0
    list_products = 4000.0
    discount_percent = 10.0
    expected_products = list_products * (1.0 - discount_percent / 100.0)

    totals = _SERVICE.compute_totals(
        order_data,
        discount_percent=discount_percent,
        logistics_cost=0.0,
    )

    assert totals["total_with_vat"] == pytest.approx(expected_products)
    assert totals["vat_amount"] == pytest.approx(round(expected_products * 0.22, 2))

    plates_only = _SERVICE.compute_totals(
        [order_data[0]],
        discount_percent=discount_percent,
        logistics_cost=0.0,
    )
    # Mixed must include discounted pile contribution (2000 * 0.9 = 1800), not plates alone.
    assert totals["total_with_vat"] == pytest.approx(
        plates_only["total_with_vat"] + 1800.0
    )
