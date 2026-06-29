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
