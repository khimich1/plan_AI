from __future__ import annotations

from app.schemas.commercial import WizardNextRequiredAction, WizardStepId
from app.services.commercial_calculation_service import (
    ERR_INVALID_WIDTHS,
    ERR_UNPRICED_PLATES,
    ERR_WIDE_PLATES,
    CommercialCalculationService,
)
from app.services.commercial_wizard_step_service import CommercialWizardStepService
from app.services.draft_store import DraftStore


def _wizard() -> CommercialWizardStepService:
    return CommercialWizardStepService(
        calculation_service=CommercialCalculationService(),
        draft_store=DraftStore(),
    )


def _invalid_line() -> dict:
    return {
        "id": "invalid-width-1",
        "name": "Плиты ПБ 29-8-8п",
        "line": "ПБ 29-8-8п 1",
        "qty": 1,
        "width_mm": 800,
        "replacements": [
            {"width_mm": 720, "width_label": "7,2"},
            {"width_mm": 860, "width_label": "8,6"},
        ],
    }


def test_wizard_blocks_on_unresolved_invalid_widths() -> None:
    payload = {
        "metadata": {
            "current_step": WizardStepId.plates.value,
            "wide_plate_lines": [],
            "wide_plates_resolved": True,
            "invalid_width_lines": [_invalid_line()],
            "invalid_widths_resolved": False,
            "unpriced_plate_lines": [],
            "unpriced_plates_resolved": True,
        },
        "order_data": [
            {"name": "Плиты ПБ 29-8-8п", "qty": 1, "length_m": 2.9, "width_m": 0.8, "unit_price": 1}
        ],
    }
    state = _wizard().build_wizard_state(payload)
    assert state["next_required_action"] == WizardNextRequiredAction.resolve_invalid_widths
    assert state["can_proceed_to"] == []
    assert ERR_INVALID_WIDTHS in state["validation_errors"]


def test_wizard_wide_has_priority_over_invalid_width() -> None:
    payload = {
        "metadata": {
            "current_step": WizardStepId.plates.value,
            "wide_plate_lines": [{"id": "w1", "line": "X", "qty": 1}],
            "wide_plates_resolved": False,
            "invalid_width_lines": [_invalid_line()],
            "invalid_widths_resolved": False,
            "unpriced_plate_lines": [],
            "unpriced_plates_resolved": True,
        },
        "order_data": [{"name": "n", "qty": 1, "length_m": 1, "width_m": 1, "unit_price": 1}],
    }
    state = _wizard().build_wizard_state(payload)
    assert state["next_required_action"] == WizardNextRequiredAction.resolve_wide_plates
    assert ERR_WIDE_PLATES in state["validation_errors"]


def test_wizard_invalid_width_has_priority_over_unpriced() -> None:
    payload = {
        "metadata": {
            "current_step": WizardStepId.plates.value,
            "wide_plate_lines": [],
            "wide_plates_resolved": True,
            "invalid_width_lines": [_invalid_line()],
            "invalid_widths_resolved": False,
            "unpriced_plate_lines": [{"id": "unpriced-1", "line": "Y", "qty": 1, "replacements": []}],
            "unpriced_plates_resolved": False,
        },
        "order_data": [{"name": "n", "qty": 1, "length_m": 1, "width_m": 1, "unit_price": None}],
    }
    state = _wizard().build_wizard_state(payload)
    assert state["next_required_action"] == WizardNextRequiredAction.resolve_invalid_widths
    assert ERR_INVALID_WIDTHS in state["validation_errors"]
    assert ERR_UNPRICED_PLATES not in state["validation_errors"]


def test_wizard_unpriced_only_unchanged() -> None:
    payload = {
        "metadata": {
            "current_step": WizardStepId.plates.value,
            "wide_plate_lines": [],
            "wide_plates_resolved": True,
            "invalid_width_lines": [],
            "invalid_widths_resolved": True,
            "unpriced_plate_lines": [{"id": "unpriced-1", "line": "Y", "qty": 1, "replacements": []}],
            "unpriced_plates_resolved": False,
        },
        "order_data": [{"name": "n", "qty": 1, "length_m": 1, "width_m": 1, "unit_price": None}],
    }
    state = _wizard().build_wizard_state(payload)
    assert state["next_required_action"] == WizardNextRequiredAction.resolve_unpriced_plates


def test_calculate_blocked_while_invalid_widths_open() -> None:
    metadata = {
        "manager_id": 1,
        "client_name": "ООО Тест",
        "conditions_mode": "standard",
        "wide_plate_lines": [],
        "wide_plates_resolved": True,
        "invalid_width_lines": [_invalid_line()],
        "invalid_widths_resolved": False,
        "unpriced_plate_lines": [],
        "unpriced_plates_resolved": True,
    }
    service = CommercialCalculationService()
    errors = service.validate_calculate_prerequisites(
        order_data=[{"name": "n", "qty": 1, "length_m": 1, "width_m": 0.8, "unit_price": 1}],
        metadata=metadata,
    )
    assert errors[0] == ERR_INVALID_WIDTHS
    assert service.invalid_width_lines_blocking(metadata)
