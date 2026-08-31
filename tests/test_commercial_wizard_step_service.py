from __future__ import annotations

import pytest

from app.schemas.commercial import WizardNextRequiredAction, WizardStepId, _coerce_wizard_step_id
from app.services.commercial_calculation_service import CommercialCalculationService
from app.services.commercial_wizard_step_service import CommercialWizardStepService
from app.services.draft_store import DraftStore


@pytest.fixture
def wizard_service() -> CommercialWizardStepService:
    return CommercialWizardStepService(
        calculation_service=CommercialCalculationService(),
        draft_store=DraftStore(),
    )


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("plates", WizardStepId.plates),
        ("client", WizardStepId.client),
        ("result", WizardStepId.result),
        ("wide-plates", WizardStepId.plates),
        ("wide_plates", WizardStepId.plates),
        ("manager", WizardStepId.client),
        ("calculate", WizardStepId.client),
        ("", WizardStepId.plates),
        ("unknown-step", WizardStepId.plates),
    ],
)
def test_coerce_wizard_step_id_legacy_aliases(raw: str, expected: WizardStepId) -> None:
    assert _coerce_wizard_step_id(raw) == expected


def test_wizard_state_ingest_plates_validation_errors(wizard_service: CommercialWizardStepService) -> None:
    payload = {
        "metadata": {"current_step": WizardStepId.plates.value, "wide_plate_lines": [], "wide_plates_resolved": True},
        "order_data": [],
    }
    state = wizard_service.build_wizard_state(payload)
    assert state["next_required_action"] == WizardNextRequiredAction.ingest_plates
    assert state["validation_errors"] == ["Список плит пустой."]
    assert state["can_proceed_to"] == []


def test_wizard_state_wide_plates_blocks_forward_on_plates(wizard_service: CommercialWizardStepService) -> None:
    payload = {
        "metadata": {
            "current_step": WizardStepId.plates.value,
            "wide_plate_lines": [{"id": "w1", "line": "X", "qty": 1}],
            "wide_plates_resolved": False,
        },
        "order_data": [{"name": "n", "qty": 1, "length_m": 1, "width_m": 1, "unit_price": 1}],
    }
    state = wizard_service.build_wizard_state(payload)
    assert state["current_step"] == WizardStepId.plates
    assert state["can_proceed_to"] == []
    assert state["next_required_action"] == WizardNextRequiredAction.resolve_wide_plates
    assert state["validation_errors"] == ["Сначала примите решение по позициям шире стандартной."]


def test_wizard_state_legacy_wide_plates_step_maps_to_plates(wizard_service: CommercialWizardStepService) -> None:
    payload = {
        "metadata": {
            "current_step": "wide-plates",
            "wide_plate_lines": [{"id": "w1", "line": "X", "qty": 1}],
            "wide_plates_resolved": False,
        },
        "order_data": [{"name": "n", "qty": 1, "length_m": 1, "width_m": 1, "unit_price": 1}],
    }
    state = wizard_service.build_wizard_state(payload)
    assert state["current_step"] == WizardStepId.plates
    assert state["can_proceed_to"] == []
    assert state["next_required_action"] == WizardNextRequiredAction.resolve_wide_plates


def test_wizard_state_ingest_piles_validation_errors(wizard_service: CommercialWizardStepService) -> None:
    payload = {
        "metadata": {
            "current_step": WizardStepId.piles.value,
            "product_type": "piles",
            "wide_plates_resolved": True,
        },
        "order_data": [],
    }
    state = wizard_service.build_wizard_state(payload)
    assert state["next_required_action"] == WizardNextRequiredAction.ingest_piles
    assert state["validation_errors"] == ["Список свай пустой."]
    assert state["can_proceed_to"] == []


def test_wizard_state_piles_no_wide_plates_gate(wizard_service: CommercialWizardStepService) -> None:
    payload = {
        "metadata": {
            "current_step": WizardStepId.piles.value,
            "product_type": "piles",
            "wide_plate_lines": [{"id": "w1", "line": "X", "qty": 1}],
            "wide_plates_resolved": False,
        },
        "order_data": [
            {
                "product_type": "piles",
                "product_kind": "pile",
                "mark": "С120.35-12",
                "qty": 1,
                "unit_price": 1.0,
                "concrete_grade": "B25",
            }
        ],
    }
    state = wizard_service.build_wizard_state(payload)
    assert state["current_step"] == WizardStepId.piles
    assert state["can_proceed_to"] == [WizardStepId.client]
    assert state["next_required_action"] == WizardNextRequiredAction.select_manager


def test_wizard_state_piles_unpriced_blocks_proceed(wizard_service: CommercialWizardStepService) -> None:
    payload = {
        "metadata": {
            "current_step": WizardStepId.piles.value,
            "product_type": "piles",
            "wide_plates_resolved": True,
        },
        "order_data": [
            {
                "product_kind": "pile",
                "mark": "С120.35-99",
                "name": "С120.35-99",
                "qty": 2,
                "unit_price": None,
                "concrete_grade": "B25",
            }
        ],
    }
    state = wizard_service.build_wizard_state(payload)
    assert state["can_proceed_to"] == []
    assert any("С120.35-99" in err for err in state["validation_errors"])


def test_wizard_state_plates_can_proceed_to_client_when_ready(wizard_service: CommercialWizardStepService) -> None:
    payload = {
        "metadata": {
            "current_step": WizardStepId.plates.value,
            "wide_plate_lines": [],
            "wide_plates_resolved": True,
        },
        "order_data": [{"name": "n", "qty": 1, "length_m": 1, "width_m": 1, "unit_price": 1}],
    }
    state = wizard_service.build_wizard_state(payload)
    assert state["current_step"] == WizardStepId.plates
    assert state["can_proceed_to"] == [WizardStepId.client]


def test_wizard_state_client_requires_calculate(wizard_service: CommercialWizardStepService) -> None:
    payload = {
        "metadata": {
            "current_step": WizardStepId.client.value,
            "manager_id": 1,
            "client_name": "ООО А",
            "conditions_mode": "standard",
            "wide_plate_lines": [],
            "wide_plates_resolved": True,
        },
        "order_data": [{"name": "n", "qty": 1, "length_m": 1, "width_m": 1, "unit_price": 1}],
    }
    state = wizard_service.build_wizard_state(payload)
    assert state["current_step"] == WizardStepId.client
    assert state["can_proceed_to"] == []
    assert state["next_required_action"] == WizardNextRequiredAction.post_calculate
    assert state["validation_errors"] == []


def test_wizard_state_result_after_calculate(wizard_service: CommercialWizardStepService) -> None:
    payload = {
        "metadata": {
            "current_step": WizardStepId.result.value,
            "manager_id": 1,
            "client_name": "ООО А",
            "conditions_mode": "standard",
            "wide_plate_lines": [],
            "wide_plates_resolved": True,
        },
        "order_data": [{"name": "n", "qty": 1, "length_m": 1, "width_m": 1, "unit_price": 1}],
    }
    state = wizard_service.build_wizard_state(payload)
    assert state["current_step"] == WizardStepId.result
    assert state["next_required_action"] == WizardNextRequiredAction.none
    assert state["validation_errors"] == []


def test_wizard_state_select_manager_on_client_step(wizard_service: CommercialWizardStepService) -> None:
    payload = {
        "metadata": {
            "current_step": "manager",
            "manager_id": None,
            "client_name": "",
            "conditions_mode": "standard",
            "wide_plate_lines": [],
            "wide_plates_resolved": True,
        },
        "order_data": [{"name": "n", "qty": 1, "length_m": 1, "width_m": 1, "unit_price": 1}],
    }
    state = wizard_service.build_wizard_state(payload)
    assert state["current_step"] == WizardStepId.client
    assert state["next_required_action"] == WizardNextRequiredAction.select_manager
    assert state["validation_errors"] == ["Выберите менеджера."]


def test_wizard_state_product_step_hides_manager_validation_error(
    wizard_service: CommercialWizardStepService,
) -> None:
    """На шаге ввода ступеней next_action может быть select_manager, но UI не должен показывать это."""
    payload = {
        "metadata": {
            "current_step": "steps",
            "product_type": "steps",
            "manager_id": None,
            "client_name": "",
            "conditions_mode": "standard",
            "wide_plate_lines": [],
            "wide_plates_resolved": True,
        },
        "order_data": [
            {
                "product_kind": "step",
                "name": "ЛС11",
                "mark": "ЛС11",
                "qty": 1,
                "unit_price": 1409.91,
            }
        ],
    }
    state = wizard_service.build_wizard_state(payload)
    assert state["current_step"] == WizardStepId.steps
    assert state["next_required_action"] == WizardNextRequiredAction.select_manager
    assert "Выберите менеджера." not in state["validation_errors"]
    assert state["can_proceed_to"] == [WizardStepId.client]


# --- MNA-104: skip client on cycle ≥2 / resume ---


def _priced_plate_line() -> dict:
    return {"name": "n", "qty": 1, "length_m": 1, "width_m": 1, "unit_price": 1}


def _mono_first_cycle_meta(**overrides: object) -> dict:
    meta = {
        "current_step": WizardStepId.plates.value,
        "product_type": "plates",
        "client_name": "",
        "append_batches": [],
        "resume_kp_id": None,
        "wide_plate_lines": [],
        "wide_plates_resolved": True,
    }
    meta.update(overrides)
    return meta


@pytest.mark.parametrize(
    ("metadata", "expected"),
    [
        ({"client_name": "", "append_batches": [], "resume_kp_id": None}, False),
        ({"client_name": "   ", "append_batches": [], "resume_kp_id": None}, False),
        ({"client_name": "ООО А", "append_batches": [], "resume_kp_id": None}, True),
        (
            {
                "client_name": "",
                "append_batches": [
                    {"batch_id": "b1", "product_type": "plates", "line_ids": ["ln1"]}
                ],
                "resume_kp_id": None,
            },
            True,
        ),
        ({"client_name": "", "append_batches": [], "resume_kp_id": 42}, True),
        (
            {
                "client_name": "ООО А",
                "append_batches": [
                    {"batch_id": "b1", "product_type": "piles", "line_ids": ["ln1"]}
                ],
                "resume_kp_id": 7,
            },
            True,
        ),
    ],
)
def test_should_skip_client_step_mna104(
    wizard_service: CommercialWizardStepService,
    metadata: dict,
    expected: bool,
) -> None:
    assert wizard_service.should_skip_client_step(metadata) is expected


def test_wizard_step_order_mono_first_cycle_includes_client(
    wizard_service: CommercialWizardStepService,
) -> None:
    """Mono first cycle: product → client → result (unchanged)."""
    order = wizard_service.wizard_step_order(_mono_first_cycle_meta())
    assert order == [WizardStepId.plates, WizardStepId.client, WizardStepId.result]


@pytest.mark.parametrize(
    ("product_type", "product_step"),
    [
        ("plates", WizardStepId.plates),
        ("piles", WizardStepId.piles),
        ("steps", WizardStepId.steps),
        ("marches", WizardStepId.marches),
        ("bridge_piles", WizardStepId.bridge_piles),
        ("fbs", WizardStepId.fbs),
    ],
)
def test_wizard_step_order_skips_client_when_client_name_set(
    wizard_service: CommercialWizardStepService,
    product_type: str,
    product_step: WizardStepId,
) -> None:
    order = wizard_service.wizard_step_order(
        _mono_first_cycle_meta(
            product_type=product_type,
            current_step=product_step.value,
            client_name="ООО Клиент",
        )
    )
    assert order == [product_step, WizardStepId.result]
    assert WizardStepId.client not in order


def test_wizard_step_order_skips_client_when_append_batches_nonempty(
    wizard_service: CommercialWizardStepService,
) -> None:
    order = wizard_service.wizard_step_order(
        _mono_first_cycle_meta(
            append_batches=[
                {"batch_id": "batch-1", "product_type": "plates", "line_ids": ["ln1"]}
            ]
        )
    )
    assert order == [WizardStepId.plates, WizardStepId.result]


def test_wizard_step_order_skips_client_when_resume_kp_id_set(
    wizard_service: CommercialWizardStepService,
) -> None:
    order = wizard_service.wizard_step_order(
        _mono_first_cycle_meta(resume_kp_id=99)
    )
    assert order == [WizardStepId.plates, WizardStepId.result]


def test_can_proceed_to_result_when_skip_client_client_name(
    wizard_service: CommercialWizardStepService,
) -> None:
    """Cycle ≥2: from product step proceed to result, not client."""
    payload = {
        "metadata": _mono_first_cycle_meta(client_name="ООО А"),
        "order_data": [_priced_plate_line()],
    }
    state = wizard_service.build_wizard_state(payload)
    assert state["current_step"] == WizardStepId.plates
    assert state["can_proceed_to"] == [WizardStepId.result]


def test_can_proceed_to_result_when_skip_client_append_batches(
    wizard_service: CommercialWizardStepService,
) -> None:
    payload = {
        "metadata": _mono_first_cycle_meta(
            append_batches=[
                {"batch_id": "b1", "product_type": "plates", "line_ids": ["ln1"]}
            ]
        ),
        "order_data": [_priced_plate_line()],
    }
    state = wizard_service.build_wizard_state(payload)
    assert state["can_proceed_to"] == [WizardStepId.result]


def test_can_proceed_to_result_when_skip_client_resume_kp_id(
    wizard_service: CommercialWizardStepService,
) -> None:
    payload = {
        "metadata": _mono_first_cycle_meta(resume_kp_id=15),
        "order_data": [_priced_plate_line()],
    }
    state = wizard_service.build_wizard_state(payload)
    assert state["can_proceed_to"] == [WizardStepId.result]


def test_can_proceed_to_still_client_on_mono_first_cycle(
    wizard_service: CommercialWizardStepService,
) -> None:
    """Regression: first cycle without client/append/resume still goes to client."""
    payload = {
        "metadata": _mono_first_cycle_meta(),
        "order_data": [_priced_plate_line()],
    }
    state = wizard_service.build_wizard_state(payload)
    assert wizard_service.should_skip_client_step(payload["metadata"]) is False
    assert state["can_proceed_to"] == [WizardStepId.client]
