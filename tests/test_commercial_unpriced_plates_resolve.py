from __future__ import annotations

from typing import Any

import pytest

from app.domain.models.optimization_context import OptimizationContext
from app.domain.models.parse_result import ParseResult
from app.domain.models.plate_order import PlateOrder
from app.schemas.commercial import WizardNextRequiredAction, WizardStepId
from app.services.commercial_calculation_service import (
    CommercialCalculationService,
    ERR_UNPRICED_PLATES,
    ERR_WIDE_PLATES,
)
from app.services.commercial_service import CommercialPreviewResult
from app.services.commercial_wizard_step_service import CommercialWizardStepService
from app.services.commercial_workflow_service import CommercialWorkflowService
from app.services.draft_store import DraftStore
from core.plate_order_context import PlateOrderContext


@pytest.fixture()
def wizard_service() -> CommercialWizardStepService:
    return CommercialWizardStepService(
        calculation_service=CommercialCalculationService(),
        draft_store=DraftStore(),
    )


def test_wizard_state_unpriced_plates_blocks_forward(
    wizard_service: CommercialWizardStepService,
) -> None:
    payload = {
        "metadata": {
            "current_step": WizardStepId.plates.value,
            "wide_plate_lines": [],
            "wide_plates_resolved": True,
            "unpriced_plate_lines": [
                {
                    "id": "unpriced-1",
                    "name": "Плиты ПБ 75-12-12п",
                    "line": "ПБ 75-12-12п 1",
                    "qty": 1,
                    "length_m": 7.5,
                    "width_m": 1.2,
                    "load_class": 1200,
                    "replacements": [{"load_code": 10, "price": 31890.0}],
                }
            ],
            "unpriced_plates_resolved": False,
        },
        "order_data": [
            {
                "name": "Плиты ПБ 75-12-12п",
                "qty": 1,
                "length_m": 7.5,
                "width_m": 1.2,
                "unit_price": None,
                "load_class": 1200,
            }
        ],
    }
    state = wizard_service.build_wizard_state(payload)
    assert state["current_step"] == WizardStepId.plates
    assert state["can_proceed_to"] == []
    assert state["next_required_action"] == WizardNextRequiredAction.resolve_unpriced_plates
    assert ERR_UNPRICED_PLATES in state["validation_errors"]


def test_wizard_state_wide_plates_have_priority_over_unpriced(
    wizard_service: CommercialWizardStepService,
) -> None:
    payload = {
        "metadata": {
            "current_step": WizardStepId.plates.value,
            "wide_plate_lines": [{"id": "w1", "line": "X", "qty": 1}],
            "wide_plates_resolved": False,
            "unpriced_plate_lines": [
                {"id": "unpriced-1", "line": "Y", "qty": 1, "replacements": []}
            ],
            "unpriced_plates_resolved": False,
        },
        "order_data": [{"name": "n", "qty": 1, "length_m": 1, "width_m": 1, "unit_price": 1}],
    }
    state = wizard_service.build_wizard_state(payload)
    assert state["next_required_action"] == WizardNextRequiredAction.resolve_wide_plates
    assert ERR_WIDE_PLATES in state["validation_errors"]


def test_wizard_state_unpriced_resolved_clears_action(
    wizard_service: CommercialWizardStepService,
) -> None:
    payload = {
        "metadata": {
            "current_step": WizardStepId.plates.value,
            "wide_plate_lines": [],
            "wide_plates_resolved": True,
            "unpriced_plate_lines": [],
            "unpriced_plates_resolved": True,
        },
        "order_data": [{"name": "n", "qty": 1, "length_m": 1, "width_m": 1, "unit_price": 1}],
    }
    state = wizard_service.build_wizard_state(payload)
    assert state["next_required_action"] == WizardNextRequiredAction.select_manager
    assert state["can_proceed_to"] == [WizardStepId.client]


def _fake_preview(text: str, order_data: list[dict[str, Any]]) -> CommercialPreviewResult:
    order = PlateOrder()
    lines = [line for line in text.split("\n") if line.strip()]
    return CommercialPreviewResult(
        parse_result=ParseResult(order=order, normalized_text=text, normalized_lines=lines),
        optimization_context=OptimizationContext(order=order),
        order_data=order_data,
        price_rows=[],
        breakdown_tables=[],
        total_sum=0.0,
        unpriced_plate_lines=[],
    )


def test_resolve_unpriced_plates_replace_load(monkeypatch: pytest.MonkeyPatch) -> None:
    workflow = CommercialWorkflowService()
    draft_payload = {
        "order": PlateOrder(),
        "optimization_context": OptimizationContext(order=PlateOrder()),
        "order_data": [],
        "metadata": {
            "source_type": "text",
            "original_text": "ПБ 75-12-12п 1",
            "ocr_text": "",
            "input_text": "ПБ 75-12-12п 1\nПБ 60-12-8п 1",
            "normalized_lines": ["ПБ 75-12-12п 1", "ПБ 60-12-8п 1"],
            "wide_plate_lines": [],
            "wide_plates_resolved": True,
            "unpriced_plate_lines": [
                {
                    "id": "unpriced-1",
                    "name": "Плиты ПБ 75-12-12п",
                    "line": "ПБ 75-12-12п 1",
                    "qty": 1,
                    "length_m": 7.5,
                    "width_m": 1.2,
                    "load_class": 1200,
                    "replacements": [
                        {"load_code": 10, "price": 31890.0},
                        {"load_code": 8, "price": 29316.0},
                    ],
                }
            ],
            "unpriced_plates_resolved": False,
            "plate_batches": [
                {
                    "source_type": "text",
                    "original_text": "",
                    "normalized_text": "ПБ 75-12-12п 1\nПБ 60-12-8п 1",
                    "ocr_text": "",
                    "filename": "",
                }
            ],
            "last_source_filename": "",
        },
    }
    monkeypatch.setattr(workflow, "_load_draft_or_raise", lambda _draft_id: draft_payload)

    captured: dict[str, Any] = {}

    def fake_generate_preview(
        *,
        text: str | None = None,
        parse_result: ParseResult | None = None,
        plate_order_ctx: Any = None,
    ) -> CommercialPreviewResult:
        captured["text"] = text or ""
        return _fake_preview(
            text or "",
            [
                {
                    "name": "Плиты ПБ 75-12-10п",
                    "qty": 1,
                    "length_m": 7.5,
                    "width_m": 1.2,
                    "load_class": 1000,
                    "unit_price": 31890.0,
                },
                {
                    "name": "Плиты ПБ 60-12-8п",
                    "qty": 1,
                    "length_m": 6.0,
                    "width_m": 1.2,
                    "load_class": 800,
                    "unit_price": 20000.0,
                },
            ],
        )

    monkeypatch.setattr(workflow.commercial_service, "generate_preview", fake_generate_preview)

    saved: dict[str, Any] = {}

    def fake_replace_preview(draft_id: str, **kwargs: Any) -> str:
        saved["order_data"] = kwargs.get("order_data")
        saved["metadata"] = kwargs.get("metadata")
        return draft_id

    monkeypatch.setattr(workflow.draft_store, "replace_preview", fake_replace_preview)
    monkeypatch.setattr(workflow, "get_draft_details", lambda draft_id: {
        "draft_id": draft_id,
        "metadata": saved["metadata"],
        "order_data": saved["order_data"],
    })
    monkeypatch.setattr(workflow, "_persist_wizard_step", lambda *_args, **_kwargs: None)

    result = workflow.resolve_unpriced_plates(
        "draft-1",
        decisions=[{"line_id": "unpriced-1", "action": "replace_load", "load_code": 10}],
        plate_order_ctx=PlateOrderContext.fresh_empty(),
    )

    assert captured["text"] == "ПБ 75-12-10п 1\nПБ 60-12-8п 1"
    assert result["metadata"]["unpriced_plates_resolved"] is True
    assert result["metadata"]["unpriced_plate_decisions"]
    assert result["metadata"]["plate_batches"][0]["normalized_text"] == "ПБ 75-12-10п 1\nПБ 60-12-8п 1"


def test_resolve_unpriced_plates_rejects_foreign_load_code(monkeypatch: pytest.MonkeyPatch) -> None:
    workflow = CommercialWorkflowService()
    draft_payload = {
        "order": PlateOrder(),
        "optimization_context": OptimizationContext(order=PlateOrder()),
        "order_data": [],
        "metadata": {
            "source_type": "text",
            "input_text": "ПБ 75-12-12п 1",
            "normalized_lines": ["ПБ 75-12-12п 1"],
            "wide_plates_resolved": True,
            "unpriced_plate_lines": [
                {
                    "id": "unpriced-1",
                    "line": "ПБ 75-12-12п 1",
                    "name": "Плиты ПБ 75-12-12п",
                    "qty": 1,
                    "length_m": 7.5,
                    "width_m": 1.2,
                    "load_class": 1200,
                    "replacements": [{"load_code": 10, "price": 31890.0}],
                }
            ],
            "unpriced_plates_resolved": False,
            "plate_batches": [],
        },
    }
    monkeypatch.setattr(workflow, "_load_draft_or_raise", lambda _draft_id: draft_payload)

    with pytest.raises(ValueError, match="не входит в предложенные"):
        workflow.resolve_unpriced_plates(
            "draft-1",
            decisions=[{"line_id": "unpriced-1", "action": "replace_load", "load_code": 6}],
            plate_order_ctx=PlateOrderContext.fresh_empty(),
        )


def test_resolve_unpriced_plates_exclude_when_no_replacements(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workflow = CommercialWorkflowService()
    draft_payload = {
        "order": PlateOrder(),
        "optimization_context": OptimizationContext(order=PlateOrder()),
        "order_data": [],
        "metadata": {
            "source_type": "text",
            "input_text": "ПБ 99-12-12п 1\nПБ 60-12-8п 1",
            "normalized_lines": ["ПБ 99-12-12п 1", "ПБ 60-12-8п 1"],
            "wide_plates_resolved": True,
            "unpriced_plate_lines": [
                {
                    "id": "unpriced-1",
                    "line": "ПБ 99-12-12п 1",
                    "name": "Плиты ПБ 99-12-12п",
                    "qty": 1,
                    "length_m": 9.9,
                    "width_m": 1.2,
                    "load_class": 1200,
                    "replacements": [],
                }
            ],
            "unpriced_plates_resolved": False,
            "plate_batches": [],
            "last_source_filename": "",
        },
    }
    monkeypatch.setattr(workflow, "_load_draft_or_raise", lambda _draft_id: draft_payload)

    with pytest.raises(ValueError, match="только исключение"):
        workflow.resolve_unpriced_plates(
            "draft-1",
            decisions=[{"line_id": "unpriced-1", "action": "replace_load", "load_code": 10}],
            plate_order_ctx=PlateOrderContext.fresh_empty(),
        )

    captured: dict[str, Any] = {}

    def fake_generate_preview(
        *,
        text: str | None = None,
        parse_result: ParseResult | None = None,
        plate_order_ctx: Any = None,
    ) -> CommercialPreviewResult:
        captured["text"] = text or ""
        return _fake_preview(
            text or "",
            [
                {
                    "name": "Плиты ПБ 60-12-8п",
                    "qty": 1,
                    "length_m": 6.0,
                    "width_m": 1.2,
                    "load_class": 800,
                    "unit_price": 20000.0,
                }
            ],
        )

    monkeypatch.setattr(workflow.commercial_service, "generate_preview", fake_generate_preview)

    saved: dict[str, Any] = {}

    def fake_replace_preview(draft_id: str, **kwargs: Any) -> str:
        saved["metadata"] = kwargs.get("metadata")
        saved["order_data"] = kwargs.get("order_data")
        return draft_id

    monkeypatch.setattr(workflow.draft_store, "replace_preview", fake_replace_preview)
    monkeypatch.setattr(
        workflow,
        "get_draft_details",
        lambda draft_id: {"draft_id": draft_id, "metadata": saved["metadata"], "order_data": saved["order_data"]},
    )
    monkeypatch.setattr(workflow, "_persist_wizard_step", lambda *_args, **_kwargs: None)

    result = workflow.resolve_unpriced_plates(
        "draft-1",
        decisions=[{"line_id": "unpriced-1", "action": "exclude"}],
        plate_order_ctx=PlateOrderContext.fresh_empty(),
    )
    assert captured["text"] == "ПБ 60-12-8п 1"
    assert result["metadata"]["unpriced_plates_resolved"] is True
