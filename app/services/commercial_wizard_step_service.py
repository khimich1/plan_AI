from __future__ import annotations

from typing import Any

from app.schemas.commercial import WizardNextRequiredAction, WizardStepId
from app.services.commercial_calculation_service import (
    ERR_EMPTY_PLATES,
    ERR_NO_MANAGER,
    ERR_WIDE_PLATES,
    CommercialCalculationService,
)
from app.services.draft_store import DraftStore


class CommercialWizardStepService:
    """Wizard step inference and persistence (extracted from CommercialWorkflowService)."""

    def __init__(
        self,
        *,
        calculation_service: CommercialCalculationService,
        draft_store: DraftStore,
    ) -> None:
        self.calculation_service = calculation_service
        self.draft_store = draft_store

    def wide_lines_blocking(self, metadata: dict[str, Any]) -> bool:
        return self.calculation_service.wide_lines_blocking(metadata)

    def meta_ready_for_calculate(self, metadata: dict[str, Any]) -> bool:
        return self.calculation_service.meta_ready_for_calculate(metadata)

    def normalize_stored_step(self, metadata: dict[str, Any]) -> WizardStepId:
        raw = str(metadata.get("current_step") or "").strip().lower()
        aliases = {"wide_plates": WizardStepId.wide_plates.value, "calculate": WizardStepId.client.value}
        raw = aliases.get(raw, raw)
        try:
            return WizardStepId(raw) if raw else WizardStepId.plates
        except ValueError:
            return WizardStepId.plates

    def wizard_step_after_plate_snapshot(self, metadata: dict[str, Any], order_data: list[Any]) -> WizardStepId:
        return WizardStepId.plates

    def persist_wizard_step(self, draft_id: str, step: WizardStepId) -> None:
        self.draft_store.update_metadata(draft_id, current_step=step.value)

    def infer_wizard_current_step(self, payload: dict[str, Any]) -> WizardStepId:
        """Эффективный шаг для UI: хранимый шаг + защита от «перепрыгивания» узких плит + валидность result."""
        metadata = dict(payload.get("metadata") or {})
        order_data = payload.get("order_data") or []
        stored = self.normalize_stored_step(metadata)

        if stored == WizardStepId.result:
            if (
                order_data
                and not self.wide_lines_blocking(metadata)
                and self.meta_ready_for_calculate(metadata)
            ):
                return WizardStepId.result
            return WizardStepId.client

        if order_data and self.wide_lines_blocking(metadata):
            if stored in (WizardStepId.manager, WizardStepId.client, WizardStepId.result):
                return WizardStepId.wide_plates
            if stored == WizardStepId.wide_plates:
                return WizardStepId.wide_plates

        return stored

    def infer_next_required_action(
        self,
        payload: dict[str, Any],
        effective_step: WizardStepId,
    ) -> WizardNextRequiredAction:
        metadata = dict(payload.get("metadata") or {})
        order_data = payload.get("order_data") or []
        errors = self.calculation_service.validate_calculate_prerequisites(
            order_data=order_data,
            metadata=metadata,
        )

        if errors:
            first = errors[0]
            if first == ERR_EMPTY_PLATES:
                return WizardNextRequiredAction.ingest_plates
            if first == ERR_WIDE_PLATES:
                return WizardNextRequiredAction.resolve_wide_plates
            if first == ERR_NO_MANAGER:
                return WizardNextRequiredAction.select_manager
            return WizardNextRequiredAction.complete_client_terms

        if effective_step == WizardStepId.result:
            return WizardNextRequiredAction.none

        return WizardNextRequiredAction.post_calculate

    def infer_can_proceed_to(
        self,
        payload: dict[str, Any],
        effective_step: WizardStepId,
        next_action: WizardNextRequiredAction,
    ) -> list[WizardStepId]:
        metadata = dict(payload.get("metadata") or {})
        order_data = payload.get("order_data") or []

        if effective_step == WizardStepId.plates:
            if not order_data:
                return []
            if self.wide_lines_blocking(metadata):
                return [WizardStepId.wide_plates]
            return [WizardStepId.manager]

        if effective_step == WizardStepId.wide_plates:
            return []

        if effective_step == WizardStepId.manager:
            if metadata.get("manager_id"):
                return [WizardStepId.client]
            return []

        if effective_step == WizardStepId.client:
            if next_action == WizardNextRequiredAction.post_calculate:
                return []
            return []

        return []

    def collect_wizard_validation_errors(
        self,
        payload: dict[str, Any],
        next_action: WizardNextRequiredAction,
    ) -> list[str]:
        """Сообщения, согласованные с проверками ``calculate_draft`` / ``next_required_action``."""
        if next_action == WizardNextRequiredAction.none:
            return []

        metadata = dict(payload.get("metadata") or {})
        order_data = payload.get("order_data") or []
        errors = self.calculation_service.validate_calculate_prerequisites(
            order_data=order_data,
            metadata=metadata,
        )

        if next_action == WizardNextRequiredAction.complete_client_terms:
            return errors

        action_message = {
            WizardNextRequiredAction.ingest_plates: ERR_EMPTY_PLATES,
            WizardNextRequiredAction.resolve_wide_plates: ERR_WIDE_PLATES,
            WizardNextRequiredAction.select_manager: ERR_NO_MANAGER,
        }.get(next_action)
        if action_message is None:
            return []
        return [action_message]

    def build_wizard_state(self, payload: dict[str, Any]) -> dict[str, Any]:
        current = self.infer_wizard_current_step(payload)
        next_action = self.infer_next_required_action(payload, current)
        can_proceed = self.infer_can_proceed_to(payload, current, next_action)
        validation_errors = self.collect_wizard_validation_errors(payload, next_action)
        return {
            "current_step": current,
            "can_proceed_to": can_proceed,
            "next_required_action": next_action,
            "validation_errors": validation_errors,
        }
