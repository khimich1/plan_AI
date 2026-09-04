from __future__ import annotations

from typing import Any

from app.schemas.commercial import WizardNextRequiredAction, WizardStepId, _coerce_wizard_step_id
from app.services.commercial_calculation_service import (
    ERR_EMPTY_PILES,
    ERR_EMPTY_BRIDGE_PILES,
    ERR_EMPTY_FBS,
    ERR_EMPTY_MARCHES,
    ERR_EMPTY_PLATES,
    ERR_EMPTY_STEPS,
    ERR_NO_CLIENT,
    ERR_NO_DELIVERY,
    ERR_NO_MANAGER,
    ERR_NO_PAYMENT,
    ERR_INVALID_WIDTHS,
    ERR_UNPRICED_PLATES,
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

    def unpriced_lines_blocking(self, metadata: dict[str, Any]) -> bool:
        return self.calculation_service.unpriced_lines_blocking(metadata)

    def invalid_width_lines_blocking(self, metadata: dict[str, Any]) -> bool:
        return self.calculation_service.invalid_width_lines_blocking(metadata)

    def meta_ready_for_calculate(self, metadata: dict[str, Any]) -> bool:
        return self.calculation_service.meta_ready_for_calculate(metadata)

    def normalize_stored_step(self, metadata: dict[str, Any]) -> WizardStepId:
        return _coerce_wizard_step_id(metadata.get("current_step"))

    def is_pile_draft(self, metadata: dict[str, Any]) -> bool:
        return self.calculation_service.is_pile_draft(metadata)

    def is_step_draft(self, metadata: dict[str, Any]) -> bool:
        return self.calculation_service.is_step_draft(metadata)

    def is_march_draft(self, metadata: dict[str, Any]) -> bool:
        return self.calculation_service.is_march_draft(metadata)

    def is_bridge_pile_draft(self, metadata: dict[str, Any]) -> bool:
        return self.calculation_service.is_bridge_pile_draft(metadata)

    def is_fbs_draft(self, metadata: dict[str, Any]) -> bool:
        return self.calculation_service.is_fbs_draft(metadata)

    def product_step(self, metadata: dict[str, Any]) -> WizardStepId:
        if self.is_step_draft(metadata):
            return WizardStepId.steps
        if self.is_march_draft(metadata):
            return WizardStepId.marches
        if self.is_bridge_pile_draft(metadata):
            return WizardStepId.bridge_piles
        if self.is_fbs_draft(metadata):
            return WizardStepId.fbs
        if self.is_pile_draft(metadata):
            return WizardStepId.piles
        return WizardStepId.plates

    def should_skip_client_step(self, metadata: dict[str, Any]) -> bool:
        """Skip client when header is sticky (cycle ≥2, resume, or client already known)."""
        client_name = str(metadata.get("client_name") or "").strip()
        if client_name:
            return True
        append_batches = metadata.get("append_batches") or []
        if append_batches:
            return True
        resume_kp_id = metadata.get("resume_kp_id")
        return resume_kp_id is not None

    def wizard_step_order(self, metadata: dict[str, Any]) -> list[WizardStepId]:
        """Product → [client] → result; omits client when ``should_skip_client_step``."""
        product = self.product_step(metadata)
        if self.should_skip_client_step(metadata):
            return [product, WizardStepId.result]
        return [product, WizardStepId.client, WizardStepId.result]

    def wizard_step_after_plate_snapshot(self, metadata: dict[str, Any], order_data: list[Any]) -> WizardStepId:
        return self.product_step(metadata)

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
                and not self.invalid_width_lines_blocking(metadata)
                and not self.unpriced_lines_blocking(metadata)
                and self.meta_ready_for_calculate(metadata)
            ):
                return WizardStepId.result
            return WizardStepId.client

        if order_data and (
            self.wide_lines_blocking(metadata)
            or self.invalid_width_lines_blocking(metadata)
            or self.unpriced_lines_blocking(metadata)
        ):
            product_step = self.product_step(metadata)
            if stored in (WizardStepId.client, WizardStepId.result):
                return product_step
            return product_step

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
            if first == ERR_EMPTY_PILES:
                return WizardNextRequiredAction.ingest_piles
            if first == ERR_EMPTY_STEPS:
                return WizardNextRequiredAction.ingest_steps
            if first == ERR_EMPTY_MARCHES:
                return WizardNextRequiredAction.ingest_marches
            if first == ERR_EMPTY_BRIDGE_PILES:
                return WizardNextRequiredAction.ingest_bridge_piles
            if first == ERR_WIDE_PLATES:
                return WizardNextRequiredAction.resolve_wide_plates
            if first == ERR_INVALID_WIDTHS:
                return WizardNextRequiredAction.resolve_invalid_widths
            if first == ERR_UNPRICED_PLATES:
                return WizardNextRequiredAction.resolve_unpriced_plates
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

        if effective_step in (
            WizardStepId.plates,
            WizardStepId.piles,
            WizardStepId.steps,
            WizardStepId.marches,
            WizardStepId.bridge_piles,
            WizardStepId.fbs,
        ):
            if not order_data:
                return []
            if self.wide_lines_blocking(metadata):
                return []
            if self.invalid_width_lines_blocking(metadata):
                return []
            if self.unpriced_lines_blocking(metadata):
                return []
            if self.calculation_service.unpriced_position_labels(order_data):
                return []
            if self.should_skip_client_step(metadata):
                return [WizardStepId.result]
            return [WizardStepId.client]

        if effective_step == WizardStepId.client:
            if next_action == WizardNextRequiredAction.post_calculate:
                return []
            return []

        return []

    def collect_wizard_validation_errors(
        self,
        payload: dict[str, Any],
        next_action: WizardNextRequiredAction,
        *,
        effective_step: WizardStepId | None = None,
    ) -> list[str]:
        """Сообщения, согласованные с проверками ``calculate_draft`` / ``next_required_action``.

        На шаге ввода позиций (плиты/сваи/ступени) не показываем ошибки шага «Клиент»
        (менеджер, клиент, условия) — они нужны только на client / calculate.
        """
        if next_action == WizardNextRequiredAction.none:
            return []

        metadata = dict(payload.get("metadata") or {})
        order_data = payload.get("order_data") or []
        errors = self.calculation_service.validate_calculate_prerequisites(
            order_data=order_data,
            metadata=metadata,
        )

        messages: list[str] = []
        if next_action == WizardNextRequiredAction.complete_client_terms:
            messages.extend(errors)
        else:
            action_message = {
                WizardNextRequiredAction.ingest_plates: ERR_EMPTY_PLATES,
                WizardNextRequiredAction.ingest_piles: ERR_EMPTY_PILES,
                WizardNextRequiredAction.ingest_steps: ERR_EMPTY_STEPS,
                WizardNextRequiredAction.ingest_marches: ERR_EMPTY_MARCHES,
                WizardNextRequiredAction.ingest_bridge_piles: ERR_EMPTY_BRIDGE_PILES,
                WizardNextRequiredAction.ingest_fbs: ERR_EMPTY_FBS,
                WizardNextRequiredAction.resolve_wide_plates: ERR_WIDE_PLATES,
                WizardNextRequiredAction.resolve_invalid_widths: ERR_INVALID_WIDTHS,
                WizardNextRequiredAction.resolve_unpriced_plates: ERR_UNPRICED_PLATES,
                WizardNextRequiredAction.select_manager: ERR_NO_MANAGER,
            }.get(next_action)
            if action_message is not None:
                messages.append(action_message)

        unpriced = self.calculation_service.unpriced_position_labels(order_data)
        if unpriced:
            messages.append(f"Нет цен для позиций: {', '.join(unpriced)}")

        step = effective_step if effective_step is not None else self.infer_wizard_current_step(payload)
        if step in (
            WizardStepId.plates,
            WizardStepId.piles,
            WizardStepId.steps,
            WizardStepId.marches,
            WizardStepId.bridge_piles,
            WizardStepId.fbs,
        ):
            client_step_errors = {
                ERR_NO_MANAGER,
                ERR_NO_CLIENT,
                ERR_NO_DELIVERY,
                ERR_NO_PAYMENT,
            }
            messages = [msg for msg in messages if msg not in client_step_errors]

        return messages

    def build_wizard_state(self, payload: dict[str, Any]) -> dict[str, Any]:
        current = self.infer_wizard_current_step(payload)
        next_action = self.infer_next_required_action(payload, current)
        can_proceed = self.infer_can_proceed_to(payload, current, next_action)
        validation_errors = self.collect_wizard_validation_errors(
            payload, next_action, effective_step=current
        )
        return {
            "current_step": current,
            "can_proceed_to": can_proceed,
            "next_required_action": next_action,
            "validation_errors": validation_errors,
        }
