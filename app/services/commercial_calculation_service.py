from __future__ import annotations

from typing import Any

from core.commercial_offer_xlsx import calculate_total_cost


class CommercialCalculationService:
    """Totals and prerequisite validation for commercial draft calculation."""

    def wide_lines_blocking(self, metadata: dict[str, Any]) -> bool:
        lines = metadata.get("wide_plate_lines") or []
        return bool(lines) and not bool(metadata.get("wide_plates_resolved"))

    def meta_ready_for_calculate(self, metadata: dict[str, Any]) -> bool:
        if not metadata.get("manager_id"):
            return False
        if not str(metadata.get("client_name", "")).strip():
            return False
        if metadata.get("conditions_mode") == "custom":
            if not str(metadata.get("delivery_conditions", "")).strip():
                return False
            if not str(metadata.get("payment_conditions", "")).strip():
                return False
        return True

    def validate_calculate_prerequisites(
        self,
        *,
        order_data: list[Any],
        metadata: dict[str, Any],
    ) -> None:
        if order_data == []:
            raise ValueError("Список плит пустой.")
        if metadata.get("wide_plate_lines") and not metadata.get("wide_plates_resolved"):
            raise ValueError("Сначала обработайте плиты шире 12 дм.")
        if not metadata.get("manager_id"):
            raise ValueError("Выберите менеджера.")
        if not str(metadata.get("client_name", "")).strip():
            raise ValueError("Укажите клиента.")
        if metadata.get("conditions_mode") == "custom":
            if not str(metadata.get("delivery_conditions", "")).strip():
                raise ValueError("Укажите условия поставки.")
            if not str(metadata.get("payment_conditions", "")).strip():
                raise ValueError("Укажите условия оплаты.")

    def compute_totals(
        self,
        order_data: list[Any],
        *,
        discount_percent: float,
        logistics_cost: float,
    ) -> dict[str, Any]:
        return calculate_total_cost(
            order_data,
            discount_percent,
            logistics_cost=logistics_cost,
        )
