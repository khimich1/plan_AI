from __future__ import annotations

from typing import Any

from core.commercial_offer_xlsx import DB_PATH
from core.commercial_pricing import calculate_total_cost, ensure_order_priced

# Wizard / calculate prerequisite messages (single source of truth).
ERR_EMPTY_PLATES = "Список плит пустой."
ERR_EMPTY_PILES = "Список свай пустой."
ERR_WIDE_PLATES = "Сначала примите решение по позициям шире стандартной."
ERR_NO_MANAGER = "Выберите менеджера."
ERR_NO_CLIENT = "Укажите клиента."
ERR_NO_DELIVERY = "Укажите условия поставки."
ERR_NO_PAYMENT = "Укажите условия оплаты."


class CommercialCalculationService:
    """Totals and prerequisite validation for commercial draft calculation."""

    @staticmethod
    def is_pile_draft(metadata: dict[str, Any]) -> bool:
        return str(metadata.get("product_type", "plates") or "plates").lower() == "piles"

    def _wide_plate_errors(self, metadata: dict[str, Any]) -> list[str]:
        if self.is_pile_draft(metadata):
            return []
        lines = metadata.get("wide_plate_lines") or []
        if lines and not metadata.get("wide_plates_resolved"):
            return [ERR_WIDE_PLATES]
        return []

    def _metadata_errors(self, metadata: dict[str, Any]) -> list[str]:
        errors: list[str] = []
        if not metadata.get("manager_id"):
            errors.append(ERR_NO_MANAGER)
        if not str(metadata.get("client_name", "")).strip():
            errors.append(ERR_NO_CLIENT)
        if metadata.get("conditions_mode") == "custom":
            if not str(metadata.get("delivery_conditions", "")).strip():
                errors.append(ERR_NO_DELIVERY)
            if not str(metadata.get("payment_conditions", "")).strip():
                errors.append(ERR_NO_PAYMENT)
        return errors

    def validate_calculate_prerequisites(
        self,
        *,
        order_data: list[Any],
        metadata: dict[str, Any],
    ) -> list[str]:
        errors: list[str] = []
        if order_data == []:
            errors.append(
                ERR_EMPTY_PILES if self.is_pile_draft(metadata) else ERR_EMPTY_PLATES
            )
        errors.extend(self._wide_plate_errors(metadata))
        errors.extend(self._metadata_errors(metadata))
        return errors

    def wide_lines_blocking(self, metadata: dict[str, Any]) -> bool:
        return bool(self._wide_plate_errors(metadata))

    def meta_ready_for_calculate(self, metadata: dict[str, Any]) -> bool:
        return not self._metadata_errors(metadata)

    def enforce_calculate_prerequisites(
        self,
        *,
        order_data: list[Any],
        metadata: dict[str, Any],
    ) -> None:
        errors = self.validate_calculate_prerequisites(
            order_data=order_data,
            metadata=metadata,
        )
        if errors:
            raise ValueError(errors[0])
        ensure_order_priced(list(order_data), db_path=str(DB_PATH))

    def unpriced_position_labels(self, order_data: list[Any]) -> list[str]:
        from core.commercial_pricing import collect_unpriced_positions

        return collect_unpriced_positions(list(order_data), db_path=str(DB_PATH))

    def compute_totals(
        self,
        order_data: list[Any],
        *,
        discount_percent: float,
        logistics_cost: float,
        require_all_priced: bool = False,
    ) -> dict[str, Any]:
        return calculate_total_cost(
            order_data,
            discount_percent,
            logistics_cost=logistics_cost,
            db_path=str(DB_PATH),
            require_all_priced=require_all_priced,
        )
