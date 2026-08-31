from __future__ import annotations

from typing import Any

from core.commercial_offer_xlsx import DB_PATH
from core.commercial_pricing import calculate_total_cost, ensure_order_priced

# Wizard / calculate prerequisite messages (single source of truth).
ERR_EMPTY_PLATES = "Список плит пустой."
ERR_EMPTY_PILES = "Список свай пустой."
ERR_EMPTY_STEPS = "Список ступеней пустой."
ERR_EMPTY_MARCHES = "Список маршей пустой."
ERR_EMPTY_BRIDGE_PILES = "Список мостовых свай пустой."
ERR_EMPTY_FBS = "Список ФБС пустой."
ERR_WIDE_PLATES = "Сначала примите решение по позициям шире стандартной."
ERR_UNPRICED_PLATES = "Сначала примите решение по позициям без цены в прайсе."
ERR_NO_MANAGER = "Выберите менеджера."
ERR_NO_CLIENT = "Укажите клиента."
ERR_NO_DELIVERY = "Укажите условия поставки."
ERR_NO_PAYMENT = "Укажите условия оплаты."


class CommercialCalculationService:
    """Totals and prerequisite validation for commercial draft calculation."""

    @staticmethod
    def is_pile_draft(metadata: dict[str, Any]) -> bool:
        return str(metadata.get("product_type", "plates") or "plates").lower() == "piles"

    @staticmethod
    def is_step_draft(metadata: dict[str, Any]) -> bool:
        return str(metadata.get("product_type", "plates") or "plates").lower() == "steps"

    @staticmethod
    def is_march_draft(metadata: dict[str, Any]) -> bool:
        return str(metadata.get("product_type", "plates") or "plates").lower() == "marches"

    @staticmethod
    def is_bridge_pile_draft(metadata: dict[str, Any]) -> bool:
        return str(metadata.get("product_type", "plates") or "plates").lower() == "bridge_piles"

    @staticmethod
    def is_fbs_draft(metadata: dict[str, Any]) -> bool:
        return str(metadata.get("product_type", "plates") or "plates").lower() == "fbs"

    @staticmethod
    def order_has_plates(order_data: list[Any]) -> bool:
        """True if any line is plates; missing product_type counts as plates (legacy)."""
        for item in order_data or []:
            if not isinstance(item, dict):
                continue
            line_type = str(item.get("product_type") or "plates").strip().lower()
            if line_type == "plates":
                return True
        return False

    def _non_plate_cycle(self, metadata: dict[str, Any]) -> bool:
        return (
            self.is_pile_draft(metadata)
            or self.is_step_draft(metadata)
            or self.is_march_draft(metadata)
            or self.is_bridge_pile_draft(metadata)
            or self.is_fbs_draft(metadata)
        )

    def _wide_plate_errors(
        self,
        metadata: dict[str, Any],
        *,
        order_data: list[Any] | None = None,
    ) -> list[str]:
        # Prefer actual plate lines over metadata cycle type (mixed drafts).
        if order_data is not None:
            if not self.order_has_plates(order_data):
                return []
        elif self._non_plate_cycle(metadata):
            return []
        lines = metadata.get("wide_plate_lines") or []
        if lines and not metadata.get("wide_plates_resolved"):
            return [ERR_WIDE_PLATES]
        return []

    def _unpriced_plate_errors(
        self,
        metadata: dict[str, Any],
        *,
        order_data: list[Any] | None = None,
    ) -> list[str]:
        if order_data is not None:
            if not self.order_has_plates(order_data):
                return []
        elif self._non_plate_cycle(metadata):
            return []
        lines = metadata.get("unpriced_plate_lines") or []
        if lines and not metadata.get("unpriced_plates_resolved"):
            return [ERR_UNPRICED_PLATES]
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
            if self.is_step_draft(metadata):
                errors.append(ERR_EMPTY_STEPS)
            elif self.is_march_draft(metadata):
                errors.append(ERR_EMPTY_MARCHES)
            elif self.is_bridge_pile_draft(metadata):
                errors.append(ERR_EMPTY_BRIDGE_PILES)
            elif self.is_fbs_draft(metadata):
                errors.append(ERR_EMPTY_FBS)
            elif self.is_pile_draft(metadata):
                errors.append(ERR_EMPTY_PILES)
            else:
                errors.append(ERR_EMPTY_PLATES)
        errors.extend(self._wide_plate_errors(metadata, order_data=order_data))
        errors.extend(self._unpriced_plate_errors(metadata, order_data=order_data))
        errors.extend(self._metadata_errors(metadata))
        return errors

    def wide_lines_blocking(
        self,
        metadata: dict[str, Any],
        *,
        order_data: list[Any] | None = None,
    ) -> bool:
        return bool(self._wide_plate_errors(metadata, order_data=order_data))

    def unpriced_lines_blocking(
        self,
        metadata: dict[str, Any],
        *,
        order_data: list[Any] | None = None,
    ) -> bool:
        return bool(self._unpriced_plate_errors(metadata, order_data=order_data))

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
