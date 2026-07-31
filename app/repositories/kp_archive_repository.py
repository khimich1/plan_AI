from __future__ import annotations

from typing import Literal

from app.core.settings import get_settings
from app.repositories.kp_offers_repository import KpOffersRepository
from core.kp import offers_write


ArchiveSection = Literal["archived", "in_production", "completed"]


class KpArchiveRepository:
    """Archive access to saved KP — reads via KpOffersRepository, writes via offers_write."""

    def __init__(self, db_path: str | None = None) -> None:
        settings = get_settings()
        self.db_path = db_path or str(settings.plita_db_path)
        self._offers = KpOffersRepository(self.db_path)

    def list_by_section(
        self,
        section: ArchiveSection,
        *,
        product_type: str | None = None,
        **list_filters,
    ) -> list[dict]:
        return self._offers.list_by_section(section, product_type=product_type, **list_filters)

    def get_by_id(self, kp_id: int) -> dict | None:
        return self._offers.get_by_id(kp_id)

    def search_by_customer_name(
        self,
        name: str,
        *,
        limit: int = 50,
        product_type: str | None = None,
        **list_filters,
    ) -> tuple[list[dict], int]:
        return self._offers.search_by_customer_name(
            name,
            limit=limit,
            product_type=product_type,
            **list_filters,
        )

    def get_completion_percentage(self, kp_id: int) -> dict:
        return self._offers.get_completion_percentage(kp_id)

    def update_discount(self, kp_id: int, discount_percent: float) -> bool:
        return offers_write.update_kp_discount(kp_id, discount_percent, self.db_path)

    def update_logistics_cost(self, kp_id: int, logistics_cost: float) -> bool:
        return offers_write.update_kp_logistics_cost(kp_id, logistics_cost, self.db_path)

    def update_status(self, kp_id: int, status: str) -> bool:
        return offers_write.update_kp_status(kp_id, status, self.db_path)

    def update_execution_date(self, kp_id: int, execution_date: str) -> bool:
        return offers_write.update_kp_execution_date(kp_id, execution_date, self.db_path)

    def delete(self, kp_id: int) -> bool:
        return offers_write.delete_kp_by_id(kp_id, self.db_path)
