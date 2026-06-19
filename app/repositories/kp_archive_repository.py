from __future__ import annotations

from typing import Literal

from app.core.settings import get_settings
from core import kp_db_offers


ArchiveSection = Literal["archived", "in_production", "completed"]


class KpArchiveRepository:
    """Тонкая обёртка над core.kp_db_offers для доступа к сохранённым КП."""

    def __init__(self, db_path: str | None = None) -> None:
        settings = get_settings()
        self.db_path = db_path or str(settings.plita_db_path)

    def list_by_section(self, section: ArchiveSection, **list_filters) -> list[dict]:
        grouped = kp_db_offers.get_all_kp_list(self.db_path, **list_filters)
        return list(grouped.get(section, []))

    def get_by_id(self, kp_id: int) -> dict | None:
        return kp_db_offers.get_kp_by_id(kp_id, self.db_path)

    def search_by_customer_name(
        self,
        name: str,
        *,
        limit: int = 50,
        **list_filters,
    ) -> tuple[list[dict], int]:
        return kp_db_offers.search_kp_by_customer_name(
            name,
            limit=limit,
            db_path=self.db_path,
            **list_filters,
        )

    def get_completion_percentage(self, kp_id: int) -> dict:
        return kp_db_offers.get_kp_completion_percentage(kp_id, self.db_path)

    def update_discount(self, kp_id: int, discount_percent: float) -> bool:
        return kp_db_offers.update_kp_discount(kp_id, discount_percent, self.db_path)

    def update_logistics_cost(self, kp_id: int, logistics_cost: float) -> bool:
        return kp_db_offers.update_kp_logistics_cost(kp_id, logistics_cost, self.db_path)

    def update_status(self, kp_id: int, status: str) -> bool:
        return kp_db_offers.update_kp_status(kp_id, status, self.db_path)

    def update_execution_date(self, kp_id: int, execution_date: str) -> bool:
        return kp_db_offers.update_kp_execution_date(kp_id, execution_date, self.db_path)

    def delete(self, kp_id: int) -> bool:
        return kp_db_offers.delete_kp_by_id(kp_id, self.db_path)
