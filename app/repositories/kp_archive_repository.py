from __future__ import annotations

from typing import Literal

from app.core.settings import get_settings
from core import kp_db


ArchiveSection = Literal["archived", "in_production", "completed"]


class KpArchiveRepository:
    """Тонкая обёртка над core.kp_db для доступа к сохранённым КП и их статусам."""

    def __init__(self, db_path: str | None = None) -> None:
        settings = get_settings()
        self.db_path = db_path or str(settings.plita_db_path)
        kp_db.init_schema(self.db_path)

    def list_by_section(self, section: ArchiveSection) -> list[dict]:
        grouped = kp_db.get_all_kp_list(self.db_path)
        return list(grouped.get(section, []))

    def get_by_id(self, kp_id: int) -> dict | None:
        return kp_db.get_kp_by_id(kp_id, self.db_path)

    def get_completion_percentage(self, kp_id: int) -> dict:
        return kp_db.get_kp_completion_percentage(kp_id, self.db_path)

    def update_discount(self, kp_id: int, discount_percent: float) -> bool:
        return kp_db.update_kp_discount(kp_id, discount_percent, self.db_path)

    def update_logistics_cost(self, kp_id: int, logistics_cost: float) -> bool:
        return kp_db.update_kp_logistics_cost(kp_id, logistics_cost, self.db_path)

    def delete(self, kp_id: int) -> bool:
        return kp_db.delete_kp_by_id(kp_id, self.db_path)

    def update_status(self, kp_id: int, new_status: str) -> bool:
        return kp_db.update_kp_status(kp_id, new_status, self.db_path)

    def update_execution_date(self, kp_id: int, new_date: str) -> bool:
        return kp_db.update_kp_execution_date(kp_id, new_date, self.db_path)
