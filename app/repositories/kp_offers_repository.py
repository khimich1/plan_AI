from __future__ import annotations

from typing import Literal

from app.core.settings import get_settings
from core.kp import offers_read

ArchiveSection = Literal["archived", "in_production", "completed"]


class KpOffersRepository:
    """Read/query access to KP_offers — delegates to core.kp.offers_read."""

    def __init__(self, db_path: str | None = None) -> None:
        settings = get_settings()
        self.db_path = db_path or str(settings.plita_db_path)

    def get_by_id(self, kp_id: int) -> dict | None:
        return offers_read.get_kp_by_id(kp_id, self.db_path)

    def list_by_status(self, status: str) -> list[dict]:
        return offers_read.get_all_kp_by_status(status, self.db_path)

    def list_grouped(
        self,
        *,
        owner_user_id: int | None = None,
        readable_statuses: tuple[str, ...] | None = None,
        deny_all: bool = False,
        product_type: str | None = None,
    ) -> dict[str, list[dict]]:
        return offers_read.get_all_kp_list(
            self.db_path,
            owner_user_id=owner_user_id,
            readable_statuses=readable_statuses,
            deny_all=deny_all,
            product_type=product_type,
        )

    def list_by_section(
        self,
        section: ArchiveSection,
        *,
        product_type: str | None = None,
        **list_filters,
    ) -> list[dict]:
        grouped = self.list_grouped(product_type=product_type, **list_filters)
        return list(grouped.get(section, []))

    def search_by_customer_name(
        self,
        name: str,
        *,
        limit: int = 50,
        owner_user_id: int | None = None,
        readable_statuses: tuple[str, ...] | None = None,
        deny_all: bool = False,
        product_type: str | None = None,
    ) -> tuple[list[dict], int]:
        return offers_read.search_kp_by_customer_name(
            name,
            limit=limit,
            db_path=self.db_path,
            owner_user_id=owner_user_id,
            readable_statuses=readable_statuses,
            deny_all=deny_all,
            product_type=product_type,
        )

    def get_xlsx_file(self, kp_id: int, output_path: str | None = None) -> bytes | None:
        return offers_read.get_xlsx_file(kp_id, output_path, self.db_path)

    def get_db_stats(self) -> dict[str, int]:
        return offers_read.get_db_stats(self.db_path)

    def get_next_kp_number(self) -> int:
        return offers_read.get_next_kp_number(self.db_path)

    def get_completion_percentage(self, kp_id: int) -> dict:
        return offers_read.get_kp_completion_percentage(kp_id, self.db_path)

    def get_plates_in_plan_percentage(self, kp_id: int) -> dict:
        return offers_read.get_kp_plates_in_plan_percentage(kp_id, self.db_path)

    def get_total_length(self, kp_id: int) -> float:
        return offers_read.get_kp_total_length(kp_id, self.db_path)
