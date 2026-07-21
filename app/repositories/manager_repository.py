from __future__ import annotations

from app.core.settings import get_settings
from core.kp_db import get_all_managers, get_manager_by_id


class ManagerRepository:
    def __init__(self, db_path: str | None = None) -> None:
        settings = get_settings()
        self.db_path = db_path or str(settings.plita_db_path)

    def list_managers(self) -> list[dict]:
        return get_all_managers(self.db_path)

    def get_manager(self, manager_id: int) -> dict | None:
        return get_manager_by_id(manager_id, self.db_path)

