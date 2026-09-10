"""Directory of 1C counterparties: search, create, resolve for KP save."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from app.repositories.counterparties_repository import CounterpartiesRepository


class DuplicateCodeError(ValueError):
    """Повторный ``code_1c`` при ручном создании."""

    def __init__(self, existing: dict[str, Any]) -> None:
        self.existing = existing
        super().__init__("Контрагент с таким кодом 1С уже есть")


class CounterpartyValidationError(ValueError):
    """Контрагент не подходит для нового КП (нет / не клиент / неактивен)."""


@dataclass(frozen=True)
class CreateResult:
    item: dict[str, Any]
    warning: str | None = None


class CounterpartiesService:
    def __init__(
        self,
        *,
        db_path: str,
        repository: CounterpartiesRepository | None = None,
    ) -> None:
        self.db_path = db_path
        self.repo = repository or CounterpartiesRepository(db_path=db_path)

    def get_by_id(self, counterparty_id: int) -> dict[str, Any] | None:
        return self.repo.get_by_id(counterparty_id)

    def get_by_code_1c(self, code_1c: str) -> dict[str, Any] | None:
        return self.repo.get_by_code_1c(code_1c)

    def search(
        self,
        q: str,
        *,
        limit: int = 10,
        clients_only: bool = True,
    ) -> list[dict[str, Any]]:
        return self.repo.search(q, limit=limit, clients_only=clients_only)

    def require_active_client(self, counterparty_id: int | None) -> dict[str, Any]:
        if counterparty_id is None:
            raise CounterpartyValidationError("Контрагент не найден в справочнике")
        row = self.get_by_id(int(counterparty_id))
        if row is None or not int(row.get("is_active") or 0):
            raise CounterpartyValidationError("Контрагент не найден в справочнике")
        if not int(row.get("is_client") or 0):
            raise CounterpartyValidationError("Контрагент не отмечен как клиент в 1С")
        return row

    def create(
        self,
        *,
        name: str,
        code_1c: str,
        inn: str | None = None,
        kpp: str | None = None,
        is_client: bool = True,
    ) -> CreateResult:
        code = (code_1c or "").strip()
        trimmed_name = (name or "").strip()
        if not code or not trimmed_name:
            raise ValueError("Наименование и код 1С обязательны")
        existing = self.repo.get_by_code_1c(code)
        if existing is not None:
            raise DuplicateCodeError(existing)

        inn_value = (inn or "").strip() or None
        kpp_value = (kpp or "").strip() or None
        warning: str | None = None
        if inn_value:
            twins = self.repo.find_by_inn(inn_value)
            if twins:
                warning = "В справочнике уже есть контрагент с таким ИНН"

        item = self.repo.insert(
            code_1c=code,
            name=trimmed_name,
            inn=inn_value,
            kpp=kpp_value,
            is_client=is_client,
            source="manual",
        )
        return CreateResult(item=item, warning=warning)
