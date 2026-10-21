"""Контакты менеджера для шапки PDF и XLSX из архива.

Совпадение ФИО — strip и casefold, первая карточка в переданном списке.
Возвращает телефон и почту. Id менеджера не возвращается и не подставляется.
"""

from __future__ import annotations

from pathlib import Path

from app.repositories.manager_repository import ManagerRepository


def lookup_manager_contacts(
    manager_name: str | None,
    managers: list[dict],
) -> tuple[str, str]:
    name = (manager_name or "").strip()
    if not name:
        return "", ""
    needle = name.casefold()
    for manager in managers:
        fio = str(manager.get("fio") or "").strip()
        if fio.casefold() != needle:
            continue
        return (
            str(manager.get("contact_number") or "").strip(),
            str(manager.get("email") or "").strip(),
        )
    return "", ""


def header_contacts_for_archive(
    manager_name: str | None,
    managers: list[dict],
) -> tuple[str | None, str | None, bool]:
    name = (manager_name or "").strip()
    if not name:
        return None, None, False
    phone, email = lookup_manager_contacts(name, managers)
    return phone or None, email or None, True


def load_managers_if_db_exists(db_path: object) -> list[dict]:
    """Читает справочник только если db_path — существующий файл.

    Иначе пустой список и без sqlite.connect: тесты с MagicMock и
    несуществующим plita.db не должны создавать файл базы.
    """
    if not isinstance(db_path, (str, Path)):
        return []
    if not Path(db_path).is_file():
        return []
    return ManagerRepository(str(db_path)).list_managers()
