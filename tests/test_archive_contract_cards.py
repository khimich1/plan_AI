"""Карточки договора: вставка (CDC-003), номер, имя и реквизиты (CDC-004), сервис (CDC-005)."""

from __future__ import annotations

import sqlite3
from datetime import date, datetime
from pathlib import Path

import pytest
from fastapi import HTTPException

from app.repositories.supply_contract_repository import SupplyContractRepository
from app.schemas.supply_contract import SupplyContractCreate
from app.services.supply_contract_service import (
    SupplyContractNotFoundError,
    SupplyContractService,
    SupplyContractValidationError,
)
from core import kp_db_schema
from core.kp_db_common import _connect

_BLANK_COLUMNS = (
    "counterparty_id",
    "contract_date",
    "manager_name",
    "scan_note",
    "scan_path",
    "legal_form",
    "full_name",
    "short_name",
    "signatory_position",
    "signatory_name",
    "signatory_verb",
    "authority_basis",
    "poa_number",
    "poa_date",
    "inn",
    "kpp",
    "ogrn",
    "legal_address",
    "postal_address",
    "phone",
    "email",
    "bank_name",
    "account",
    "corr_account",
    "bik",
    "edo_operator",
    "edo_id",
    "okved",
)

# Колонки поиска: ключи имени/ИНН/номера и поля, которые отдаёт карточка.
_SEARCH_COLUMNS = (
    "id",
    "imported_name",
    "full_name",
    "short_name",
    "inn",
    "number",
    "counterparty_id",
    "contract_date",
    "manager_name",
    "status",
    "scan_note",
    "scan_path",
    "legal_form",
    "signatory_position",
    "signatory_name",
    "signatory_verb",
    "authority_basis",
    "poa_number",
    "poa_date",
    "kpp",
    "ogrn",
    "legal_address",
    "postal_address",
    "phone",
    "email",
    "bank_name",
    "account",
    "corr_account",
    "bik",
    "edo_operator",
    "edo_id",
    "okved",
)


def _repo(tmp_path: Path) -> SupplyContractRepository:
    db_path = str(tmp_path / "cards.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return SupplyContractRepository(db_path=db_path)


def _service(tmp_path: Path, *, today: date | None = None) -> SupplyContractService:
    db_path = str(tmp_path / "cards.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return SupplyContractService(db_path=db_path, today=today)


def _writer(role: str = "manager", user_id: int = 7) -> dict:
    return {"id": user_id, "role": role, "username": role}


def _count(db_path: str, table: str) -> int:
    with _connect(db_path) as conn:
        row = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()
    return int(row[0])


def _events(db_path: str) -> list[dict]:
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT contract_id, field, old_value, new_value, user_id, created_at
            FROM supply_contract_event
            ORDER BY id
            """
        ).fetchall()
    return [dict(row) for row in rows]


def _seed_counterparty(db_path: str) -> None:
    with _connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO counterparties (code_1c, name, name_normalized, inn)
            VALUES ('00-0001', 'ООО Ромашка', 'ооо ромашка', '7701000001')
            """
        )


def test_insert_card_without_requisites_or_date_reads_back_nulls(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    stored = repo.insert_card(
        imported_name="ИП Секлюцкий",
        number="0786/1/25",
        user_id=7,
    )
    loaded = repo.get_by_number("0786/1/25")

    assert loaded == stored
    assert loaded["imported_name"] == "ИП Секлюцкий"
    assert loaded["number"] == "0786/1/25"
    assert loaded["status"] == "нет"
    assert loaded["created_by_user_id"] == 7
    assert isinstance(loaded["created_at"], str) and loaded["created_at"]
    for column in _BLANK_COLUMNS:
        assert loaded[column] is None, column
    assert _count(repo.db_path, "supply_contract") == 1
    assert _count(repo.db_path, "supply_contract_event") == 0


def test_insert_card_stores_inn_and_date_without_other_requisites(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    repo.insert_card(
        imported_name="ИП Секлюцкий",
        number="0786/1/25",
        user_id=7,
        inn="440117229007",
        contract_date=date(2025, 1, 1),
    )

    loaded = repo.get_by_number("0786/1/25")
    assert loaded is not None
    assert loaded["inn"] == "440117229007"
    assert loaded["contract_date"] == "2025-01-01"
    assert loaded["status"] == "нет"
    assert loaded["manager_name"] is None
    assert loaded["counterparty_id"] is None
    assert loaded["legal_form"] is None
    assert loaded["full_name"] is None
    assert loaded["email"] is None


def test_same_number_does_not_create_a_second_row(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    repo.insert_card(imported_name="ИП Секлюцкий", number="0786/1/25", user_id=7)

    with pytest.raises(sqlite3.IntegrityError) as caught:
        repo.insert_card(imported_name="Другое имя", number="0786/1/25", user_id=8)

    assert "в реестре" not in str(caught.value)
    assert _count(repo.db_path, "supply_contract") == 1
    loaded = repo.get_by_number("0786/1/25")
    assert loaded is not None
    assert loaded["imported_name"] == "ИП Секлюцкий"
    assert loaded["created_by_user_id"] == 7


def test_insert_card_does_not_add_a_counterparty(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    _seed_counterparty(repo.db_path)
    before = _count(repo.db_path, "counterparties")

    stored = repo.insert_card(
        imported_name="ИП Секлюцкий",
        number="0786/1/25",
        user_id=7,
        inn="440117229007",
    )

    assert _count(repo.db_path, "counterparties") == before
    assert stored["counterparty_id"] is None


def test_get_by_number_returns_none_for_a_free_number(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    assert repo.get_by_number("0786/1/25") is None


def test_list_for_search_returns_name_inn_number_and_card_fields(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    repo.insert_card(
        imported_name="ИП Секлюцкий",
        number="0786/1/25",
        user_id=7,
        inn="440117229007",
    )
    with _connect(repo.db_path) as conn:
        conn.execute(
            """
            INSERT INTO supply_contract (
                imported_name, full_name, short_name, inn, number,
                contract_date, manager_name, status, legal_form,
                created_at, created_by_user_id
            ) VALUES (
                NULL, 'ООО Ромашка', 'Ромашка', '7701000001', '0001/10/26',
                '2026-10-01', 'Иван', 'есть', 'ooo',
                '2026-10-01T09:00:00', 7
            )
            """
        )

    rows = repo.list_for_search()
    by_number = {row["number"]: row for row in rows}

    assert set(by_number) == {"0786/1/25", "0001/10/26"}
    assert set(by_number["0786/1/25"]) == set(_SEARCH_COLUMNS)
    assert "created_at" not in by_number["0786/1/25"]
    assert "created_by_user_id" not in by_number["0786/1/25"]
    assert by_number["0786/1/25"]["imported_name"] == "ИП Секлюцкий"
    assert by_number["0786/1/25"]["inn"] == "440117229007"
    assert by_number["0786/1/25"]["full_name"] is None
    assert by_number["0001/10/26"]["imported_name"] is None
    assert by_number["0001/10/26"]["full_name"] == "ООО Ромашка"
    assert by_number["0001/10/26"]["short_name"] == "Ромашка"
    assert by_number["0001/10/26"]["status"] == "есть"


def test_change_number_keeps_contract_date_and_adds_one_event(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    stored = repo.insert_card(
        imported_name="ИП Секлюцкий",
        number="0786/1/25",
        user_id=7,
        inn="440117229007",
        contract_date=date(2025, 1, 1),
    )

    updated = repo.change_number(int(stored["id"]), " 1000 / 10 / 26 ", user_id=9)

    assert updated is not None
    assert updated["id"] == stored["id"]
    assert updated["number"] == "1000/10/26"
    assert updated["contract_date"] == "2025-01-01"
    assert updated["imported_name"] == "ИП Секлюцкий"
    assert updated["inn"] == "440117229007"
    assert repo.get_by_number("0786/1/25") is None
    assert repo.get_by_number("1000/10/26")["id"] == stored["id"]
    events = _events(repo.db_path)
    assert len(events) == 1
    event = events[0]
    assert event["contract_id"] == stored["id"]
    assert event["field"] == "number"
    assert event["old_value"] == "0786/1/25"
    assert event["new_value"] == "1000/10/26"
    assert event["user_id"] == 9
    datetime.fromisoformat(event["created_at"])


def test_same_number_after_normalize_does_not_write_event(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    stored = repo.insert_card(
        imported_name="ИП Секлюцкий",
        number="0786/1/25",
        user_id=7,
        contract_date=date(2025, 1, 1),
    )

    updated = repo.change_number(int(stored["id"]), "0786 / 1 / 25", user_id=9)

    assert updated is not None
    assert updated["number"] == "0786/1/25"
    assert updated["contract_date"] == "2025-01-01"
    assert _events(repo.db_path) == []
    assert _count(repo.db_path, "supply_contract") == 1


def test_change_number_without_a_digit_does_not_write(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    repo.insert_card(
        imported_name="ИП Секлюцкий",
        number="0786/1/25",
        user_id=7,
        contract_date=date(2025, 1, 1),
    )

    with pytest.raises(ValueError, match="В номере договора нужна хотя бы одна цифра"):
        repo.change_number(1, "без цифр", user_id=9)

    loaded = repo.get_by_number("0786/1/25")
    assert loaded is not None
    assert loaded["contract_date"] == "2025-01-01"
    assert _events(repo.db_path) == []


def test_occupied_number_raises_integrity_error_without_commit(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    first = repo.insert_card(
        imported_name="ИП Секлюцкий",
        number="0786/1/25",
        user_id=7,
        contract_date=date(2025, 1, 1),
    )
    second = repo.insert_card(
        imported_name="Трест",
        number="0793/02/25",
        user_id=7,
        contract_date=date(2025, 2, 1),
    )

    with pytest.raises(sqlite3.IntegrityError) as caught:
        repo.change_number(int(first["id"]), "0793 / 02 / 25", user_id=9)

    assert "в реестре" not in str(caught.value)
    assert repo.get_by_number("0786/1/25")["id"] == first["id"]
    assert repo.get_by_number("0786/1/25")["contract_date"] == "2025-01-01"
    assert repo.get_by_number("0793/02/25")["id"] == second["id"]
    assert repo.get_by_number("0793/02/25")["imported_name"] == "Трест"
    assert _events(repo.db_path) == []
    assert _count(repo.db_path, "supply_contract") == 2


def test_missing_card_updates_return_none(tmp_path: Path) -> None:
    repo = _repo(tmp_path)

    assert repo.change_number(404, "1000/10/26", user_id=9) is None
    assert repo.update_imported_name_and_inn(404, imported_name="Имя", inn="1") is None
    assert repo.save_requisites(404, {"legal_form": "ip"}, manager_name="ivan") is None
    assert _count(repo.db_path, "supply_contract") == 0
    assert _events(repo.db_path) == []


def test_update_imported_name_and_inn_does_not_write_event(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    stored = repo.insert_card(
        imported_name="ИП Секлюцкий",
        number="0786/1/25",
        user_id=7,
        inn="440117229007",
        contract_date=date(2025, 1, 1),
    )

    updated = repo.update_imported_name_and_inn(
        int(stored["id"]),
        imported_name="ИП Секлюцкий А.А.",
        inn="7702000002",
    )

    assert updated is not None
    assert updated["imported_name"] == "ИП Секлюцкий А.А."
    assert updated["inn"] == "7702000002"
    assert updated["number"] == "0786/1/25"
    assert updated["contract_date"] == "2025-01-01"
    assert updated["full_name"] is None
    assert _events(repo.db_path) == []


def test_save_requisites_keeps_number_and_existing_manager(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    stored = repo.insert_card(
        imported_name="ИП Секлюцкий",
        number="0786/1/25",
        user_id=7,
        inn="440117229007",
        contract_date=date(2025, 1, 1),
    )
    with _connect(repo.db_path) as conn:
        conn.execute(
            "UPDATE supply_contract SET manager_name = ? WHERE id = ?",
            ("Иван", int(stored["id"])),
        )

    updated = repo.save_requisites(
        int(stored["id"]),
        {
            "legal_form": "ip",
            "full_name": "Секлюцкий Иван Иванович",
            "short_name": "Секлюцкий",
            "email": "sek@example.com",
            "number": "9999/01/99",
            "imported_name": "Чужое имя",
            "contract_date": "1999-01-01",
            "manager_name": "Чужой менеджер",
        },
        manager_name="Пётр",
    )

    assert updated is not None
    assert updated["number"] == "0786/1/25"
    assert updated["contract_date"] == "2025-01-01"
    assert updated["manager_name"] == "Иван"
    assert updated["imported_name"] == "ИП Секлюцкий"
    assert updated["inn"] == "440117229007"
    assert updated["legal_form"] == "ip"
    assert updated["full_name"] == "Секлюцкий Иван Иванович"
    assert updated["short_name"] == "Секлюцкий"
    assert updated["email"] == "sek@example.com"
    assert updated["phone"] is None
    assert _events(repo.db_path) == []


def test_save_requisites_writes_date_only_when_given(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    stored = repo.insert_card(
        imported_name="ИП Секлюцкий",
        number="0786/1/25",
        user_id=7,
        contract_date=date(2025, 1, 1),
    )

    kept = repo.save_requisites(
        int(stored["id"]),
        {"legal_form": "ip"},
        manager_name="ivan",
    )
    changed = repo.save_requisites(
        int(stored["id"]),
        {"full_name": "Секлюцкий Иван Иванович"},
        manager_name="petr",
        contract_date=date(2026, 3, 15),
    )

    assert kept is not None
    assert kept["contract_date"] == "2025-01-01"
    assert kept["number"] == "0786/1/25"
    assert changed is not None
    assert changed["contract_date"] == "2026-03-15"
    assert changed["number"] == "0786/1/25"
    assert changed["legal_form"] == "ip"
    assert changed["full_name"] == "Секлюцкий Иван Иванович"
    assert _events(repo.db_path) == []


def test_save_requisites_sets_manager_only_when_empty(tmp_path: Path) -> None:
    repo = _repo(tmp_path)
    stored = repo.insert_card(
        imported_name="ИП Секлюцкий",
        number="0786/1/25",
        user_id=7,
        contract_date=date(2025, 1, 1),
    )

    filled = repo.save_requisites(
        int(stored["id"]),
        {"legal_form": "ip"},
        manager_name="  ivan  ",
    )
    again = repo.save_requisites(
        int(stored["id"]),
        {"legal_form": "ooo"},
        manager_name="petr",
    )

    assert filled is not None
    assert filled["manager_name"] == "ivan"
    assert again is not None
    assert again["manager_name"] == "ivan"
    assert again["legal_form"] == "ooo"
    assert again["number"] == "0786/1/25"
    assert again["contract_date"] == "2025-01-01"


def test_create_card_without_number_uses_next_seq_and_card_date(tmp_path: Path) -> None:
    service = _service(tmp_path, today=date(2026, 10, 20))
    service.repository.insert_card(
        imported_name="Занятый префикс",
        number="1033/09/26",
        user_id=1,
        contract_date=date(2026, 9, 1),
    )
    service.repository.insert_card(
        imported_name="Не в счётчик",
        number="ТБ-1/10/25",
        user_id=1,
        contract_date=date(2025, 10, 1),
    )
    service.repository.insert_card(
        imported_name="Дефисный",
        number="0844-06-25",
        user_id=1,
        contract_date=date(2025, 6, 1),
    )

    created = service.create_card(name="  Новый контрагент  ", user=_writer())

    assert created["number"] == "1034/10/26"
    assert created["contract_date"] == "2026-10-20"
    assert created["imported_name"] == "Новый контрагент"
    assert created["name"] == "Новый контрагент"
    assert created["inn"] is None
    assert created["status"] == "нет"
    assert created["counterparty_id"] is None
    assert created["manager_name"] is None
    assert _count(service.repository.db_path, "counterparties") == 0
    assert _count(service.repository.db_path, "supply_contract") == 4


def test_manual_number_is_stored_literally_and_duplicate_uses_spec_text(
    tmp_path: Path,
) -> None:
    service = _service(tmp_path, today=date(2026, 10, 20))
    user = _writer()

    created = service.create_card(
        name="ИП Секлюцкий",
        inn="440117229007",
        number="0786/1/25",
        user=user,
    )

    assert created["number"] == "0786/1/25"
    assert created["contract_date"] == "2025-01-01"
    assert created["contract_date"] != "2026-10-20"
    assert created["inn"] == "440117229007"
    with pytest.raises(SupplyContractValidationError) as caught:
        service.create_card(
            name="Другой",
            number="0786 / 1 / 25",
            user=user,
        )
    assert str(caught.value) == "Номер 0786/1/25 уже есть"
    assert "реестре" not in str(caught.value)
    assert "У контрагента" not in str(caught.value)
    assert _count(service.repository.db_path, "supply_contract") == 1
    assert _count(service.repository.db_path, "counterparties") == 0


def test_number_without_month_and_year_stores_null_date(tmp_path: Path) -> None:
    service = _service(tmp_path, today=date(2026, 10, 20))

    created = service.create_card(name="Без даты", number="ТБ-1", user=_writer())

    assert created["number"] == "ТБ-1"
    assert created["contract_date"] is None


def test_change_number_keeps_date_writes_event_and_old_number_is_not_found(
    tmp_path: Path,
) -> None:
    service = _service(tmp_path, today=date(2026, 10, 20))
    user = _writer()
    created = service.create_card(
        name="ИП Секлюцкий",
        inn="440117229007",
        number="0786/1/25",
        user=user,
    )

    updated = service.update_card(
        int(created["id"]),
        number="1000/10/26",
        user=_writer(user_id=9),
    )

    assert updated["number"] == "1000/10/26"
    assert updated["contract_date"] == "2025-01-01"
    assert updated["imported_name"] == "ИП Секлюцкий"
    assert updated["inn"] == "440117229007"
    events = _events(service.repository.db_path)
    assert len(events) == 1
    assert events[0]["field"] == "number"
    assert events[0]["old_value"] == "0786/1/25"
    assert events[0]["new_value"] == "1000/10/26"
    assert events[0]["user_id"] == 9
    assert service.search_cards("0786/1/25", user=user) == []
    found = service.search_cards("1000/10", user=user)
    assert [row["number"] for row in found] == ["1000/10/26"]


def test_occupied_number_on_update_keeps_row_and_uses_spec_text(tmp_path: Path) -> None:
    service = _service(tmp_path, today=date(2026, 10, 20))
    user = _writer()
    first = service.create_card(name="ИП Секлюцкий", number="0786/1/25", user=user)
    service.create_card(name="Трест", number="0793/02/25", user=user)

    with pytest.raises(SupplyContractValidationError) as caught:
        service.update_card(
            int(first["id"]),
            number="0793 / 02 / 25",
            name="Не должно сохраниться",
            user=_writer(user_id=9),
        )

    assert str(caught.value) == "Номер 0793/02/25 уже есть"
    loaded = service.repository.get_by_id(int(first["id"]))
    assert loaded is not None
    assert loaded["number"] == "0786/1/25"
    assert loaded["contract_date"] == "2025-01-01"
    assert loaded["imported_name"] == "ИП Секлюцкий"
    assert _events(service.repository.db_path) == []


def test_update_name_and_inn_keeps_number_and_date(tmp_path: Path) -> None:
    service = _service(tmp_path, today=date(2026, 10, 20))
    created = service.create_card(
        name="ИП Секлюцкий",
        inn="440117229007",
        number="0786/1/25",
        user=_writer(),
    )

    updated = service.update_card(
        int(created["id"]),
        name="  ИП Секлюцкий А.А.  ",
        inn="7702000002",
        user=_writer(),
    )

    assert updated["imported_name"] == "ИП Секлюцкий А.А."
    assert updated["name"] == "ИП Секлюцкий А.А."
    assert updated["inn"] == "7702000002"
    assert updated["number"] == "0786/1/25"
    assert updated["contract_date"] == "2025-01-01"
    assert _events(service.repository.db_path) == []


def test_search_finds_three_cards_by_name_inn_and_number_fragment(tmp_path: Path) -> None:
    service = _service(tmp_path, today=date(2026, 10, 20))
    user = _writer()
    for name, number in (
        ("ИП Секлюцкий А.А.", "0793/02/25"),
        ("Секлюцкий", "0794/02/25"),
        ("ИП Секлюцкий", "0933/03/26"),
    ):
        service.create_card(name=name, inn="440117229007", number=number, user=user)
    service.create_card(name="Трест", inn="5001000000", number="1000/10/26", user=user)
    _seed_counterparty(service.repository.db_path)
    with _connect(service.repository.db_path) as conn:
        conn.execute(
            """
            INSERT INTO counterparties (code_1c, name, name_normalized, inn)
            VALUES ('00-0002', 'Секлюцкий из справочника', 'секлюцкий из справочника', '440117229007')
            """
        )

    by_name = service.search_cards("  секлюц  ", user=user)
    by_inn = service.search_cards("440117229007", user=user)
    by_number = service.search_cards("0793/02", user=user)
    by_upper = service.search_cards("СЕКЛЮЦ", user=user)

    assert [row["number"] for row in by_name] == ["0793/02/25", "0794/02/25", "0933/03/26"]
    assert [row["name"] for row in by_name] == [
        "ИП Секлюцкий А.А.",
        "Секлюцкий",
        "ИП Секлюцкий",
    ]
    assert [row["number"] for row in by_inn] == ["0793/02/25", "0794/02/25", "0933/03/26"]
    assert [row["number"] for row in by_number] == ["0793/02/25"]
    assert [row["number"] for row in by_upper] == ["0793/02/25", "0794/02/25", "0933/03/26"]
    assert [row["number"] for row in service.search_cards("", user=user)] == [
        "0793/02/25",
        "0794/02/25",
        "0933/03/26",
        "1000/10/26",
    ]
    assert service.search_cards("с", user=user) == []
    assert service.search_cards("%", user=user) == []
    assert service.search_cards("Секлюцкий из справочника", user=user) == []
    assert _count(service.repository.db_path, "counterparties") == 2


def test_search_treats_percent_and_underscore_as_literals(tmp_path: Path) -> None:
    service = _service(tmp_path, today=date(2026, 10, 20))
    user = _writer()
    service.create_card(name="ООО 100%", number="1100/01/26", user=user)
    service.create_card(name="ООО 1000", number="1101/01/26", user=user)
    service.create_card(name="A_B завод", number="1102/01/26", user=user)
    service.create_card(name="AXB завод", number="1103/01/26", user=user)

    assert [row["number"] for row in service.search_cards("0%", user=user)] == ["1100/01/26"]
    assert [row["number"] for row in service.search_cards("A_", user=user)] == ["1102/01/26"]
    assert service.search_cards("%%", user=user) == []


def test_search_name_falls_back_to_full_name_then_short_name(tmp_path: Path) -> None:
    service = _service(tmp_path, today=date(2026, 10, 20))
    user = _writer()
    with_full = service.create_card(name="Черновик", number="0001/10/26", user=user)
    with_short = service.create_card(name="Ещё черновик", number="0002/10/26", user=user)
    with_both = service.create_card(name="Завод", number="0003/10/26", user=user)
    with _connect(service.repository.db_path) as conn:
        conn.execute(
            """
            UPDATE supply_contract
            SET imported_name = NULL, full_name = ?, short_name = ?
            WHERE id = ?
            """,
            ("ООО Ромашка", "Ромашка", int(with_full["id"])),
        )
        conn.execute(
            """
            UPDATE supply_contract
            SET imported_name = '   ', full_name = NULL, short_name = ?
            WHERE id = ?
            """,
            ("Птицефабрика", int(with_short["id"])),
        )
        conn.execute(
            "UPDATE supply_contract SET full_name = ? WHERE id = ?",
            ("Птицефабрика Центральная", int(with_both["id"])),
        )

    by_full = service.search_cards("ромаш", user=user)
    by_short = service.search_cards("птицефабрика", user=user)

    assert [row["number"] for row in by_full] == ["0001/10/26"]
    assert by_full[0]["name"] == "ООО Ромашка"
    assert [row["number"] for row in by_short] == ["0002/10/26", "0003/10/26"]
    names = {row["number"]: row["name"] for row in by_short}
    assert names["0002/10/26"] == "Птицефабрика"
    assert names["0003/10/26"] == "Завод"


@pytest.mark.parametrize("role", ["admin", "manager"])
def test_admin_and_manager_can_create_and_search(tmp_path: Path, role: str) -> None:
    service = _service(tmp_path, today=date(2026, 10, 20))
    user = _writer(role=role, user_id=3)

    created = service.create_card(name="Ромашка", user=user)

    assert created["number"] == "0001/10/26"
    assert created["contract_date"] == "2026-10-20"
    found = service.search_cards("рома", user=user)
    assert [row["number"] for row in found] == ["0001/10/26"]
    assert found[0]["name"] == "Ромашка"


def test_other_role_cannot_create_search_or_update(tmp_path: Path) -> None:
    service = _service(tmp_path, today=date(2026, 10, 20))
    created = service.create_card(name="Ромашка", number="0786/1/25", user=_writer())
    outsider = {"id": 4, "role": "production", "username": "shop"}

    with pytest.raises(HTTPException) as denied_create:
        service.create_card(name="Чужой", user=outsider)
    with pytest.raises(HTTPException) as denied_search:
        service.search_cards("рома", user=outsider)
    with pytest.raises(HTTPException) as denied_update:
        service.update_card(int(created["id"]), number="1000/10/26", user=outsider)

    assert denied_create.value.status_code == 403
    assert denied_search.value.status_code == 403
    assert denied_update.value.status_code == 403
    loaded = service.repository.get_by_number("0786/1/25")
    assert loaded is not None
    assert loaded["contract_date"] == "2025-01-01"
    assert service.repository.get_by_number("1000/10/26") is None
    assert _count(service.repository.db_path, "supply_contract") == 1
    assert _events(service.repository.db_path) == []


def test_empty_name_does_not_insert_card(tmp_path: Path) -> None:
    service = _service(tmp_path, today=date(2026, 10, 20))

    with pytest.raises(SupplyContractValidationError):
        service.create_card(name="   ", user=_writer())

    assert _count(service.repository.db_path, "supply_contract") == 0


def _buyer() -> SupplyContractCreate:
    return SupplyContractCreate.model_validate(
        {
            "legal_form": "ooo",
            "full_name": "ООО Ромашка",
            "short_name": "Ромашка",
            "signatory_position": "Генерального директора",
            "signatory_name": "Иванов Иван",
            "signatory_verb": "действующего",
            "authority_basis": "устав",
            "inn": "7701000001",
            "kpp": "770101001",
            "ogrn": "1027700000001",
            "legal_address": "Москва",
            "email": "a@b.ru",
            "bank_name": "Банк",
            "account": "40702810000000000001",
            "corr_account": "30101810000000000001",
            "bik": "044525225",
        }
    )


def _blank_requisite_columns() -> dict[str, str]:
    """Те же поля, без которых Word/PDF отвечают «Сначала заполните реквизиты договора»."""
    return {
        "legal_form": "   ",
        "full_name": "   ",
        "short_name": " \t ",
        "signatory_name": " ",
        "authority_basis": " ",
        "inn": " ",
        "ogrn": " ",
        "legal_address": " ",
        "email": " ",
        "bank_name": " ",
        "account": " ",
        "corr_account": " ",
        "bik": " ",
        "contract_date": " \t ",
    }


def test_search_marks_blank_card_incomplete_and_saved_requisites_complete(
    tmp_path: Path,
) -> None:
    service = _service(tmp_path, today=date(2026, 10, 20))
    user = _writer()
    empty = service.create_card(name="Импорт", number="0793/02/25", user=user)
    spaced = service.create_card(name="Пробелы", number="0844-06-25", user=user)
    filled = service.create_card(name="Ромашка", number="0001/10/26", user=user)
    assignments = ", ".join(f"{name} = ?" for name in _blank_requisite_columns())
    with _connect(service.repository.db_path) as conn:
        conn.execute(
            f"UPDATE supply_contract SET {assignments} WHERE id = ?",
            (*_blank_requisite_columns().values(), int(spaced["id"])),
        )
    service.save_requisites(int(filled["id"]), _buyer(), user=user)

    found = {
        row["number"]: row["requisites_complete"]
        for row in service.search_cards("", user=user)
    }
    opened_empty = service.get_card(int(empty["id"]), user=user)
    opened_filled = service.get_card(int(filled["id"]), user=user)

    assert found["0793/02/25"] is False
    assert found["0844-06-25"] is False
    assert found["0001/10/26"] is True
    assert opened_empty["requisites_complete"] is False
    assert opened_filled["requisites_complete"] is True
    assert opened_filled["number"] == "0001/10/26"


def test_update_missing_card_raises_not_found(tmp_path: Path) -> None:
    service = _service(tmp_path, today=date(2026, 10, 20))

    with pytest.raises(SupplyContractNotFoundError):
        service.update_card(404, number="1000/10/26", user=_writer())

    assert _count(service.repository.db_path, "supply_contract") == 0
    assert _events(service.repository.db_path) == []
