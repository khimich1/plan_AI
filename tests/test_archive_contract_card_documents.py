"""Сохранение бланка на ту же карточку и отказ Word/PDF/ЭДО без реквизитов (CDC-007)."""

from __future__ import annotations

import asyncio
from datetime import date
from io import BytesIO
from pathlib import Path

import pytest
from fastapi import HTTPException
from openpyxl import Workbook

from app.schemas.supply_contract import SupplyContractCreate
from app.services.supply_contract_service import (
    SupplyContractFieldError,
    SupplyContractNotFoundError,
    SupplyContractService,
    SupplyContractValidationError,
)
from core import kp_db_schema
from core.kp_db_common import _connect

_FILL_REQUISITES = "Сначала заполните реквизиты договора"

_REQUISITE_VALUES = {
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


def _service(tmp_path: Path) -> SupplyContractService:
    db_path = str(tmp_path / "cards.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return SupplyContractService(db_path=db_path, today=date(2026, 10, 20))


def _writer(username: str = "ivanov") -> dict:
    return {"id": 7, "role": "manager", "username": username}


def _count(db_path: str, table: str) -> int:
    with _connect(db_path) as conn:
        row = conn.execute(f"SELECT COUNT(*) FROM {table}").fetchone()
    return int(row[0])


def _buyer(**overrides: object) -> SupplyContractCreate:
    data: dict[str, object] = {
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
    data.update(overrides)
    return SupplyContractCreate.model_validate(data)


def _set_columns(db_path: str, contract_id: int, **columns: object) -> None:
    assignments = ", ".join(f"{name} = ?" for name in columns)
    with _connect(db_path) as conn:
        conn.execute(
            f"UPDATE supply_contract SET {assignments} WHERE id = ?",
            (*columns.values(), int(contract_id)),
        )


def _bind_offer(db_path: str, contract_id: int) -> None:
    with _connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO counterparties (code_1c, name, name_normalized, inn)
            VALUES ('00-0001', 'ООО Ромашка', 'ооо ромашка', '7701000001')
            """
        )
        client_id = int(conn.execute("SELECT last_insert_rowid()").fetchone()[0])
        conn.execute(
            """
            INSERT INTO KP_offers (
                kp_id, creation_date, customer_name, manager_name, counterparty_id
            ) VALUES (1, '2026-10-01', 'РОМАШКА ООО', 'Иван Иванов', ?)
            """,
            (client_id,),
        )
        conn.execute(
            "INSERT INTO kp_meta (kp_id, status, owner_user_id) VALUES (1, 'в архиве', 7)"
        )
        conn.execute(
            "UPDATE supply_contract SET counterparty_id = ? WHERE id = ?",
            (client_id, int(contract_id)),
        )


def _card_xlsx() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None
    sheet["A1"] = "ИНН 7604010011"
    sheet["A2"] = "Р/с 40702810000000000007"
    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def _assert_exact_refusal(caught: pytest.ExceptionInfo[SupplyContractValidationError]) -> None:
    assert str(caught.value) == _FILL_REQUISITES


@pytest.mark.parametrize("file_format", ["docx", "pdf"])
def test_empty_card_refuses_word_pdf_and_edo(tmp_path: Path, file_format: str) -> None:
    service = _service(tmp_path)
    created = service.create_card(name="ИП Секлюцкий", number="0786/1/25", user=_writer())

    with pytest.raises(SupplyContractValidationError) as word:
        service.document_for_contract(
            int(created["id"]),
            user=_writer(),
            file_format=file_format,
        )
    with pytest.raises(SupplyContractValidationError) as edo:
        service.edo_for_contract(
            int(created["id"]),
            user=_writer(),
            file_format=file_format,
        )

    _assert_exact_refusal(word)
    _assert_exact_refusal(edo)
    assert _count(service.repository.db_path, "supply_contract") == 1


def test_whitespace_requisites_count_as_empty(tmp_path: Path) -> None:
    service = _service(tmp_path)
    created = service.create_card(name="ИП Секлюцкий", number="0786/1/25", user=_writer())
    blanks = {name: "   " for name in _REQUISITE_VALUES}
    blanks["contract_date"] = " \t "
    _set_columns(service.repository.db_path, int(created["id"]), **blanks)

    with pytest.raises(SupplyContractValidationError) as caught:
        service.document_for_contract(int(created["id"]), user=_writer())

    _assert_exact_refusal(caught)


def test_null_contract_date_refuses_download_even_with_other_fields(tmp_path: Path) -> None:
    service = _service(tmp_path)
    created = service.create_card(name="Без даты", number="ТБ-1", user=_writer())
    _set_columns(service.repository.db_path, int(created["id"]), **_REQUISITE_VALUES)

    with pytest.raises(SupplyContractValidationError) as caught:
        service.document_for_contract(int(created["id"]), user=_writer(), file_format="pdf")

    _assert_exact_refusal(caught)
    loaded = service.repository.get_by_id(int(created["id"]))
    assert loaded is not None
    assert loaded["contract_date"] is None
    assert loaded["legal_form"] == "ooo"


def test_save_requisites_keeps_the_same_row_number_and_date(tmp_path: Path) -> None:
    service = _service(tmp_path)
    user = _writer()
    created = service.create_card(name="ИП Секлюцкий", number="0786/1/25", user=user)
    contract_id = int(created["id"])
    parties_before = _count(service.repository.db_path, "counterparties")

    saved = service.save_requisites(contract_id, _buyer(), user=user)
    again = service.save_requisites(
        contract_id,
        _buyer(full_name="ООО Ромашка обновлённая"),
        user=user,
    )
    payload, filename = service.document_for_contract(contract_id, user=user)
    edo, edo_name = service.edo_for_contract(contract_id, user=user)

    assert saved["id"] == again["id"] == contract_id
    assert saved["number"] == again["number"] == "0786/1/25"
    assert saved["contract_date"] == "2025-01-01"
    assert again["contract_date"] == "2025-01-01"
    assert again["full_name"] == "ООО Ромашка обновлённая"
    assert again["imported_name"] == "ИП Секлюцкий"
    assert again["counterparty_id"] is None
    assert again["manager_name"] == "ivanov"
    assert payload.startswith(b"PK")
    assert filename == "dogovor-0786-1-25.docx"
    assert edo.startswith(b"PK")
    assert edo_name == "soglashenie-edo-0786-1-25.docx"
    assert _count(service.repository.db_path, "supply_contract") == 1
    assert _count(service.repository.db_path, "counterparties") == parties_before
    loaded = service.repository.get_by_number("0786/1/25")
    assert loaded is not None
    assert loaded["id"] == contract_id
    assert loaded["contract_date"] == "2025-01-01"
    assert loaded["email"] == "a@b.ru"


def test_save_requisites_changes_date_only_when_the_body_has_one(tmp_path: Path) -> None:
    service = _service(tmp_path)
    user = _writer()
    created = service.create_card(name="ИП Секлюцкий", number="0786/1/25", user=user)

    updated = service.save_requisites(
        int(created["id"]),
        _buyer(contract_date=date(2026, 4, 15)),
        user=user,
    )

    assert updated["id"] == created["id"]
    assert updated["number"] == "0786/1/25"
    assert updated["contract_date"] == "2026-04-15"
    assert _count(service.repository.db_path, "supply_contract") == 1


def test_blank_manager_becomes_username_and_filled_manager_stays(tmp_path: Path) -> None:
    service = _service(tmp_path)
    user = _writer("petrov")
    blank = service.create_card(name="Пустой менеджер", number="0786/1/25", user=user)
    spaced = service.create_card(name="Пробелы", number="0844-06-25", user=user)
    named = service.create_card(name="Уже есть", number="0001/10/26", user=user)
    _set_columns(service.repository.db_path, int(spaced["id"]), manager_name="   ")
    _set_columns(service.repository.db_path, int(named["id"]), manager_name="Пургина")

    from_null = service.save_requisites(int(blank["id"]), _buyer(), user=user)
    from_spaces = service.save_requisites(int(spaced["id"]), _buyer(), user=user)
    kept = service.save_requisites(int(named["id"]), _buyer(), user=user)

    assert from_null["manager_name"] == "petrov"
    assert from_spaces["manager_name"] == "petrov"
    assert kept["manager_name"] == "Пургина"
    assert kept["number"] == "0001/10/26"
    assert _count(service.repository.db_path, "supply_contract") == 3


def test_invalid_requisites_do_not_insert_a_second_contract(tmp_path: Path) -> None:
    service = _service(tmp_path)
    created = service.create_card(name="ИП Секлюцкий", number="0786/1/25", user=_writer())

    with pytest.raises(SupplyContractFieldError, match="КПП"):
        service.save_requisites(int(created["id"]), _buyer(kpp=None), user=_writer())

    loaded = service.repository.get_by_id(int(created["id"]))
    assert loaded is not None
    assert loaded["legal_form"] is None
    assert loaded["manager_name"] is None
    assert loaded["number"] == "0786/1/25"
    assert _count(service.repository.db_path, "supply_contract") == 1


def test_filled_0001_still_passes_contract_and_kp_gate(tmp_path: Path) -> None:
    service = _service(tmp_path)
    created = service.create_card(name="ООО Ромашка", number="0001/10/26", user=_writer())
    _set_columns(service.repository.db_path, int(created["id"]), **_REQUISITE_VALUES)
    _bind_offer(service.repository.db_path, int(created["id"]))

    payload, filename = service.document_for_contract(int(created["id"]), user=_writer())
    via_kp, kp_name = service.document_for_kp(1, user=_writer())
    edo, edo_name = service.edo_for_contract(int(created["id"]), user=_writer())

    assert payload.startswith(b"PK")
    assert via_kp.startswith(b"PK")
    assert edo.startswith(b"PK")
    assert filename == "dogovor-0001-10-26.docx"
    assert kp_name == "dogovor-0001-10-26.docx"
    assert edo_name == "soglashenie-edo-0001-10-26.docx"
    loaded = service.repository.get_by_number("0001/10/26")
    assert loaded is not None
    assert loaded["contract_date"] == "2026-10-01"
    assert loaded["full_name"] == "ООО Ромашка"
    assert _count(service.repository.db_path, "supply_contract") == 1


def test_kp_downloads_refuse_card_without_requisites(tmp_path: Path) -> None:
    service = _service(tmp_path)
    created = service.create_card(name="ИП Секлюцкий", number="0786/1/25", user=_writer())
    _bind_offer(service.repository.db_path, int(created["id"]))

    with pytest.raises(SupplyContractValidationError) as word:
        service.document_for_kp(1, user=_writer(), file_format="docx")
    with pytest.raises(SupplyContractValidationError) as pdf:
        service.document_for_kp(1, user=_writer(), file_format="pdf")
    with pytest.raises(SupplyContractValidationError) as edo:
        service.edo_agreement_for_kp(1, user=_writer(), file_format="pdf")

    _assert_exact_refusal(word)
    _assert_exact_refusal(pdf)
    _assert_exact_refusal(edo)


def test_missing_email_is_an_empty_blank_not_the_edo_buyer_message(tmp_path: Path) -> None:
    service = _service(tmp_path)
    created = service.create_card(name="ООО Ромашка", number="0001/10/26", user=_writer())
    filled = dict(_REQUISITE_VALUES)
    filled["email"] = None
    _set_columns(service.repository.db_path, int(created["id"]), **filled)

    with pytest.raises(SupplyContractValidationError) as caught:
        service.edo_for_contract(int(created["id"]), user=_writer())

    _assert_exact_refusal(caught)


def test_parse_for_contract_does_not_need_kp_or_counterparty(tmp_path: Path) -> None:
    service = _service(tmp_path)
    created = service.create_card(name="ИП Секлюцкий", number="0786/1/25", user=_writer())

    parsed = asyncio.run(
        service.parse_for_contract(
            int(created["id"]),
            _card_xlsx(),
            filename="card.xlsx",
            user=_writer(),
        )
    )

    assert parsed["fields"]["inn"] == "7604010011"
    assert parsed["fields"]["account"] == "40702810000000000007"
    loaded = service.repository.get_by_id(int(created["id"]))
    assert loaded is not None
    assert loaded["counterparty_id"] is None
    assert loaded["legal_form"] is None
    assert _count(service.repository.db_path, "supply_contract") == 1
    assert _count(service.repository.db_path, "counterparties") == 0
    with _connect(service.repository.db_path) as conn:
        offers = conn.execute("SELECT COUNT(*) FROM KP_offers").fetchone()
    assert int(offers[0]) == 0


def test_missing_card_is_not_found_for_blank_save_and_download(tmp_path: Path) -> None:
    service = _service(tmp_path)
    user = _writer()

    with pytest.raises(SupplyContractNotFoundError):
        service.save_requisites(404, _buyer(), user=user)
    with pytest.raises(SupplyContractNotFoundError):
        service.document_for_contract(404, user=user)
    with pytest.raises(SupplyContractNotFoundError):
        asyncio.run(
            service.parse_for_contract(404, _card_xlsx(), filename="card.xlsx", user=user)
        )

    assert _count(service.repository.db_path, "supply_contract") == 0


def test_other_role_cannot_save_or_download_the_blank(tmp_path: Path) -> None:
    service = _service(tmp_path)
    created = service.create_card(name="ИП Секлюцкий", number="0786/1/25", user=_writer())
    outsider = {"id": 4, "role": "production", "username": "shop"}

    with pytest.raises(HTTPException) as denied_save:
        service.save_requisites(int(created["id"]), _buyer(), user=outsider)
    with pytest.raises(HTTPException) as denied_word:
        service.document_for_contract(int(created["id"]), user=outsider)

    assert denied_save.value.status_code == 403
    assert denied_word.value.status_code == 403
    loaded = service.repository.get_by_id(int(created["id"]))
    assert loaded is not None
    assert loaded["legal_form"] is None
    assert _count(service.repository.db_path, "supply_contract") == 1
