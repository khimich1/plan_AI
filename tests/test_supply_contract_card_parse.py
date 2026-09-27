"""Разбор карточки покупателя: текст и снимок (SC-006, SC-007)."""

from __future__ import annotations

import sqlite3
from io import BytesIO
from pathlib import Path

from docx import Document

import pytest
from openpyxl import Workbook

from app.services.supply_contract_service import SupplyContractService, SupplyContractValidationError
from core import kp_db_schema
from core.kp_db_common import _connect
from core.supply_contract_card_parse import parse_card_text

ADMIN = {"id": 7, "role": "admin", "username": "admin"}


def _fresh_db(tmp_path: Path) -> str:
    db_path = str(tmp_path / "supply.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return db_path


def _seed(db_path: str) -> None:
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        cur = conn.execute(
            """
            INSERT INTO counterparties (code_1c, name, name_normalized, inn, kpp)
            VALUES ('00-1', 'РОМАШКА ООО', 'ромашка ооо', '7701000001', '770101001')
            """
        )
        client_id = int(cur.lastrowid)
        conn.execute(
            """
            INSERT INTO KP_offers (
                kp_id, creation_date, customer_name, manager_name, counterparty_id
            ) VALUES (1, '2026-09-01', 'РОМАШКА ООО', 'Иван Иванов', ?)
            """,
            (client_id,),
        )
        conn.execute(
            "INSERT INTO kp_meta (kp_id, status, owner_user_id) VALUES (1, 'в архиве', 7)"
        )


def _card_docx() -> bytes:
    document = Document()
    document.add_paragraph("ООО «Ромашка»")
    document.add_paragraph("Общество с ограниченной ответственностью «Ромашка»")
    document.add_paragraph("ИНН 760 4 010011")
    document.add_paragraph("КПП 760401001")
    document.add_paragraph("ОГРН 1027600000009")
    document.add_paragraph("Генеральный директор Иванов Иван Иванович")
    document.add_paragraph("на основании Устава")
    document.add_paragraph("Юридический адрес: г. Ярославль, ул. Ленина, 1")
    document.add_paragraph("Р/с 40702810000000000007")
    document.add_paragraph("Банк: ПАО «Тест Банк»")
    document.add_paragraph("К/с 30101810000000000760")
    document.add_paragraph("БИК 044525225")
    document.add_paragraph("E-mail: is-ag @mail.ru")
    buffer = BytesIO()
    document.save(buffer)
    return buffer.getvalue()


def _card_pdf() -> bytes:
    """Небольшой PDF с текстовым слоем, не карточка клиента."""
    text = (
        "ООО «Ромашка»\\n"
        "ИНН 760 4 010011\\n"
        "Р/с 40702810000000000007\\n"
        "Генеральный директор Иванов Иван Иванович"
    )
    stream = f"BT ({text}) Tj ET".encode()
    return b"%PDF-1.4\n1 0 obj\n<< /Length " + str(len(stream)).encode() + (
        b" >>\nstream\n" + stream + b"\nendstream\nendobj\n%%EOF"
    )


def _count(db_path: str) -> int:
    with _connect(db_path) as conn:
        return int(conn.execute("SELECT COUNT(*) FROM supply_contract").fetchone()[0])


class _BoomVision:
    async def __call__(self, data: bytes, *, filename: str) -> dict:
        raise AssertionError("зрение не должно вызываться для текстового файла")


def test_checksum_rejects_bad_inn_and_keeps_name_in_doubtful() -> None:
    result = parse_card_text("ИНН 7604010010\nКПП 760401001")

    assert result.fields["inn"] is None
    assert "inn" in result.doubtful


def test_checksum_drops_kpp_and_corr_account_of_wrong_length() -> None:
    result = parse_card_text("КПП 76040100\nК/с 3010181000000000076")

    assert result.fields["kpp"] is None
    assert result.fields["corr_account"] is None
    assert "kpp" in result.doubtful
    assert "corr_account" in result.doubtful


def test_checksum_does_not_promote_18_digit_row_to_account() -> None:
    result = parse_card_text("Р/с 407028105771000149")

    assert result.fields["account"] is None
    assert "account" in result.doubtful


def test_checksum_keeps_20_digit_account_without_bik_and_marks_doubtful() -> None:
    result = parse_card_text("Р/с 40702810000000000007")

    assert result.fields["account"] == "40702810000000000007"
    assert "account" in result.doubtful


def test_checksum_drops_account_when_accepted_bik_key_fails() -> None:
    result = parse_card_text("БИК 044525225\nР/с 40702810000000000001")

    assert result.fields["bik"] == "044525225"
    assert result.fields["account"] is None
    assert "account" in result.doubtful


def test_layout_slash_does_not_put_inn_into_kpp() -> None:
    result = parse_card_text("ИНН/КПП 7604010011/760401001")

    assert result.fields["inn"] == "7604010011"
    assert result.fields["kpp"] == "760401001"


def test_layout_reads_ogrn_from_header_and_from_its_own_line() -> None:
    header = parse_card_text("ИНН 7604010011 КПП 760401001 ОГРН 1027600000009")
    line = parse_card_text("ОГРН 1027600000009")

    assert header.fields["ogrn"] == "1027600000009"
    assert header.fields["inn"] == "7604010011"
    assert line.fields["ogrn"] == "1027600000009"


def test_layout_bank_branch_without_the_word_bank() -> None:
    result = parse_card_text("Ярославское отделение № 17")

    assert result.fields["bank_name"] == "Ярославское отделение № 17"


def test_layout_kfh_head_fills_position_and_name() -> None:
    result = parse_card_text("Глава КФХ Петров Пётр Петрович")

    assert result.fields["signatory_position"] == "Глава КФХ"
    assert result.fields["signatory_name"] == "Петров Пётр Петрович"


def test_layout_ip_fills_position_and_name() -> None:
    result = parse_card_text("Индивидуальный предприниматель Сидоров Сидор Сидорович")

    assert result.fields["signatory_position"] == "Индивидуальный предприниматель"
    assert result.fields["signatory_name"] == "Сидоров Сидор Сидорович"


def test_layout_authority_follows_the_chosen_signatory() -> None:
    result = parse_card_text(
        "Генеральный директор Иванов Иван Иванович, действующего на основании Устава.\n"
        "Представитель Петров Пётр Петрович действует на основании доверенности № 15 от 01.02.2024"
    )

    assert result.fields["signatory_name"] == "Иванов Иван Иванович"
    assert result.fields["authority_basis"] == "устав"
    assert result.fields["poa_number"] is None


def test_layout_two_accounts_stay_in_the_list_and_leave_the_field_empty() -> None:
    result = parse_card_text("Р/с 40702810000000000007\nР/с 40802810000000000015")

    assert result.as_dict()["accounts"] == ["40702810000000000007", "40802810000000000015"]
    assert result.fields["account"] in (None, "")


def _card_xlsx() -> bytes:
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "ИНН 7604010011"
    sheet["A2"] = "Р/с 40702810000000000007"
    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def test_xlsx_card_fills_inn_and_account_without_vision(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    _seed(db_path)
    service = SupplyContractService(db_path=db_path, card_recognizer=_BoomVision())

    result = __import__("asyncio").run(
        service.parse_for_kp(1, _card_xlsx(), filename="card.xlsx", user=ADMIN)
    )

    assert result["fields"]["inn"] == "7604010011"
    assert result["fields"]["account"] == "40702810000000000007"
    assert _count(db_path) == 0


def test_xlsx_legacy_xls_is_not_accepted(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    _seed(db_path)
    service = SupplyContractService(db_path=db_path, card_recognizer=_BoomVision())

    with pytest.raises(SupplyContractValidationError, match="Неподдерживаемый формат"):
        __import__("asyncio").run(
            service.parse_for_kp(1, b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1xls", filename="card.xls", user=ADMIN)
        )

    assert _count(db_path) == 0


def test_docx_card_glues_inn_and_fills_account_and_director(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    _seed(db_path)
    service = SupplyContractService(db_path=db_path, card_recognizer=_BoomVision())

    result = __import__("asyncio").run(
        service.parse_for_kp(1, _card_docx(), filename="card.docx", user=ADMIN)
    )

    assert result["fields"]["inn"] == "7604010011"
    assert result["fields"]["account"] == "40702810000000000007"
    assert result["fields"]["signatory_name"] == "Иванов Иван Иванович"
    assert result["fields"]["email"] == "is-ag@mail.ru"
    assert "inn" not in result["doubtful"]
    assert _count(db_path) == 0


def test_text_pdf_does_not_call_vision(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    _seed(db_path)
    service = SupplyContractService(db_path=db_path, card_recognizer=_BoomVision())

    result = __import__("asyncio").run(
        service.parse_for_kp(1, _card_pdf(), filename="card.pdf", user=ADMIN)
    )

    assert result["fields"]["inn"] == "7604010011"
    assert result["verify_failed"] is False
    assert _count(db_path) == 0


def test_not_a_card_rejects_text_without_requisites(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    _seed(db_path)
    document = Document()
    document.add_paragraph("Паспорт гражданина")
    buffer = BytesIO()
    document.save(buffer)
    service = SupplyContractService(db_path=db_path, card_recognizer=_BoomVision())

    with pytest.raises(SupplyContractValidationError, match="это не карточка контрагента"):
        __import__("asyncio").run(
            service.parse_for_kp(1, buffer.getvalue(), filename="passport.docx", user=ADMIN)
        )

    assert _count(db_path) == 0


def test_not_a_card_keeps_upload_that_has_an_inn(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    _seed(db_path)
    service = SupplyContractService(db_path=db_path, card_recognizer=_BoomVision())

    result = __import__("asyncio").run(
        service.parse_for_kp(1, _card_docx(), filename="card.docx", user=ADMIN)
    )

    assert result["fields"]["inn"] == "7604010011"
    assert _count(db_path) == 0


def test_next_line_reads_inn_and_address_not_the_parentheses() -> None:
    result = parse_card_text(
        "Юридический адрес (в соответствии с учредительными документами)\n"
        "г. Ярославль, ул. Ленина, д. 1\n"
        "ИНН\n"
        "7604010011"
    )

    assert result.fields["inn"] == "7604010011"
    assert result.fields["legal_address"] == "г. Ярославль, ул. Ленина, д. 1"


def test_next_line_bik_bank_label_and_ogrn_with_date() -> None:
    result = parse_card_text(
        "БИК банка:\n044525225\nОГРН 1027600000009 от 15.03.2002"
    )

    assert result.fields["bik"] == "044525225"
    assert result.fields["ogrn"] == "1027600000009"


def test_suspicious_parenthetical_address_is_not_a_field() -> None:
    result = parse_card_text(
        "Юридический адрес: (в соответствии с учредительными документами)"
    )

    assert result.fields["legal_address"] is None
    assert "legal_address" in result.doubtful


def test_suspicious_address_fragment_is_not_a_field() -> None:
    result = parse_card_text("Юридический адрес организации")

    assert result.fields["legal_address"] is None
    assert "legal_address" in result.doubtful


def test_suspicious_bank_line_with_account_is_not_a_field() -> None:
    result = parse_card_text("Банк: ПАО «Сбер» р/с 40702810000000000007")

    assert result.fields["bank_name"] is None
    assert "bank_name" in result.doubtful


def test_suspicious_email_glued_to_phone_is_not_a_field() -> None:
    glued = parse_card_text("E-mail: 89001234567info@mail.ru")
    clean = parse_card_text("E-mail: is-ag@mail.ru")

    assert glued.fields["email"] is None
    assert "email" in glued.doubtful
    assert clean.fields["email"] == "is-ag@mail.ru"
    assert "email" not in clean.doubtful


def test_bundle_two_banks_leave_four_fields_empty() -> None:
    result = parse_card_text(
        "Банк: Альфа\n"
        "Р/с 40702810000000000007\n"
        "К/с 30101810000000000760\n"
        "БИК 044525225\n"
        "Банк: ВТБ\n"
        "Р/с 40702810900000000009\n"
        "К/с 30101810400000000225\n"
        "БИК 044525974"
    )

    assert len(result.banks) == 2
    assert result.banks[0]["account"] == "40702810000000000007"
    assert result.banks[0]["bik"] == "044525225"
    assert result.banks[1]["bank_name"] == "ВТБ"
    assert result.banks[1]["account"] == "40702810900000000009"
    assert result.fields["account"] is None
    assert result.fields["bik"] is None
    assert result.fields["corr_account"] is None
    assert result.fields["bank_name"] is None


def test_bundle_one_bik_for_two_accounts_does_not_pick() -> None:
    result = parse_card_text(
        "Р/с 40702810000000000007\n"
        "Р/с 40702810000000000010\n"
        "БИК 044525225"
    )

    assert result.banks == []
    assert result.fields["account"] is None
    assert result.accounts == ["40702810000000000007", "40702810000000000010"]


def test_bundle_two_corr_accounts_are_not_written_into_the_field() -> None:
    result = parse_card_text("К/с 30101810000000000760\nК/с 30101810400000000225")

    assert result.fields["corr_account"] is None


def test_two_inn_values_stay_doubtful() -> None:
    result = parse_card_text("ИНН 7604010011\nИНН 7707083893")

    assert result.fields["inn"] in {"7604010011", "7707083893"}
    assert "inn" in result.doubtful


def test_source_text_is_trimmed_to_twenty_thousand_chars() -> None:
    body = "ИНН 7604010011\n" + ("Я" * 25000)
    result = parse_card_text(body)

    assert result.source_text.startswith("ИНН 7604010011")
    assert len(result.source_text) == 20000
    assert result.as_dict()["source_text"] == result.source_text


def test_empty_docx_does_not_create_contract(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    _seed(db_path)
    document = Document()
    document.add_paragraph("")
    buffer = BytesIO()
    document.save(buffer)
    service = SupplyContractService(db_path=db_path, card_recognizer=_BoomVision())

    with pytest.raises(SupplyContractValidationError, match="это не карточка контрагента"):
        __import__("asyncio").run(
            service.parse_for_kp(1, buffer.getvalue(), filename="empty.docx", user=ADMIN)
        )

    assert _count(db_path) == 0

