"""Архив передаёт в PDF и XLSX контакты менеджера, найденные по ФИО."""

from __future__ import annotations

import asyncio
from pathlib import Path
from typing import Any
from unittest.mock import MagicMock

import pytest

from tests.test_archive_service import ADMIN, _make_raw, _make_service

ALEXANDER = {
    "id": 1,
    "fio": "Александр",
    "contact_number": "79000000001",
    "email": "alexander@example.ru",
}
ZUBOV = {
    "id": 2,
    "fio": "Зубов Алексей Юрьевич",
    "contact_number": "79621872265",
    "email": "zubov@gbkstart.ru",
}
MANAGERS = [ALEXANDER, ZUBOV]


class _FakeBuffer:
    def getvalue(self) -> bytes:
        return b"fake-document"


def _install_managers(
    monkeypatch: pytest.MonkeyPatch,
    managers: list[dict[str, Any]],
) -> None:
    monkeypatch.setattr(
        "app.services.archive_service.load_managers_if_db_exists",
        lambda _db_path: list(managers),
    )


def _install_generator(
    monkeypatch: pytest.MonkeyPatch,
    name: str,
) -> MagicMock:
    fake = MagicMock(return_value=_FakeBuffer())
    monkeypatch.setattr(f"app.services.archive_service.{name}", fake)
    return fake


def _contacts(call_kwargs: dict[str, Any]) -> tuple[Any, Any, Any]:
    return (
        call_kwargs["manager_phone"],
        call_kwargs["manager_email"],
        call_kwargs["suppress_company_contacts"],
    )


def test_archive_pdf_passes_zubov_contacts_and_suppress_flag(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw(
        manager_name="Зубов Алексей Юрьевич",
    )
    service = _make_service(repository, tmp_path)
    _install_managers(monkeypatch, MANAGERS)
    fake_pdf = _install_generator(monkeypatch, "generate_commercial_offer_pdf")

    path = asyncio.run(service.generate_document(42, "pdf", user=ADMIN))

    assert path.name == "КП_42.pdf"
    fake_pdf.assert_called_once()
    kwargs = fake_pdf.call_args.kwargs
    assert kwargs["manager_name"] == "Зубов Алексей Юрьевич"
    assert _contacts(kwargs) == ("79621872265", "zubov@gbkstart.ru", True)


def test_archive_xlsx_and_delivery_in_unit_share_header_contacts(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw(
        manager_name="Зубов Алексей Юрьевич",
    )
    service = _make_service(repository, tmp_path)
    _install_managers(monkeypatch, MANAGERS)
    fake_xlsx = _install_generator(monkeypatch, "generate_commercial_offer_xlsx")

    xlsx_path = asyncio.run(service.generate_document(42, "xlsx", user=ADMIN))
    delivery_path = asyncio.run(
        service.generate_document(42, "xlsx_delivery_in_unit", user=ADMIN)
    )

    assert xlsx_path.name == "КП_42.xlsx"
    assert delivery_path.name == "КП_42_с_доставкой_в_цене.xlsx"
    assert fake_xlsx.call_count == 2
    plain = fake_xlsx.call_args_list[0].kwargs
    delivery = fake_xlsx.call_args_list[1].kwargs
    expected = ("79621872265", "zubov@gbkstart.ru", True)
    assert _contacts(plain) == expected
    assert _contacts(delivery) == expected
    assert plain["embed_delivery_in_unit_price"] is False
    assert delivery["embed_delivery_in_unit_price"] is True


def test_archive_unmatched_fio_sets_suppress_and_empty_contacts(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw(manager_name="Петров Пётр")
    service = _make_service(repository, tmp_path)
    _install_managers(monkeypatch, MANAGERS)
    fake_pdf = _install_generator(monkeypatch, "generate_commercial_offer_pdf")

    asyncio.run(service.generate_document(42, "pdf", user=ADMIN))

    kwargs = fake_pdf.call_args.kwargs
    assert kwargs["manager_name"] == "Петров Пётр"
    assert _contacts(kwargs) == (None, None, True)
    assert kwargs["manager_phone"] != ALEXANDER["contact_number"]
    assert kwargs["manager_email"] != ALEXANDER["email"]


def test_archive_blank_manager_name_keeps_company_fallback(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _install_managers(monkeypatch, MANAGERS)
    fake_pdf = _install_generator(monkeypatch, "generate_commercial_offer_pdf")

    for blank in (None, "", "   "):
        repository = MagicMock()
        repository.get_by_id.return_value = _make_raw(manager_name=blank)
        service = _make_service(repository, tmp_path)
        fake_pdf.reset_mock()

        asyncio.run(service.generate_document(42, "pdf", user=ADMIN))

        kwargs = fake_pdf.call_args.kwargs
        assert kwargs["manager_name"] == blank
        assert _contacts(kwargs) == (None, None, False)


def test_archive_partial_phone_passes_phone_and_suppress(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    phone_only = {
        "id": 2,
        "fio": "Зубов Алексей Юрьевич",
        "contact_number": "79621872265",
        "email": "",
    }
    repository = MagicMock()
    repository.get_by_id.return_value = _make_raw(
        manager_name="Зубов Алексей Юрьевич",
    )
    service = _make_service(repository, tmp_path)
    _install_managers(monkeypatch, [ALEXANDER, phone_only])
    fake_pdf = _install_generator(monkeypatch, "generate_commercial_offer_pdf")
    fake_xlsx = _install_generator(monkeypatch, "generate_commercial_offer_xlsx")

    asyncio.run(service.generate_document(42, "pdf", user=ADMIN))
    asyncio.run(service.generate_document(42, "xlsx", user=ADMIN))

    assert _contacts(fake_pdf.call_args.kwargs) == ("79621872265", None, True)
    assert _contacts(fake_xlsx.call_args.kwargs) == ("79621872265", None, True)
