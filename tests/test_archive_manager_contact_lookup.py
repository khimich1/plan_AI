"""Поиск контактов менеджера для шапки архива: без id и без заглушки id=1."""

from __future__ import annotations

import sqlite3
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from app.services.manager_contact_lookup import (
    header_contacts_for_archive,
    load_managers_if_db_exists,
    lookup_manager_contacts,
)

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


def test_lookup_matches_fio_casefold_exact() -> None:
    found = lookup_manager_contacts("  зубов алексей юрьевич  ", MANAGERS)

    assert found == ("79621872265", "zubov@gbkstart.ru")
    assert "id" not in found

    earlier = {
        "id": 9,
        "fio": "зубов алексей юрьевич",
        "contact_number": "70000000000",
        "email": "first@example.ru",
    }
    assert lookup_manager_contacts(
        "Зубов Алексей Юрьевич",
        [earlier, ZUBOV],
    ) == ("70000000000", "first@example.ru")
    assert lookup_manager_contacts("Зубов  Алексей Юрьевич", MANAGERS) == ("", "")


def test_lookup_unmatched_name_does_not_return_id1_contacts() -> None:
    assert lookup_manager_contacts("Чужое ФИО", MANAGERS) == ("", "")
    assert lookup_manager_contacts("Чужое ФИО", MANAGERS) != (
        ALEXANDER["contact_number"],
        ALEXANDER["email"],
    )


def test_lookup_blank_name_returns_empty_pair() -> None:
    assert lookup_manager_contacts(None, MANAGERS) == ("", "")
    assert lookup_manager_contacts("", MANAGERS) == ("", "")
    assert lookup_manager_contacts("   ", MANAGERS) == ("", "")


def test_lookup_partial_card_returns_only_filled_field() -> None:
    phone_only = {
        "id": 2,
        "fio": "Зубов Алексей Юрьевич",
        "contact_number": "79621872265",
        "email": "",
    }
    email_only = {
        "id": 2,
        "fio": "Зубов Алексей Юрьевич",
        "contact_number": None,
        "email": "  zubov@gbkstart.ru  ",
    }

    assert lookup_manager_contacts("Зубов Алексей Юрьевич", [ALEXANDER, phone_only]) == (
        "79621872265",
        "",
    )
    assert lookup_manager_contacts("Зубов Алексей Юрьевич", [ALEXANDER, email_only]) == (
        "",
        "zubov@gbkstart.ru",
    )


def test_header_contacts_suppress_only_when_name_present() -> None:
    assert header_contacts_for_archive(None, MANAGERS) == (None, None, False)
    assert header_contacts_for_archive("", MANAGERS) == (None, None, False)
    assert header_contacts_for_archive("   ", MANAGERS) == (None, None, False)

    assert header_contacts_for_archive("Нет такого", MANAGERS) == (None, None, True)
    assert header_contacts_for_archive("Зубов Алексей Юрьевич", MANAGERS) == (
        "79621872265",
        "zubov@gbkstart.ru",
        True,
    )

    empty_card = {
        "id": 2,
        "fio": "Зубов Алексей Юрьевич",
        "contact_number": "   ",
        "email": "",
    }
    assert header_contacts_for_archive("Зубов Алексей Юрьевич", [empty_card]) == (
        None,
        None,
        True,
    )
    phone_only = {
        "id": 2,
        "fio": "Зубов Алексей Юрьевич",
        "contact_number": "79621872265",
        "email": None,
    }
    assert header_contacts_for_archive("Зубов Алексей Юрьевич", [phone_only]) == (
        "79621872265",
        None,
        True,
    )


def test_load_managers_missing_file_returns_empty_and_does_not_create_db(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def _forbid_connect(*_args: object, **_kwargs: object) -> None:
        raise AssertionError("sqlite connect must not run when the database file is missing")

    monkeypatch.setattr(sqlite3, "connect", _forbid_connect)
    import core.kp_db_common as kp_common

    monkeypatch.setattr(kp_common.sqlite3, "connect", _forbid_connect)

    missing = tmp_path / "plita.db"
    assert load_managers_if_db_exists(missing) == []
    assert load_managers_if_db_exists(str(missing)) == []
    assert load_managers_if_db_exists(MagicMock()) == []
    assert load_managers_if_db_exists(None) == []
    assert missing.exists() is False
