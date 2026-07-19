"""Tests for ``init_default_managers`` loading from JSON seed."""

from __future__ import annotations

import json

import pytest

from core import kp_db
from tests.helpers import kp_db_fixtures as fx


@pytest.fixture
def iso_db(tmp_path):
    return fx.make_iso_db(tmp_path)


def test_init_default_managers_from_seed_file(iso_db: str, tmp_path) -> None:
    seed = tmp_path / "managers.json"
    seed.write_text(
        json.dumps(
            [
                {
                    "fio": "Тест Менеджер",
                    "contact_number": "79990001122",
                    "email": "test.manager@example.local",
                }
            ]
        ),
        encoding="utf-8",
    )

    added = kp_db.init_default_managers(iso_db, seed_path=str(seed))

    assert added == 1
    managers = kp_db.get_all_managers(iso_db)
    assert len(managers) == 1
    assert managers[0]["email"] == "test.manager@example.local"


def test_init_default_managers_skips_duplicate_email(iso_db: str, tmp_path) -> None:
    seed = tmp_path / "managers.json"
    seed.write_text(
        json.dumps(
            [
                {
                    "fio": "Тест Менеджер",
                    "contact_number": "79990001122",
                    "email": "dup@example.local",
                }
            ]
        ),
        encoding="utf-8",
    )

    assert kp_db.init_default_managers(iso_db, seed_path=str(seed)) == 1
    assert kp_db.init_default_managers(iso_db, seed_path=str(seed)) == 0
    assert len(kp_db.get_all_managers(iso_db)) == 1


def test_init_default_managers_missing_seed_returns_zero(iso_db: str, tmp_path) -> None:
    missing = tmp_path / "no_seed.json"
    assert kp_db.init_default_managers(iso_db, seed_path=str(missing)) == 0
