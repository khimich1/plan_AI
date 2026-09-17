"""GPS-006: POST /api/v1/nomenclature/import-1c."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from core.nomenclature_guid import ensure_schema, upsert
from core.pricelist_1c_parser import PricelistRow
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.production_api_fixtures import VALID_APP_SECRET_KEY, session_cookie

API = "/api/v1/nomenclature/import-1c"
PLAIN_GUID = "11111111-1111-1111-1111-111111111111"
MARK = "С30.30-3"
PILE_FILENAME = "Прайс сваи.xls"

TEST_USERS = [
    {
        "id": 1,
        "username": "admin",
        "role": "admin",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
    {
        "id": 2,
        "username": "prod_user",
        "role": "production",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
    {
        "id": 3,
        "username": "manager_a",
        "role": "manager",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
]


@pytest.fixture()
def import_env(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    from app.core.settings import get_settings

    pb_db = tmp_path / "pb.db"
    plita_db = tmp_path / "plita.db"
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PB_DB_PATH", str(pb_db))
    monkeypatch.setenv("PLITA_DB_PATH", str(plita_db))
    get_settings.cache_clear()
    patch_auth_users(monkeypatch, TEST_USERS)

    conn = sqlite3.connect(pb_db)
    try:
        ensure_schema(conn)
        upsert(conn, "pile", MARK, match_status="missing")
        conn.commit()
    finally:
        conn.close()
    return pb_db


@pytest.fixture()
def client(import_env: Path):
    from app.main import create_app
    from tests.helpers.csrf import CsrfAwareTestClient

    del import_env
    return CsrfAwareTestClient(create_app())


def _admin() -> dict[str, str]:
    return session_cookie(1, "admin", "admin")


def _pile_row(source: str = PILE_FILENAME) -> PricelistRow:
    return PricelistRow(
        guid=PLAIN_GUID,
        name="Сваи С 30.30-3",
        price=None,
        unit="шт",
        source_file=source,
        row_index=3,
    )


@pytest.fixture()
def mock_parse(monkeypatch: pytest.MonkeyPatch):
    def fake_parse(path: Path | str) -> list[PricelistRow]:
        return [_pile_row(Path(path).name)]

    monkeypatch.setattr(
        "app.services.nomenclature_import_service.parse_pricelist_xls",
        fake_parse,
    )


def _upload(
    client,
    *,
    filename: str,
    content: bytes = b"xls-bytes",
    params: dict | None = None,
    data: dict | None = None,
):
    return client.post(
        API,
        params=params,
        data=data,
        files={"file": (filename, content, "application/vnd.ms-excel")},
        cookies=_admin(),
    )


def test_happy_path_fills_guid(client, mock_parse) -> None:
    response = _upload(client, filename=PILE_FILENAME)
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["new_guids_count"] == 1
    assert body["waiting_price"] == 0
    assert body["product_kind"] == "pile"
    assert body["new_guids"][0]["mark"] == MARK
    assert body["new_guids"][0]["guid"] == PLAIN_GUID
    assert body["new_guids"][0]["field"] == "guid_1c"
    assert "+1 GUID" in body["summary"]


def test_second_upload_idempotent(client, mock_parse) -> None:
    first = _upload(client, filename=PILE_FILENAME)
    assert first.status_code == 200, first.text
    assert first.json()["new_guids_count"] == 1

    second = _upload(client, filename=PILE_FILENAME)
    assert second.status_code == 200, second.text
    body = second.json()
    assert body["new_guids_count"] == 0
    assert body["new_guids"] == []
    assert body["unchanged"] >= 1


def test_bad_file_returns_422(client) -> None:
    response = _upload(
        client,
        filename=PILE_FILENAME,
        content=b"this is not an xls workbook",
    )
    assert response.status_code == 422
    detail = response.json()["detail"]
    assert "это не выгрузка Прайс-лист 1С" in detail


def test_mixed_filename_without_product_kind_returns_422(client, mock_parse) -> None:
    response = _upload(
        client,
        filename="Прайс ЛМ, ЛП, ЛС.xls",
        content=b"xls-bytes",
    )
    assert response.status_code == 422
    detail = response.json()["detail"]
    assert "product_kind" in detail
    assert "Прайс ЛМ" in detail


def test_mixed_filename_with_product_kind_query_ok(client, mock_parse) -> None:
    response = _upload(
        client,
        filename="Прайс ЛМ, ЛП, ЛС.xls",
        params={"product_kind": "pile"},
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["new_guids_count"] == 1
    assert body["product_kind"] == "pile"
