"""PD-008 / PD-011: HTTP стол прайсов + 403 экономиста на чужие API."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pandas as pd
import pytest

from app.core.settings import get_settings
from app.main import create_app
from app.services.price_desk_service import MSG_NEED_SHA, MSG_SHA_MISMATCH, file_sha256
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient
from tests.helpers.production_api_fixtures import VALID_APP_SECRET_KEY, session_cookie

STATUS = "/api/v1/prices/status"
PREVIEW = "/api/v1/prices/preview"
APPLY = "/api/v1/prices/apply"
XLSX_MEDIA = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

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
        "id": 6,
        "username": "economist_user",
        "role": "economist",
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
def desk_env(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    pb_db = tmp_path / "pb.db"
    plita_db = tmp_path / "plita.db"
    pb_db.touch()
    plita_db.touch()
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PB_DB_PATH", str(pb_db))
    monkeypatch.setenv("PLITA_DB_PATH", str(plita_db))
    get_settings.cache_clear()
    patch_auth_users(monkeypatch, TEST_USERS)
    return pb_db


@pytest.fixture()
def client(desk_env: Path) -> CsrfAwareTestClient:
    del desk_env
    return CsrfAwareTestClient(create_app())


def _economist() -> dict[str, str]:
    return session_cookie(6, "economist", "economist_user")


def _admin() -> dict[str, str]:
    return session_cookie(1, "admin", "admin")


def _manager() -> dict[str, str]:
    return session_cookie(3, "manager", "manager_a")


def _write_nikita_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", "Цены по прайсу, который прислал Никита"],
        [1, "ПБ 17-12-6", 6049],
        [2, "ПБ 17-12-8", 6100],
    ]
    pd.DataFrame(rows).to_excel(path, sheet_name="Прайс", index=False, header=False)


def _upload(client: CsrfAwareTestClient, url: str, path: Path, *, cookies, data: dict | None = None):
    return client.post(
        url,
        files={"file": (path.name, path.read_bytes(), XLSX_MEDIA)},
        data=data,
        cookies=cookies,
    )


def test_economist_status_200(client: CsrfAwareTestClient) -> None:
    response = client.get(STATUS, cookies=_economist())
    assert response.status_code == 200, response.text
    body = response.json()
    assert len(body["groups"]) == 7
    assert {g["product_kind"] for g in body["groups"]} == {
        "plates",
        "fbs",
        "march",
        "step",
        "bridge_pile",
        "pile",
        "composite_pile",
    }


def test_manager_forbidden_on_prices(client: CsrfAwareTestClient, tmp_path: Path) -> None:
    xlsx = tmp_path / "Расчет новых цен на ПБ 17.08.2026.xlsx"
    _write_nikita_xlsx(xlsx)
    assert client.get(STATUS, cookies=_manager()).status_code == 403
    assert _upload(client, PREVIEW, xlsx, cookies=_manager()).status_code == 403
    assert _upload(
        client, APPLY, xlsx, cookies=_manager(), data={"file_sha256": "ab" * 32}
    ).status_code == 403


def test_economist_preview_does_not_write(
    client: CsrfAwareTestClient, desk_env: Path, tmp_path: Path
) -> None:
    xlsx = tmp_path / "Расчет новых цен на ПБ 17.08.2026.xlsx"
    _write_nikita_xlsx(xlsx)
    response = _upload(client, PREVIEW, xlsx, cookies=_economist())
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["product_kind"] == "plates"
    assert body["parsed_rows"] == 2
    assert body["new"] == 2
    conn = sqlite3.connect(desk_env)
    try:
        tables = {
            row[0]
            for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")
        }
        if "prices" in tables:
            assert conn.execute("SELECT COUNT(*) FROM prices").fetchone()[0] == 0
    finally:
        conn.close()


def test_apply_without_sha256_422(client: CsrfAwareTestClient, tmp_path: Path) -> None:
    xlsx = tmp_path / "Расчет новых цен на ПБ 17.08.2026.xlsx"
    _write_nikita_xlsx(xlsx)
    response = _upload(client, APPLY, xlsx, cookies=_economist())
    assert response.status_code == 422
    assert MSG_NEED_SHA in response.json()["detail"]


def test_apply_wrong_sha256_409_no_write(
    client: CsrfAwareTestClient, desk_env: Path, tmp_path: Path
) -> None:
    xlsx = tmp_path / "Расчет новых цен на ПБ 17.08.2026.xlsx"
    _write_nikita_xlsx(xlsx)
    response = _upload(
        client, APPLY, xlsx, cookies=_economist(), data={"file_sha256": "ff" * 32}
    )
    assert response.status_code == 409
    assert MSG_SHA_MISMATCH in response.json()["detail"]
    conn = sqlite3.connect(desk_env)
    try:
        tables = {
            row[0]
            for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")
        }
        if "prices" in tables:
            assert conn.execute("SELECT COUNT(*) FROM prices").fetchone()[0] == 0
    finally:
        conn.close()


def test_economist_apply_then_empty_diff(
    client: CsrfAwareTestClient, desk_env: Path, tmp_path: Path
) -> None:
    xlsx = tmp_path / "Расчет новых цен на ПБ 17.08.2026.xlsx"
    _write_nikita_xlsx(xlsx)
    digest = file_sha256(xlsx.read_bytes())
    preview = _upload(client, PREVIEW, xlsx, cookies=_economist())
    assert preview.status_code == 200
    assert preview.json()["file_sha256"] == digest
    applied = _upload(
        client, APPLY, xlsx, cookies=_economist(), data={"file_sha256": digest}
    )
    assert applied.status_code == 200, applied.text
    conn = sqlite3.connect(desk_env)
    try:
        assert conn.execute("SELECT COUNT(*) FROM prices").fetchone()[0] == 2
    finally:
        conn.close()
    second = _upload(client, PREVIEW, xlsx, cookies=_economist())
    assert second.status_code == 200
    body = second.json()
    assert body["changed"] == 0
    assert body["new"] == 0
    assert body["unchanged"] == 2


def test_admin_can_apply(client: CsrfAwareTestClient, tmp_path: Path) -> None:
    xlsx = tmp_path / "Расчет новых цен на ПБ 17.08.2026.xlsx"
    _write_nikita_xlsx(xlsx)
    digest = file_sha256(xlsx.read_bytes())
    response = _upload(client, APPLY, xlsx, cookies=_admin(), data={"file_sha256": digest})
    assert response.status_code == 200, response.text


def test_composite_pile_preview(client: CsrfAwareTestClient, tmp_path: Path) -> None:
    xlsx = tmp_path / "Прайс на составные сваи от 07.09.2026.xlsx"
    rows = [
        [None, "Наименование", 15, 20, 22.5, 25, "30 на граните"],
        [1, "Сваи С 60.30-ВС.1", 100.0, 110.0, 120.0, 130.0, 140.0],
    ]
    pd.DataFrame(rows).to_excel(xlsx, sheet_name="Прайс", index=False, header=False)
    response = _upload(client, PREVIEW, xlsx, cookies=_economist())
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["product_kind"] == "composite_pile"
    assert body["parsed_rows"] == 5


@pytest.mark.parametrize(
    "path",
    [
        "/api/v1/offers",
        "/api/v1/commercial/archive",
        "/api/v1/gsm/vehicles",
        "/api/v1/production/plans",
        "/api/v1/logistics/shipments",
    ],
)
def test_economist_forbidden_on_foreign_get(
    client: CsrfAwareTestClient, path: str
) -> None:
    response = client.get(path, cookies=_economist())
    assert response.status_code == 403, response.text


def test_economist_can_post_1c_import(client: CsrfAwareTestClient) -> None:
    response = client.post(
        "/api/v1/nomenclature/import-1c",
        files={"file": ("Прайс сваи.xls", b"xls-bytes", "application/vnd.ms-excel")},
        cookies=_economist(),
    )
    assert response.status_code != 403, response.text


def test_manager_forbidden_on_1c_import(client: CsrfAwareTestClient) -> None:
    response = client.post(
        "/api/v1/nomenclature/import-1c",
        files={"file": ("Прайс сваи.xls", b"xls-bytes", "application/vnd.ms-excel")},
        cookies=_manager(),
    )
    assert response.status_code == 403, response.text
