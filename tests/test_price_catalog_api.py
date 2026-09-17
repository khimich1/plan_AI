"""HTTP tests for GET /api/v1/commercial/price-catalog."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app
from app.security.session import create_session_token
from core.pile_price_db import init_pile_prices_schema
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient
from tests.test_commercial_draft_append import _patch_price_db_path


def _seed_pile_prices(db_path: Path) -> None:
    init_pile_prices_schema(str(db_path))
    conn = sqlite3.connect(db_path)
    try:
        conn.executemany(
            "INSERT INTO pile_prices (mark, concrete_grade, price) VALUES (?, ?, ?)",
            [
                ("С110.30-9", "B25", 27585.43),
                ("С120.35-12", "B25", 44634.03),
            ],
        )
        conn.commit()
    finally:
        conn.close()


def _catalog_count(db_path: Path) -> tuple[int, list[str]]:
    conn = sqlite3.connect(db_path)
    try:
        count = conn.execute("SELECT COUNT(*) FROM pile_prices").fetchone()[0]
        stamps = [
            row[0]
            for row in conn.execute(
                "SELECT imported_at FROM pile_prices ORDER BY mark, concrete_grade"
            ).fetchall()
        ]
    finally:
        conn.close()
    return int(count), stamps


@pytest.fixture()
def catalog_env(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> Path:
    pb_db = tmp_path / "pb.db"
    _seed_pile_prices(pb_db)
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    _patch_price_db_path(monkeypatch, pb_db)
    get_settings.cache_clear()
    return pb_db


def _auth_client(monkeypatch: pytest.MonkeyPatch, *, role: str = "admin") -> TestClient:
    patch_auth_users(
        monkeypatch,
        [
            {
                "id": 1,
                "username": "tester",
                "role": role,
                "manager_id": None,
                "is_active": 1,
                "created_at": "2026-01-01 00:00:00",
            }
        ],
    )
    client = CsrfAwareTestClient(create_app())
    token = create_session_token({"id": 1, "username": "tester", "role": role}, ttl_seconds=300)
    client.cookies.set("app_session", token)
    return client


def test_price_catalog_piles_q_c110_returns_neighbor(
    catalog_env: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    client = _auth_client(monkeypatch)
    response = client.get(
        "/api/v1/commercial/price-catalog",
        params={"product_type": "piles", "q": "C110"},
    )
    assert response.status_code == 200, response.text
    items = list(response.json().get("items") or [])
    marks = {item["mark"] for item in items}
    assert "С110.30-9" in marks
    assert "С110.30-6" not in marks
    neighbor = next(item for item in items if item["mark"] == "С110.30-9")
    assert neighbor["concrete_grade"] == "B25"
    assert float(neighbor["price"]) == pytest.approx(27585.43)


def test_price_catalog_without_cookie_is_unauthorized(
    catalog_env: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    del catalog_env
    patch_auth_users(
        monkeypatch,
        [
            {
                "id": 1,
                "username": "tester",
                "role": "admin",
                "manager_id": None,
                "is_active": 1,
                "created_at": "2026-01-01 00:00:00",
            }
        ],
    )
    client = CsrfAwareTestClient(create_app())
    response = client.get(
        "/api/v1/commercial/price-catalog",
        params={"product_type": "piles"},
    )
    assert response.status_code == 401


def test_price_catalog_accountant_is_forbidden(
    catalog_env: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    del catalog_env
    client = _auth_client(monkeypatch, role="accountant")
    response = client.get(
        "/api/v1/commercial/price-catalog",
        params={"product_type": "piles"},
    )
    assert response.status_code == 403


def test_price_catalog_unknown_type_400(
    catalog_env: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    del catalog_env
    client = _auth_client(monkeypatch)
    response = client.get(
        "/api/v1/commercial/price-catalog",
        params={"product_type": "widgets"},
    )
    assert response.status_code == 400, response.text
    assert "Неизвестный" in str(response.json().get("detail") or "")


def test_price_catalog_cap_400(
    catalog_env: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    del catalog_env
    monkeypatch.setattr("core.price_catalog_query.CATALOG_CAP", 1)
    client = _auth_client(monkeypatch)
    response = client.get(
        "/api/v1/commercial/price-catalog",
        params={"product_type": "piles", "q": ""},
    )
    assert response.status_code == 400, response.text
    detail = str(response.json().get("detail") or "")
    assert "уточните поиск" in detail.lower()


def test_price_catalog_get_does_not_write_sqlite(
    catalog_env: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    count_before, stamps_before = _catalog_count(catalog_env)
    client = _auth_client(monkeypatch)
    response = client.get(
        "/api/v1/commercial/price-catalog",
        params={"product_type": "piles", "q": "C110"},
    )
    assert response.status_code == 200, response.text
    count_after, stamps_after = _catalog_count(catalog_env)
    assert count_after == count_before
    assert stamps_after == stamps_before
