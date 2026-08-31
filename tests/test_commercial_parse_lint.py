"""HTTP tests for POST /commercial/parse lint extension."""

from __future__ import annotations

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app
from app.schemas.commercial import COMMERCIAL_PARSE_TEXT_MAX_LENGTH
from app.security.session import create_session_token
from app.services.commercial_service import CommercialService
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient

PLATE_KEYS = {
    "order",
    "normalized_text",
    "normalized_lines",
    "unparsed_lines",
    "warnings",
    "wide_plate_lines",
    "dobor_pairs",
    "diagnostics",
}

TYPE_CASES: list[tuple[str, str, str]] = [
    ("plates", "ПБ 78-12-8п 2", "xyz-not-a-plate"),
    ("piles", "С120.35-12 B25 5", "ПБ 78-12-8п 2"),
    ("steps", "ЛС11 10", "С120.35-12 5"),
    ("marches", "1ЛМ 27-11-14-4 2", "ПБ 78-12-8п 2"),
    ("bridge_piles", "С7-35Т5", "С120.35-12 2"),
    ("fbs", "ФБС 9.3.6-Т 2", "С120.35-12 2"),
]


@pytest.fixture()
def client(monkeypatch: pytest.MonkeyPatch) -> TestClient:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    get_settings.cache_clear()
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
    return CsrfAwareTestClient(create_app())


@pytest.fixture()
def auth_cookie(client: TestClient) -> dict[str, str]:
    token = create_session_token({"id": 1, "username": "tester", "role": "admin"}, ttl_seconds=300)
    client.cookies.set("app_session", token)
    return {"app_session": token}


def _line_by_index(payload: dict) -> dict[int, dict]:
    return {item["index"]: item for item in payload["lines"]}


def test_parse_without_product_type_keeps_plate_keys_and_adds_lines(
    client: TestClient, auth_cookie: dict[str, str]
) -> None:
    del auth_cookie
    response = client.post(
        "/api/v1/commercial/parse",
        json={"text": "ПБ 78-12-8п 2\n\nПБ 40,3/2,6-8п"},
    )
    assert response.status_code == 200
    body = response.json()
    assert PLATE_KEYS.issubset(body.keys())
    assert body["product_type"] == "plates"
    by_index = _line_by_index(body)
    assert by_index[0]["ok"] is True
    assert by_index[0]["empty"] is False
    assert by_index[1]["empty"] is True
    assert by_index[1]["ok"] is True
    assert by_index[2]["ok"] is False
    assert by_index[2]["text"] == "ПБ 40,3/2,6-8п"


@pytest.mark.parametrize(("product_type", "ok_line", "bad_line"), TYPE_CASES)
def test_parse_lints_each_product_type(
    client: TestClient,
    auth_cookie: dict[str, str],
    product_type: str,
    ok_line: str,
    bad_line: str,
) -> None:
    del auth_cookie
    response = client.post(
        "/api/v1/commercial/parse",
        json={"text": f"{ok_line}\n{bad_line}", "product_type": product_type},
    )
    assert response.status_code == 200
    body = response.json()
    assert body["product_type"] == product_type
    assert "lines" in body
    assert "unparsed_lines" in body
    by_index = _line_by_index(body)
    assert by_index[0]["ok"] is True
    assert by_index[1]["ok"] is False
    if product_type != "plates":
        assert set(body.keys()) == {"product_type", "lines", "unparsed_lines"}
        assert "order" not in body
        assert bad_line in body["unparsed_lines"]
        assert all("(пропущено" not in item for item in body["unparsed_lines"])
    else:
        assert PLATE_KEYS.issubset(body.keys())


def test_parse_invalid_product_type_returns_422(client: TestClient, auth_cookie: dict[str, str]) -> None:
    del auth_cookie
    response = client.post(
        "/api/v1/commercial/parse",
        json={"text": "ПБ 78-12-8п 2", "product_type": "widgets"},
    )
    assert response.status_code == 422


def test_parse_without_cookie_is_unauthorized(client: TestClient) -> None:
    response = client.post("/api/v1/commercial/parse", json={"text": "ПБ 78-12-8п 2"})
    assert response.status_code == 401


def test_parse_lint_only_skips_full_plate_parse(
    client: TestClient, auth_cookie: dict[str, str], monkeypatch: pytest.MonkeyPatch
) -> None:
    del auth_cookie

    def boom(self, text: str):
        raise AssertionError("CommercialService.parse must not run for lint_only")

    monkeypatch.setattr(CommercialService, "parse", boom)
    response = client.post(
        "/api/v1/commercial/parse",
        json={"text": "ПБ 78-12-8п 2", "product_type": "plates", "lint_only": True},
    )
    assert response.status_code == 200
    body = response.json()
    assert set(body.keys()) == {"product_type", "lines", "unparsed_lines"}
    assert body["product_type"] == "plates"
    assert body["lines"][0]["ok"] is True
    assert "order" not in body


def test_parse_text_over_max_length_returns_422(client: TestClient, auth_cookie: dict[str, str]) -> None:
    del auth_cookie
    response = client.post(
        "/api/v1/commercial/parse",
        json={"text": "П" * (COMMERCIAL_PARSE_TEXT_MAX_LENGTH + 1)},
    )
    assert response.status_code == 422
