"""WP5 S5: server-side RBAC negative tests (role guards are authoritative, not the SPA)."""

from __future__ import annotations

from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient
from app.security.session import create_session_token

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"

USERS = [
    {
        "id": 1,
        "username": "admin",
        "role": "admin",
        "manager_id": None,
        "is_active": 1,
        "created_at": "2026-01-01 00:00:00",
    },
    {
        "id": 2,
        "username": "prod_user",
        "role": "production",
        "manager_id": None,
        "is_active": 1,
        "created_at": "2026-01-01 00:00:00",
    },
    {
        "id": 3,
        "username": "manager_a",
        "role": "manager",
        "manager_id": None,
        "is_active": 1,
        "created_at": "2026-01-01 00:00:00",
    },
]

ADMIN_GET_ENDPOINTS = (
    "/api/v1/admin/users",
    "/api/v1/admin/db/stats",
)

ADMIN_POST_ENDPOINTS = (
    "/api/v1/admin/db/reset/full",
    "/api/v1/admin/db/reset/kp-only",
    "/api/v1/admin/db/reset/plans-only",
    "/api/v1/admin/db/reset/calendar-only",
    "/api/v1/admin/db/recover-plates",
)

PRODUCTION_FORBIDDEN_COMMERCIAL = (
    ("GET", "/api/v1/commercial/archive?section=archived", None),
    ("POST", "/api/v1/commercial/parse", {"text": "ПБ 60-12-8 1"}),
    ("GET", "/api/v1/offers", None),
)


def _session_cookie(user_id: int, role: str, username: str) -> dict[str, str]:
    return {
        "app_session": create_session_token(
            {"id": user_id, "username": username, "role": role},
            ttl_seconds=300,
        )
    }


@pytest.fixture()
def rbac_client(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> CsrfAwareTestClient:
    db_path = tmp_path / "plita.db"
    db_path.touch()
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", str(db_path))
    get_settings.cache_clear()
    patch_auth_users(monkeypatch, USERS)
    return CsrfAwareTestClient(create_app())


@pytest.mark.parametrize("endpoint", ADMIN_GET_ENDPOINTS)
def test_production_forbidden_on_admin_get(
    rbac_client: CsrfAwareTestClient,
    endpoint: str,
) -> None:
    response = rbac_client.get(
        endpoint,
        cookies=_session_cookie(2, "production", "prod_user"),
    )
    assert response.status_code == 403
    assert response.json()["detail"] == "Forbidden"


@pytest.mark.parametrize("endpoint", ADMIN_POST_ENDPOINTS)
def test_production_forbidden_on_admin_post(
    rbac_client: CsrfAwareTestClient,
    endpoint: str,
) -> None:
    response = rbac_client.post(
        endpoint,
        cookies=_session_cookie(2, "production", "prod_user"),
    )
    assert response.status_code == 403
    assert response.json()["detail"] == "Forbidden"


@pytest.mark.parametrize("endpoint", ADMIN_GET_ENDPOINTS)
def test_manager_forbidden_on_admin_get(
    rbac_client: CsrfAwareTestClient,
    endpoint: str,
) -> None:
    response = rbac_client.get(
        endpoint,
        cookies=_session_cookie(3, "manager", "manager_a"),
    )
    assert response.status_code == 403
    assert response.json()["detail"] == "Forbidden"


@pytest.mark.parametrize("endpoint", ADMIN_POST_ENDPOINTS)
def test_manager_forbidden_on_admin_destructive_post(
    rbac_client: CsrfAwareTestClient,
    endpoint: str,
) -> None:
    response = rbac_client.post(
        endpoint,
        cookies=_session_cookie(3, "manager", "manager_a"),
    )
    assert response.status_code == 403
    assert response.json()["detail"] == "Forbidden"


@pytest.mark.parametrize("method,path,json_body", PRODUCTION_FORBIDDEN_COMMERCIAL)
def test_production_forbidden_on_commercial_pii(
    rbac_client: CsrfAwareTestClient,
    method: str,
    path: str,
    json_body: dict | None,
) -> None:
    cookies = _session_cookie(2, "production", "prod_user")
    if method == "GET":
        response = rbac_client.get(path, cookies=cookies)
    else:
        response = rbac_client.post(path, json=json_body, cookies=cookies)
    assert response.status_code == 403
    assert response.json()["detail"] == "Forbidden"


def test_manager_forbidden_on_production_plans(rbac_client: CsrfAwareTestClient) -> None:
    response = rbac_client.get(
        "/api/v1/production/plans",
        cookies=_session_cookie(3, "manager", "manager_a"),
    )
    assert response.status_code == 403
    assert response.json()["detail"] == "Forbidden"


def test_production_allowed_on_production_calendar(rbac_client: CsrfAwareTestClient) -> None:
    response = rbac_client.get(
        "/api/v1/production/calendar",
        cookies=_session_cookie(2, "production", "prod_user"),
    )
    assert response.status_code == 200
