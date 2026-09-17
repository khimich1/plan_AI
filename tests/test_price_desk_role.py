"""PD-005: роль economist + REQUIRE_PRICES."""

from __future__ import annotations

from pathlib import Path

import pytest
from fastapi import Depends, FastAPI
from fastapi.testclient import TestClient
from pydantic import ValidationError

from app.core.settings import get_settings
from app.main import create_app
from app.schemas.auth import RegisterUserRequest
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient
from tests.helpers.production_api_fixtures import VALID_APP_SECRET_KEY, session_cookie

PROBE_PATH = "/__test__/prices-probe"
_VALID_PASSWORD = "ValidPass1234"

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
        "id": 5,
        "username": "accountant_user",
        "role": "accountant",
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


def _prices_probe_app() -> FastAPI:
    from app.dependencies.auth import REQUIRE_PRICES

    app = FastAPI()

    @app.get(PROBE_PATH)
    def _probe(user: dict = Depends(REQUIRE_PRICES)) -> dict[str, object]:
        return {"ok": True, "role": user.get("role")}

    return app


@pytest.fixture()
def prices_auth_env(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    db_path = tmp_path / "plita.db"
    db_path.touch()
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", str(db_path))
    monkeypatch.setenv("PB_DB_PATH", str(db_path))
    get_settings.cache_clear()
    patch_auth_users(monkeypatch, TEST_USERS)
    return db_path


@pytest.fixture()
def prices_client(prices_auth_env: Path) -> TestClient:
    del prices_auth_env
    return TestClient(_prices_probe_app())


@pytest.fixture()
def api_client(prices_auth_env: Path) -> CsrfAwareTestClient:
    del prices_auth_env
    return CsrfAwareTestClient(create_app())


def test_default_economist_role_constant() -> None:
    from app.core.constants import DEFAULT_ECONOMIST_ROLE

    assert DEFAULT_ECONOMIST_ROLE == "economist"


@pytest.mark.parametrize(
    ("user_id", "role", "username"),
    [
        (1, "admin", "admin"),
        (6, "economist", "economist_user"),
    ],
)
def test_require_prices_allows_admin_and_economist(
    prices_client: TestClient,
    user_id: int,
    role: str,
    username: str,
) -> None:
    response = prices_client.get(
        PROBE_PATH,
        cookies=session_cookie(user_id, role, username),
    )
    assert response.status_code == 200
    assert response.json() == {"ok": True, "role": role}


@pytest.mark.parametrize(
    ("user_id", "role", "username"),
    [
        (3, "manager", "manager_a"),
        (5, "accountant", "accountant_user"),
    ],
)
def test_require_prices_forbids_other_roles(
    prices_client: TestClient,
    user_id: int,
    role: str,
    username: str,
) -> None:
    response = prices_client.get(
        PROBE_PATH,
        cookies=session_cookie(user_id, role, username),
    )
    assert response.status_code == 403
    assert response.json()["detail"] == "Forbidden"


def test_register_user_request_accepts_economist_role() -> None:
    payload = RegisterUserRequest(
        username="ekonomist",
        password=_VALID_PASSWORD,
        role="economist",
    )
    assert payload.role == "economist"


def test_register_user_request_rejects_garbage_role() -> None:
    with pytest.raises(ValidationError):
        RegisterUserRequest(
            username="nope",
            password=_VALID_PASSWORD,
            role="superadmin",  # type: ignore[arg-type]
        )


def test_admin_can_register_user_with_economist_role(api_client: CsrfAwareTestClient) -> None:
    response = api_client.post(
        "/api/v1/auth/register",
        json={
            "username": "ekonomist",
            "password": _VALID_PASSWORD,
            "role": "economist",
        },
        cookies=session_cookie(1, "admin", "admin"),
    )
    assert response.status_code == 201
    body = response.json()
    assert body["created"] is True
    assert body["user"]["username"] == "ekonomist"
    assert body["user"]["role"] == "economist"
