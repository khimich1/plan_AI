"""Task T2: роль accountant + REQUIRE_ACCOUNTING (AuthZ для модуля ГСМ).

TDD: тесты должны падать, пока worker не добавит
DEFAULT_ACCOUNTANT_ROLE, REQUIRE_ACCOUNTING и Literal «accountant».
"""

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

PROBE_PATH = "/__test__/accounting-probe"
LOGISTICS_PROBE_PATH = "/__test__/logistics-probe"

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
        "id": 4,
        "username": "logist",
        "role": "logistics",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
]


def _accounting_probe_app() -> FastAPI:
    """Минимальный probe с Depends(REQUIRE_ACCOUNTING) — без правок production-роутеров."""
    from app.dependencies.auth import REQUIRE_ACCOUNTING

    app = FastAPI()

    @app.get(PROBE_PATH)
    def _probe(user: dict = Depends(REQUIRE_ACCOUNTING)) -> dict[str, object]:
        return {"ok": True, "role": user.get("role")}

    return app


def _logistics_probe_app() -> FastAPI:
    """Регрессия: REQUIRE_LOGISTICS не сломан добавлением accountant."""
    from app.dependencies.auth import REQUIRE_LOGISTICS

    app = FastAPI()

    @app.get(LOGISTICS_PROBE_PATH)
    def _probe(user: dict = Depends(REQUIRE_LOGISTICS)) -> dict[str, object]:
        return {"ok": True, "role": user.get("role")}

    return app


@pytest.fixture()
def gsm_auth_env(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    db_path = tmp_path / "plita.db"
    db_path.touch()
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", str(db_path))
    monkeypatch.setenv("PB_DB_PATH", str(db_path))
    get_settings.cache_clear()
    patch_auth_users(monkeypatch, TEST_USERS)
    return db_path


@pytest.fixture()
def accounting_client(gsm_auth_env: Path) -> TestClient:
    del gsm_auth_env
    return TestClient(_accounting_probe_app())


@pytest.fixture()
def logistics_client(gsm_auth_env: Path) -> TestClient:
    del gsm_auth_env
    return TestClient(_logistics_probe_app())


@pytest.fixture()
def api_client(gsm_auth_env: Path) -> CsrfAwareTestClient:
    del gsm_auth_env
    return CsrfAwareTestClient(create_app())


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------


def test_default_accountant_role_constant() -> None:
    from app.core.constants import DEFAULT_ACCOUNTANT_ROLE

    assert DEFAULT_ACCOUNTANT_ROLE == "accountant"


# ---------------------------------------------------------------------------
# REQUIRE_ACCOUNTING — 200/403 matrix (probe route)
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("user_id", "role", "username"),
    [
        (1, "admin", "admin"),
        (5, "accountant", "accountant_user"),
    ],
)
def test_require_accounting_allows_admin_and_accountant(
    accounting_client: TestClient,
    user_id: int,
    role: str,
    username: str,
) -> None:
    response = accounting_client.get(
        PROBE_PATH,
        cookies=session_cookie(user_id, role, username),
    )
    assert response.status_code == 200
    assert response.json() == {"ok": True, "role": role}


@pytest.mark.parametrize(
    ("user_id", "role", "username"),
    [
        (3, "manager", "manager_a"),
        (2, "production", "prod_user"),
        (4, "logistics", "logist"),
    ],
)
def test_require_accounting_forbids_other_roles(
    accounting_client: TestClient,
    user_id: int,
    role: str,
    username: str,
) -> None:
    response = accounting_client.get(
        PROBE_PATH,
        cookies=session_cookie(user_id, role, username),
    )
    assert response.status_code == 403
    assert response.json()["detail"] == "Forbidden"


def test_require_accounting_rejects_unauthenticated(accounting_client: TestClient) -> None:
    response = accounting_client.get(PROBE_PATH)
    assert response.status_code == 401


# ---------------------------------------------------------------------------
# Admin register — роль accountant в Literal / POST /auth/register
# ---------------------------------------------------------------------------


def test_register_user_request_accepts_accountant_role() -> None:
    payload = RegisterUserRequest(
        username="buchgalter",
        password=_VALID_PASSWORD,
        role="accountant",
    )
    assert payload.role == "accountant"


def test_admin_can_register_user_with_accountant_role(api_client: CsrfAwareTestClient) -> None:
    response = api_client.post(
        "/api/v1/auth/register",
        json={
            "username": "buchgalter",
            "password": _VALID_PASSWORD,
            "role": "accountant",
        },
        cookies=session_cookie(1, "admin", "admin"),
    )
    assert response.status_code == 201
    body = response.json()
    assert body["created"] is True
    assert body["user"]["username"] == "buchgalter"
    assert body["user"]["role"] == "accountant"


# ---------------------------------------------------------------------------
# Регрессия: существующие роли и REQUIRE_LOGISTICS
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "role",
    ["admin", "manager", "production", "logistics"],
)
def test_register_user_request_still_accepts_existing_roles(role: str) -> None:
    payload = RegisterUserRequest(
        username=f"user_{role}",
        password=_VALID_PASSWORD,
        role=role,  # type: ignore[arg-type]
    )
    assert payload.role == role


def test_register_user_request_still_rejects_unknown_role() -> None:
    with pytest.raises(ValidationError):
        RegisterUserRequest(
            username="nope",
            password=_VALID_PASSWORD,
            role="superadmin",  # type: ignore[arg-type]
        )


@pytest.mark.parametrize(
    ("user_id", "role", "username"),
    [
        (1, "admin", "admin"),
        (4, "logistics", "logist"),
    ],
)
def test_require_logistics_still_allows_admin_and_logistics(
    logistics_client: TestClient,
    user_id: int,
    role: str,
    username: str,
) -> None:
    response = logistics_client.get(
        LOGISTICS_PROBE_PATH,
        cookies=session_cookie(user_id, role, username),
    )
    assert response.status_code == 200
    assert response.json() == {"ok": True, "role": role}


def test_require_logistics_still_forbids_manager(logistics_client: TestClient) -> None:
    response = logistics_client.get(
        LOGISTICS_PROBE_PATH,
        cookies=session_cookie(3, "manager", "manager_a"),
    )
    assert response.status_code == 403
    assert response.json()["detail"] == "Forbidden"


def test_admin_can_still_register_logistics_role(api_client: CsrfAwareTestClient) -> None:
    response = api_client.post(
        "/api/v1/auth/register",
        json={
            "username": "logist_new",
            "password": _VALID_PASSWORD,
            "role": "logistics",
        },
        cookies=session_cookie(1, "admin", "admin"),
    )
    assert response.status_code == 201
    assert response.json()["user"]["role"] == "logistics"
