from __future__ import annotations

import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.main import create_app
from app.repositories.auth_repository import AuthRepository
from app.security.session import create_session_token
from app.web.legacy_deprecation import (
    DEPRECATION_HEADER,
    DEPRECATION_HEADER_VALUE,
    SPA_ARCHIVE,
    SPA_LOGIN,
    SPA_NEW,
    SPA_PRODUCTION,
    spa_draft_url,
)
from tests.helpers.auth_fixtures import patch_auth_users

VALID_APP_SECRET_KEY = "test-secret-key-for-pytest-must-be-32-chars-min"


@pytest.fixture()
def client(monkeypatch: pytest.MonkeyPatch) -> TestClient:
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    get_settings.cache_clear()
    patch_auth_users(
        monkeypatch,
        [
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
                "username": "prod",
                "role": "production",
                "manager_id": None,
                "is_active": 1,
                "created_at": "2026-01-01 00:00:00",
            },
        ],
    )
    return TestClient(create_app())


def _session_cookie(*, user_id: int, username: str, role: str) -> dict[str, str]:
    token = create_session_token({"id": user_id, "username": username, "role": role}, ttl_seconds=300)
    return {"app_session": token}


@pytest.mark.parametrize(
    ("path", "expected_location"),
    [
        ("/web/login", SPA_LOGIN),
        ("/web", SPA_NEW),
        ("/web/managers", SPA_ARCHIVE),
        ("/web/offers", SPA_ARCHIVE),
        ("/web/offers/new", SPA_NEW),
        ("/web/production", SPA_PRODUCTION),
    ],
)
def test_legacy_get_routes_redirect_to_spa(
    client: TestClient,
    path: str,
    expected_location: str,
) -> None:
    cookies = _session_cookie(user_id=1, username="admin", role="admin")
    if path == "/web/login":
        cookies = {}

    response = client.get(path, cookies=cookies, follow_redirects=False)

    assert response.status_code == 303
    assert response.headers["location"] == expected_location
    assert response.headers.get(DEPRECATION_HEADER) == DEPRECATION_HEADER_VALUE
    assert "successor-version" in response.headers.get("link", "")


def test_legacy_dashboard_redirects_production_role_to_spa_production(
    client: TestClient,
) -> None:
    cookies = _session_cookie(user_id=2, username="prod", role="production")
    response = client.get("/web", cookies=cookies, follow_redirects=False)

    assert response.status_code == 303
    assert response.headers["location"] == SPA_PRODUCTION


def test_spa_shell_routes_exist(client: TestClient) -> None:
    admin = _session_cookie(user_id=1, username="admin", role="admin")
    production = _session_cookie(user_id=2, username="prod", role="production")

    login = client.get("/commercial-offer/login")
    assert login.status_code == 200

    archive = client.get("/commercial-offer/archive", cookies=admin)
    assert archive.status_code == 200

    production_page = client.get("/commercial-offer/production", cookies=production)
    assert production_page.status_code == 200


def test_web_login_post_redirects_to_spa_home(monkeypatch: pytest.MonkeyPatch, client: TestClient) -> None:
    def fake_authenticate(self, user: str, pwd: str) -> dict | None:
        if user == "admin" and pwd == "StrongPassword123!":
            return {"id": 1, "username": "admin", "role": "admin"}
        return None

    monkeypatch.setattr(AuthRepository, "authenticate", fake_authenticate)

    response = client.post(
        "/web/login",
        data={"username": "admin", "password": "StrongPassword123!"},
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert response.headers["location"] == SPA_NEW
    assert response.headers.get(DEPRECATION_HEADER) == DEPRECATION_HEADER_VALUE


def test_legacy_draft_get_redirects_to_spa_wizard(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from app.services.commercial_workflow_service import CommercialWorkflowService

    draft_id = "draft-legacy-redirect"
    monkeypatch.setattr(
        "app.web.router.check_draft_ownership",
        lambda _draft_id, _user: None,
    )
    monkeypatch.setattr(
        CommercialWorkflowService,
        "get_draft_details",
        lambda self, _draft_id: {"draft_id": draft_id},
    )

    cookies = _session_cookie(user_id=1, username="admin", role="admin")
    response = client.get(f"/web/offers/drafts/{draft_id}", cookies=cookies, follow_redirects=False)

    assert response.status_code == 303
    assert response.headers["location"] == spa_draft_url(draft_id)
    assert response.headers.get(DEPRECATION_HEADER) == DEPRECATION_HEADER_VALUE



from urllib.parse import unquote


def test_legacy_new_offer_post_invalid_manager_redirects_with_error(
    client: TestClient,
) -> None:
    cookies = _session_cookie(user_id=1, username="admin", role="admin")
    response = client.post(
        "/web/offers/new",
        data={"manager_id": "not-a-number", "client_name": "Test"},
        cookies=cookies,
        follow_redirects=False,
    )

    assert response.status_code == 303
    location = response.headers["location"]
    assert location.startswith("/commercial-offer/new?error=")
    assert unquote(location.split("error=", 1)[1]) == "Выберите менеджера."


def test_legacy_new_offer_post_invalid_manager_json_when_accept_json(
    client: TestClient,
) -> None:
    cookies = _session_cookie(user_id=1, username="admin", role="admin")
    response = client.post(
        "/web/offers/new",
        data={"manager_id": "not-a-number", "client_name": "Test"},
        cookies=cookies,
        headers={"Accept": "application/json"},
        follow_redirects=False,
    )

    assert response.status_code == 400
    assert response.json()["detail"] == "Выберите менеджера."
    assert response.headers.get(DEPRECATION_HEADER) == DEPRECATION_HEADER_VALUE


def test_legacy_new_offer_post_success_redirects_to_spa_draft(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from app.services.commercial_service import CommercialService
    from app.services.commercial_workflow_service import CommercialWorkflowService

    draft_id = "draft-web-post"
    monkeypatch.setattr(
        CommercialService,
        "list_managers",
        lambda self: [{"id": 1, "fio": "РРІР°РЅ РРІР°РЅРѕРІ", "contact_number": "", "email": ""}],
    )

    async def fake_create(self, **kwargs):
        return {"draft_id": draft_id}

    monkeypatch.setattr(CommercialWorkflowService, "create_draft_from_form", fake_create)

    cookies = _session_cookie(user_id=1, username="admin", role="admin")
    response = client.post(
        "/web/offers/new",
        data={
            "text": "РџР‘ 78-12-8Рї 2",
            "manager_id": "1",
            "client_name": "РћРћРћ РўРµСЃС‚",
            "discount_percent": "5",
            "delivery_conditions": "РЎР°РјРѕРІС‹РІРѕР·",
            "payment_conditions": "100% РїСЂРµРґРѕРїР»Р°С‚Р°",
        },
        cookies=cookies,
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert response.headers["location"] == spa_draft_url(draft_id)
    assert response.headers.get(DEPRECATION_HEADER) == DEPRECATION_HEADER_VALUE


def test_legacy_draft_generate_files_post_redirects_to_spa(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from app.services.commercial_workflow_service import CommercialWorkflowService

    draft_id = "draft-generate-files"
    monkeypatch.setattr("app.web.router.check_draft_ownership", lambda _draft_id, _user: None)
    monkeypatch.setattr(CommercialWorkflowService, "generate_files", lambda self, _draft_id, **kwargs: None)

    cookies = _session_cookie(user_id=1, username="admin", role="admin")
    response = client.post(
        f"/web/offers/drafts/{draft_id}/generate-files",
        cookies=cookies,
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert response.headers["location"] == spa_draft_url(draft_id)
    assert response.headers.get(DEPRECATION_HEADER) == DEPRECATION_HEADER_VALUE


def test_legacy_draft_save_post_redirects_to_spa(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from app.services.commercial_workflow_service import CommercialWorkflowService

    draft_id = "draft-save"
    monkeypatch.setattr("app.web.router.check_draft_ownership", lambda _draft_id, _user: None)
    monkeypatch.setattr(CommercialWorkflowService, "save_offer", lambda self, _draft_id: {"saved_offer": {"kp_id": 1}})

    cookies = _session_cookie(user_id=1, username="admin", role="admin")
    response = client.post(
        f"/web/offers/drafts/{draft_id}/save",
        cookies=cookies,
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert response.headers["location"] == spa_draft_url(draft_id)
    assert response.headers.get(DEPRECATION_HEADER) == DEPRECATION_HEADER_VALUE


def test_commercial_offer_draft_stub_redirects_to_spa_wizard(client: TestClient) -> None:
    draft_id = "draft-stub"
    cookies = _session_cookie(user_id=1, username="admin", role="admin")
    response = client.get(f"/commercial-offer/drafts/{draft_id}", cookies=cookies, follow_redirects=False)

    assert response.status_code == 303
    assert response.headers["location"] == spa_draft_url(draft_id)

