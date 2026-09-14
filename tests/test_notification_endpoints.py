"""Task 12: in-web notification list + mark-read endpoints."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from fastapi.testclient import TestClient

from app.api.v1.endpoints.notifications import get_promise_repository
from app.core.settings import get_settings
from app.main import create_app
from app.repositories.promise_repository import PromiseRepository
from app.security.session import create_session_token
from core import kp_db_schema
from core.kp_db_common import _connect
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient

MANAGER = {
    "id": 7,
    "username": "alice",
    "role": "manager",
    "manager_id": None,
    "is_active": 1,
    "created_at": "2026-01-01 00:00:00",
    "session_version": 0,
}
OTHER = {
    "id": 3,
    "username": "bob",
    "role": "manager",
    "manager_id": None,
    "is_active": 1,
    "created_at": "2026-01-01 00:00:00",
    "session_version": 0,
}
PRODUCTION = {
    "id": 2,
    "username": "planner",
    "role": "production",
    "manager_id": None,
    "is_active": 1,
    "created_at": "2026-01-01 00:00:00",
    "session_version": 0,
}


@pytest.fixture()
def fake_repo() -> MagicMock:
    return MagicMock()


@pytest.fixture()
def client(monkeypatch: pytest.MonkeyPatch, fake_repo: MagicMock) -> TestClient:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    get_settings.cache_clear()
    patch_auth_users(monkeypatch, [MANAGER, OTHER, PRODUCTION])
    app = create_app()
    app.dependency_overrides[get_promise_repository] = lambda: fake_repo
    return CsrfAwareTestClient(app)


def _cookie(user: dict) -> dict[str, str]:
    return {
        "app_session": create_session_token(
            {"id": user["id"], "username": user["username"], "role": user["role"]},
            ttl_seconds=300,
        ),
    }


def test_list_requires_auth(client: TestClient) -> None:
    response = client.get("/api/v1/notifications")
    assert response.status_code == 401


def test_list_filters_by_current_user_and_returns_unread_count(
    client: TestClient, fake_repo: MagicMock
) -> None:
    fake_repo.list_notifications.return_value = [
        {
            "id": 1,
            "user_id": 7,
            "kind": "promise_excluded",
            "payload_json": '{"kp_id": 42, "week_start": "2026-09-07", "reason": "нет арматуры"}',
            "read_at": None,
            "created_at": "2026-09-03T12:00:00",
        }
    ]
    fake_repo.count_unread_notifications.return_value = 1

    response = client.get("/api/v1/notifications", cookies=_cookie(MANAGER))

    assert response.status_code == 200
    body = response.json()
    assert body["unread_count"] == 1
    assert body["items"][0]["kind"] == "promise_excluded"
    assert body["items"][0]["payload"] == {
        "kp_id": 42,
        "week_start": "2026-09-07",
        "reason": "нет арматуры",
    }
    fake_repo.list_notifications.assert_called_once_with(user_id=7, unread_only=False)
    fake_repo.count_unread_notifications.assert_called_once_with(user_id=7)


def test_list_unread_query_asks_repo_for_unread_only(
    client: TestClient, fake_repo: MagicMock
) -> None:
    fake_repo.list_notifications.return_value = []
    fake_repo.count_unread_notifications.return_value = 0

    response = client.get("/api/v1/notifications?unread=true", cookies=_cookie(MANAGER))

    assert response.status_code == 200
    fake_repo.list_notifications.assert_called_once_with(user_id=7, unread_only=True)


def test_production_role_can_list(client: TestClient, fake_repo: MagicMock) -> None:
    fake_repo.list_notifications.return_value = []
    fake_repo.count_unread_notifications.return_value = 0

    response = client.get("/api/v1/notifications", cookies=_cookie(PRODUCTION))

    assert response.status_code == 200
    fake_repo.list_notifications.assert_called_once_with(user_id=2, unread_only=False)


def test_mark_read_own_notification(client: TestClient, fake_repo: MagicMock) -> None:
    fake_repo.mark_notification_read.return_value = {
        "id": 11,
        "user_id": 7,
        "kind": "promise_excluded",
        "payload_json": '{"kp_id": 1}',
        "read_at": "2026-09-03T18:00:00",
        "created_at": "2026-09-03T12:00:00",
    }

    response = client.post("/api/v1/notifications/11/read", cookies=_cookie(MANAGER))

    assert response.status_code == 200
    assert response.json() == {"id": 11, "read_at": "2026-09-03T18:00:00"}
    kwargs = fake_repo.mark_notification_read.call_args
    assert kwargs.args[0] == 11
    assert kwargs.kwargs["user_id"] == 7


def test_mark_read_foreign_or_missing_is_404(
    client: TestClient, fake_repo: MagicMock
) -> None:
    fake_repo.mark_notification_read.return_value = None

    response = client.post("/api/v1/notifications/99/read", cookies=_cookie(MANAGER))

    assert response.status_code == 404
    fake_repo.mark_notification_read.assert_called_once()
    assert fake_repo.mark_notification_read.call_args.kwargs["user_id"] == 7


def test_other_user_does_not_see_manager_filter(
    client: TestClient, fake_repo: MagicMock
) -> None:
    fake_repo.list_notifications.return_value = []
    fake_repo.count_unread_notifications.return_value = 0

    response = client.get("/api/v1/notifications", cookies=_cookie(OTHER))

    assert response.status_code == 200
    fake_repo.list_notifications.assert_called_once_with(user_id=3, unread_only=False)


def test_repo_unread_filter_and_mark_read(tmp_path: Path) -> None:
    db_path = str(tmp_path / "notes.db")
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    with _connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO notifications (user_id, kind, payload_json, created_at)
            VALUES (7, 'promise_excluded', '{"kp_id": 1}', '2026-09-03T12:00:00')
            """
        )
        conn.execute(
            """
            INSERT INTO notifications (user_id, kind, payload_json, read_at, created_at)
            VALUES (7, 'promise_excluded', '{"kp_id": 2}', '2026-09-03T13:00:00',
                    '2026-09-03T12:30:00')
            """
        )
        conn.execute(
            """
            INSERT INTO notifications (user_id, kind, payload_json, created_at)
            VALUES (3, 'promise_excluded', '{"kp_id": 9}', '2026-09-03T12:00:00')
            """
        )
        conn.commit()

    repo = PromiseRepository(db_path=db_path)
    assert repo.count_unread_notifications(user_id=7) == 1
    unread = repo.list_notifications(user_id=7, unread_only=True)
    assert [row["id"] for row in unread] == [1]

    marked = repo.mark_notification_read(
        1, user_id=7, read_at=datetime(2026, 9, 3, 18, 0, 0)
    )
    assert marked is not None
    assert marked["read_at"] == "2026-09-03T18:00:00"
    assert repo.count_unread_notifications(user_id=7) == 0
    assert repo.mark_notification_read(1, user_id=3, read_at=datetime.now()) is None
