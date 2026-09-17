"""GPS-018/019/020/021: API очередей GUID и проверка драфта."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from core.nomenclature_guid import ensure_schema, upsert
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient
from tests.helpers.production_api_fixtures import VALID_APP_SECRET_KEY, session_cookie

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
        "id": 3,
        "username": "manager_a",
        "role": "manager",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
    {
        "id": 4,
        "username": "econ",
        "role": "economist",
        "manager_id": None,
        "is_active": 1,
        "session_version": 0,
        "created_at": "2026-01-01 00:00:00",
    },
]


@pytest.fixture()
def env(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
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
        upsert(conn, "pile", "С30.30-3", match_status="missing")
        conn.commit()
    finally:
        conn.close()
    return pb_db


@pytest.fixture()
def client(env: Path):
    from app.main import create_app

    del env
    return CsrfAwareTestClient(create_app())


def test_tasks_economist_ok_manager_forbidden(client) -> None:
    ok = client.get(
        "/api/v1/nomenclature/tasks",
        cookies=session_cookie(4, "economist", "econ"),
    )
    assert ok.status_code == 200, ok.text
    body = ok.json()
    assert "to_create_1c" in body
    assert "to_price" in body
    assert "duplicates" in body
    assert any(item["mark"] == "С30.30-3" for item in body["to_create_1c"])

    forbidden = client.get(
        "/api/v1/nomenclature/tasks",
        cookies=session_cookie(3, "manager", "manager_a"),
    )
    assert forbidden.status_code == 403


def _econ() -> dict[str, str]:
    return session_cookie(4, "economist", "econ")


def _manager() -> dict[str, str]:
    return session_cookie(3, "manager", "manager_a")


def _admin() -> dict[str, str]:
    return session_cookie(1, "admin", "admin")


def test_price_resolve_zero_is_400(client) -> None:
    response = client.post(
        "/api/v1/nomenclature/price-queue/resolve",
        json={"guid": "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa", "price": 0},
        cookies=_econ(),
    )
    assert response.status_code in {400, 404}
    if response.status_code == 400:
        assert "цена" in response.json()["detail"]


def test_manager_forbidden_on_duplicates_and_resolve(client) -> None:
    for method, url, json_body in (
        ("GET", "/api/v1/nomenclature/duplicates", None),
        (
            "POST",
            "/api/v1/nomenclature/duplicates/resolve",
            {
                "scope": "nonplate",
                "key": "С50.30-8",
                "chosen_guid": "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
                "note": "x",
                "product_kind": "pile",
            },
        ),
        (
            "POST",
            "/api/v1/nomenclature/price-queue/resolve",
            {"guid": "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa", "price": 1500},
        ),
    ):
        response = client.request(method, url, json=json_body, cookies=_manager())
        assert response.status_code == 403, (url, response.text)


def test_economist_and_admin_get_duplicates_200(client) -> None:
    for cookies in (_econ(), _admin()):
        response = client.get("/api/v1/nomenclature/duplicates", cookies=cookies)
        assert response.status_code == 200, response.text
        assert "duplicates" in response.json()


def test_tasks_with_prays_plity_dupes_without_choice_table(client, env: Path) -> None:
    guid_a = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"
    guid_b = "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb"
    conn = sqlite3.connect(env)
    try:
        conn.execute(
            """
            CREATE TABLE prays_plity (
                "Уникальный идентификатор (Номенклатура)" TEXT,
                "Товар" TEXT
            )
            """
        )
        conn.execute("INSERT INTO prays_plity VALUES (?, ?)", (guid_a, "ПБ 60-12-8"))
        conn.execute("INSERT INTO prays_plity VALUES (?, ?)", (guid_b, "ПБ 60-12-8"))
        conn.commit()
        names = {
            row[0]
            for row in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")
        }
        assert "plate_guid_choice" not in names
    finally:
        conn.close()

    response = client.get("/api/v1/nomenclature/tasks", cookies=_econ())
    assert response.status_code == 200, response.text
    dups = [item for item in response.json()["duplicates"] if item["scope"] == "plate"]
    assert dups
    assert dups[0]["key"] == "ПБ 60-12-8"
    assert len(dups[0]["candidates"]) == 2


def test_resolve_duplicate_gone_guid_is_409(client, env: Path) -> None:
    from core.duplicate_candidates import upsert_candidates

    guid_a = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"
    guid_b = "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb"
    gone = "cccccccc-cccc-cccc-cccccccccccc"
    conn = sqlite3.connect(env)
    try:
        upsert(conn, "pile", "С50.30-8", match_status="ambiguous")
        upsert_candidates(
            conn,
            "pile",
            "С50.30-8",
            [(guid_a, "Сваи С 50.30-8", 1000.0), (guid_b, "Сваи С 50.30-8", 1100.0)],
        )
        conn.commit()
    finally:
        conn.close()

    response = client.post(
        "/api/v1/nomenclature/duplicates/resolve",
        json={
            "scope": "nonplate",
            "key": "С50.30-8",
            "chosen_guid": gone,
            "note": "бухгалтерия",
            "product_kind": "pile",
        },
        cookies=_econ(),
    )
    assert response.status_code == 409, response.text
    assert "исчез" in response.json()["detail"]

    still = client.get("/api/v1/nomenclature/duplicates", cookies=_econ())
    assert still.status_code == 200
    keys = {item["key"] for item in still.json()["duplicates"]}
    assert "С50.30-8" in keys


def test_price_resolve_economist_200_writes_auto_guid(client, env: Path) -> None:
    from core.nomenclature_guid import get_by_mark
    from core.pile_price_db import get_pile_price
    from core.price_queue_db import upsert_unmatched

    guid = "99999999-9999-9999-9999-999999999999"
    conn = sqlite3.connect(env)
    try:
        upsert_unmatched(conn, [(guid, "С999.30-1", "pile", "Сваи С 999.30-1")])
        conn.commit()
    finally:
        conn.close()

    forbidden = client.post(
        "/api/v1/nomenclature/price-queue/resolve",
        json={"guid": guid, "price": 2500},
        cookies=_manager(),
    )
    assert forbidden.status_code == 403

    response = client.post(
        "/api/v1/nomenclature/price-queue/resolve",
        json={"guid": guid, "price": 2500},
        cookies=_econ(),
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["state"] == "resolved"
    assert body["price"] == 2500
    assert get_pile_price("С999.30-1", "B25", str(env)) == 2500.0
    conn = sqlite3.connect(env)
    try:
        stored = get_by_mark(conn, "pile", "С999.30-1")
    finally:
        conn.close()
    assert stored is not None
    assert stored.match_status == "auto"
    assert stored.guid_1c == guid


def test_guid_check_lists_missing_for_manager(
    client, monkeypatch: pytest.MonkeyPatch
) -> None:
    from app.dependencies.services import get_commercial_workflow_service

    class _FakeWorkflow:
        def get_draft_details(self, draft_id: str) -> dict:
            del draft_id
            return {
                "order_data": [
                    {"product_type": "piles", "mark": "С30.30-3", "qty": 2},
                ]
            }

    monkeypatch.setattr(
        "app.services.draft_store.DraftStore.load_raw_json",
        lambda self, draft_id: {"metadata": {"owner_user_id": 3}},
    )
    client.app.dependency_overrides[get_commercial_workflow_service] = (
        lambda: _FakeWorkflow()
    )
    try:
        response = client.get(
            "/api/v1/commercial/drafts/draft-guid-1/guid-check",
            cookies=_manager(),
        )
    finally:
        client.app.dependency_overrides.pop(get_commercial_workflow_service, None)

    assert response.status_code == 200, response.text
    missing = response.json()["missing"]
    assert any(item["mark"] == "С30.30-3" for item in missing)


def test_guid_check_empty_when_guid_present(
    client, env: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from app.dependencies.services import get_commercial_workflow_service

    conn = sqlite3.connect(env)
    try:
        upsert(
            conn,
            "pile",
            "С30.30-3",
            guid_1c="11111111-1111-1111-1111-111111111111",
            match_status="auto",
        )
        conn.commit()
    finally:
        conn.close()

    class _FakeWorkflow:
        def get_draft_details(self, draft_id: str) -> dict:
            del draft_id
            return {
                "order_data": [
                    {"product_type": "piles", "mark": "С30.30-3", "qty": 1},
                ]
            }

    monkeypatch.setattr(
        "app.services.draft_store.DraftStore.load_raw_json",
        lambda self, draft_id: {"metadata": {"owner_user_id": 3}},
    )
    client.app.dependency_overrides[get_commercial_workflow_service] = (
        lambda: _FakeWorkflow()
    )
    try:
        response = client.get(
            "/api/v1/commercial/drafts/draft-guid-ready/guid-check",
            cookies=_manager(),
        )
    finally:
        client.app.dependency_overrides.pop(get_commercial_workflow_service, None)

    assert response.status_code == 200, response.text
    assert response.json()["missing"] == []
