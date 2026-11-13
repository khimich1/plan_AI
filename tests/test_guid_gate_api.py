"""GPS-018/019/020/021: API очередей GUID и проверка драфта."""

from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest

from core.guid_demand import record_guid_demand
from core.guid_gate import HINT_CHOOSE, HINT_CREATE, HINT_CREATE_U, MissingItem, OrderLine
from core.guid_gate import REASON_AMBIGUOUS, REASON_DUP, REASON_MISSING, REASON_MISSING_U
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
    assert body["to_create_1c"] == []
    assert body["duplicates"] == []

    forbidden = client.get(
        "/api/v1/nomenclature/tasks",
        cookies=session_cookie(3, "manager", "manager_a"),
    )
    assert forbidden.status_code == 403


def test_tasks_list_open_demand_with_kp_ids(client, env: Path) -> None:
    conn = sqlite3.connect(env)
    try:
        record_guid_demand(
            conn,
            12,
            [
                MissingItem(
                    line=OrderLine("pile", "С70.35-9у"),
                    reason=REASON_MISSING_U,
                    action_hint=HINT_CREATE_U,
                )
            ],
        )
        record_guid_demand(
            conn,
            8,
            [
                MissingItem(
                    line=OrderLine("pile", "С99.99-1"),
                    reason=REASON_MISSING,
                    action_hint=HINT_CREATE,
                )
            ],
        )
        conn.commit()
    finally:
        conn.close()

    body = client.get("/api/v1/nomenclature/tasks", cookies=_econ()).json()
    by_mark = {item["mark"]: item for item in body["to_create_1c"]}
    assert "С30.30-3" not in by_mark
    assert by_mark["С70.35-9у"]["kp_ids"] == [12]
    assert by_mark["С70.35-9у"]["field"] == "guid_1c_u"
    assert by_mark["С99.99-1"]["kp_ids"] == [8]


def test_open_duplicate_without_candidates_is_hidden(client, env: Path) -> None:
    conn = sqlite3.connect(env)
    try:
        record_guid_demand(
            conn,
            4,
            [
                MissingItem(
                    line=OrderLine("pile", "С50.30-8"),
                    reason=REASON_AMBIGUOUS,
                    action_hint=HINT_CHOOSE,
                )
            ],
        )
        conn.commit()
    finally:
        conn.close()

    body = client.get("/api/v1/nomenclature/tasks", cookies=_econ()).json()
    assert body["duplicates"] == []


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
    assert dups == []


def test_tasks_show_plate_dup_only_with_demand(client, env: Path) -> None:
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
        record_guid_demand(
            conn,
            18,
            [
                MissingItem(
                    line=OrderLine("plate", "ПБ 60-12-8"),
                    reason=REASON_DUP,
                    action_hint=HINT_CHOOSE,
                )
            ],
        )
        conn.commit()
    finally:
        conn.close()

    response = client.get("/api/v1/nomenclature/tasks", cookies=_econ())
    assert response.status_code == 200, response.text
    dups = [item for item in response.json()["duplicates"] if item["scope"] == "plate"]
    assert len(dups) == 1
    assert dups[0]["key"] == "ПБ 60-12-8"
    assert len(dups[0]["candidates"]) == 2
    assert dups[0]["kp_ids"] == [18]


def test_pile_u_demand_joins_candidates_by_stripped_mark(client, env: Path) -> None:
    from core.duplicate_candidates import upsert_candidates

    guid_a = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"
    guid_b = "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb"
    conn = sqlite3.connect(env)
    try:
        upsert(conn, "pile", "С70.35-9", match_status="ambiguous")
        upsert_candidates(
            conn,
            "pile",
            "С70.35-9",
            [(guid_a, "Сваи С 70.35-9", 1000.0), (guid_b, "Сваи С 70.35-9", 1100.0)],
        )
        record_guid_demand(
            conn,
            12,
            [
                MissingItem(
                    line=OrderLine("pile", "С70.35-9у"),
                    reason=REASON_AMBIGUOUS,
                    action_hint=HINT_CHOOSE,
                )
            ],
        )
        conn.commit()
    finally:
        conn.close()

    response = client.get("/api/v1/nomenclature/tasks", cookies=_econ())
    assert response.status_code == 200, response.text
    dups = [item for item in response.json()["duplicates"] if item["key"] == "С70.35-9"]
    assert len(dups) == 1
    assert dups[0]["kp_ids"] == [12]
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
        record_guid_demand(
            conn,
            12,
            [
                MissingItem(
                    line=OrderLine("pile", "С50.30-8"),
                    reason=REASON_AMBIGUOUS,
                    action_hint=HINT_CHOOSE,
                )
            ],
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


# --- NG-103: особые карточки видны в очереди и в guid-check -----------------------


def _seed_special_cards(pb: Path) -> None:
    conn = sqlite3.connect(pb)
    try:
        conn.execute(
            'CREATE TABLE IF NOT EXISTS prays_plity ('
            '"Уникальный идентификатор (Номенклатура)" TEXT, "Товар" TEXT)'
        )
        conn.executemany(
            "INSERT INTO prays_plity VALUES (?, ?)",
            [
                ("ad13b504-96c5-11eb-8f7e-04d9f5387845", "Плиты ПБ 51-12-12п Пв-7.1"),
                ("b3596b8f-96c5-11eb-8f7e-04d9f5387845", "Плиты ПБ 51-12-12п Пв-7.2"),
            ],
        )
        conn.commit()
    finally:
        conn.close()


def test_tasks_hint_lists_special_cards_for_variants(client, env: Path) -> None:
    from core.guid_gate import check_invoice_guids

    _seed_special_cards(env)
    conn = sqlite3.connect(env)
    try:
        report = check_invoice_guids([OrderLine("plate", "Плиты ПБ 51-12-12п")], conn)
        record_guid_demand(conn, 206, report.missing)
        conn.commit()
    finally:
        conn.close()

    body = client.get("/api/v1/nomenclature/tasks", cookies=_econ()).json()

    item = next(i for i in body["to_create_1c"] if i["mark"] == "Плиты ПБ 51-12-12п")
    assert item["field"] == "guid_1c"
    assert "особые" in item["hint"] and "Пв-7.1" in item["hint"]
    assert item["kp_ids"] == [206]


def test_guid_check_reports_variants_reason(client, env: Path, monkeypatch) -> None:
    from app.dependencies.services import get_commercial_workflow_service

    _seed_special_cards(env)

    class _FakeWorkflow:
        def get_draft_details(self, draft_id: str) -> dict:
            del draft_id
            return {"order_data": [{"product_type": "plates", "name": "Плиты ПБ 51-12-12п"}]}

    monkeypatch.setattr(
        "app.services.draft_store.DraftStore.load_raw_json",
        lambda self, draft_id: {"metadata": {"owner_user_id": 3}},
    )
    client.app.dependency_overrides[get_commercial_workflow_service] = lambda: _FakeWorkflow()
    try:
        response = client.get(
            "/api/v1/commercial/drafts/draft-variants/guid-check", cookies=_manager()
        )
    finally:
        client.app.dependency_overrides.pop(get_commercial_workflow_service, None)

    assert response.status_code == 200, response.text
    missing = response.json()["missing"]
    assert missing[0]["reason"] == "стандартной карточки нет, есть особые"


# --- NG-107: объяснение и поиск (менеджер и экономист) -----------------------------

_EXPLAIN_G1 = "11111111-1111-1111-1111-111111111111"
_EXPLAIN_G2 = "22222222-2222-2222-2222-222222222222"
_EXPLAIN_G3 = "33333333-3333-3333-3333-333333333333"
_TYPED_PLATE = "Плиты ПБ 22,80-12-8п"


def _seed_explain_cards(pb: Path) -> None:
    conn = sqlite3.connect(pb)
    try:
        conn.execute(
            'CREATE TABLE IF NOT EXISTS prays_plity ('
            '"Уникальный идентификатор (Номенклатура)" TEXT, "Товар" TEXT)'
        )
        conn.executemany(
            "INSERT INTO prays_plity VALUES (?, ?)",
            [
                (_EXPLAIN_G1, "Плиты ПБ 22,8-12-8п"),
                (_EXPLAIN_G2, "Плиты ПБ 22,8-10,2-8п"),
                (_EXPLAIN_G3, "Плиты ПБ 22,8-10,2-8п"),
            ],
        )
        conn.commit()
    finally:
        conn.close()


@pytest.fixture()
def own_draft(client, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setattr(
        "app.services.draft_store.DraftStore.load_raw_json",
        lambda self, draft_id: {"metadata": {"owner_user_id": 3}},
    )
    return client


def _by_mark(items: list[dict]) -> dict[str, dict]:
    return {item["mark"]: item for item in items}


def test_manager_explain_selectable_only_with_single_guid(own_draft, env: Path) -> None:
    _seed_explain_cards(env)

    response = own_draft.get(
        "/api/v1/commercial/drafts/d-1/guid-explain",
        params={"product_kind": "plate", "mark": _TYPED_PLATE},
        cookies=_manager(),
    )

    assert response.status_code == 200, response.text
    body = response.json()
    assert body["reason"] == ""
    same = _by_mark(body["same_cards"])
    assert same["Плиты ПБ 22,8-12-8п"]["guid_count"] == 1
    assert same["Плиты ПБ 22,8-12-8п"]["label"] == "лишний ноль"
    assert same["Плиты ПБ 22,8-12-8п"]["selectable"] is True
    dup = _by_mark(body["slice"])["Плиты ПБ 22,8-10,2-8п"]
    assert dup["guid_count"] == 2
    assert dup["selectable"] is False
    assert all(item["selectable"] == (item["guid_count"] == 1) for item in body["slice"])


def test_manager_cannot_become_economist_via_role_param(own_draft, env: Path) -> None:
    _seed_explain_cards(env)

    response = own_draft.get(
        "/api/v1/commercial/drafts/d-1/guid-explain",
        params={"product_kind": "plate", "mark": _TYPED_PLATE, "role": "economist"},
        cookies=_manager(),
    )

    body = response.json()
    assert _by_mark(body["slice"])["Плиты ПБ 22,8-10,2-8п"]["selectable"] is False


def test_economist_explain_returns_two_named_cards_for_one_plate(client, env: Path) -> None:
    _seed_explain_cards(env)

    response = client.get(
        "/api/v1/nomenclature/guid-explain",
        params={"product_kind": "plate", "mark": _TYPED_PLATE},
        cookies=_econ(),
    )

    assert response.status_code == 200, response.text
    dup = _by_mark(response.json()["slice"])["Плиты ПБ 22,8-10,2-8п"]
    assert dup["guid_count"] == 2
    assert dup["selectable"] is True
    assert [(card["name"], card["guid"]) for card in dup["cards"]] == [
        ("Плиты ПБ 22,8-10,2-8п", _EXPLAIN_G2),
        ("Плиты ПБ 22,8-10,2-8п", _EXPLAIN_G3),
    ]
    same = response.json()["same_cards"][0]
    assert same["guid_count"] == 1
    assert same["cards"] == []


def test_manager_explain_lists_two_cards_but_they_stay_unselectable(
    own_draft, env: Path
) -> None:
    _seed_explain_cards(env)

    response = own_draft.get(
        "/api/v1/commercial/drafts/d-1/guid-explain",
        params={"product_kind": "plate", "mark": _TYPED_PLATE},
        cookies=_manager(),
    )

    assert response.status_code == 200, response.text
    dup = _by_mark(response.json()["slice"])["Плиты ПБ 22,8-10,2-8п"]
    assert dup["selectable"] is False
    assert len(dup["cards"]) == 2


def test_economist_explain_selects_zero_and_two(client, env: Path) -> None:
    _seed_explain_cards(env)

    response = client.get(
        "/api/v1/nomenclature/guid-explain",
        params={"product_kind": "plate", "mark": _TYPED_PLATE},
        cookies=_econ(),
    )

    assert response.status_code == 200, response.text
    body = response.json()
    assert body["same_cards"][0]["selectable"] is True
    assert _by_mark(body["slice"])["Плиты ПБ 22,8-10,2-8п"]["selectable"] is True


def test_economist_explain_zero_guid_row_is_selectable(client, env: Path) -> None:
    conn = sqlite3.connect(env)
    try:
        upsert(conn, "pile", "С30.30-4", match_status="missing")
        conn.commit()
    finally:
        conn.close()

    response = client.get(
        "/api/v1/nomenclature/guid-explain",
        params={"product_kind": "pile", "mark": "С30.30-3"},
        cookies=_econ(),
    )

    assert response.status_code == 200, response.text
    row = _by_mark(response.json()["slice"])["С30.30-4"]
    assert row["guid_count"] == 0
    assert row["selectable"] is True


def test_manager_forbidden_on_economist_explain_and_search(client, env: Path) -> None:
    _seed_explain_cards(env)

    explain = client.get(
        "/api/v1/nomenclature/guid-explain",
        params={"product_kind": "plate", "mark": _TYPED_PLATE},
        cookies=_manager(),
    )
    search = client.get(
        "/api/v1/nomenclature/guid-search",
        params={"product_kind": "plate", "q": "ПБ 22,8"},
        cookies=_manager(),
    )

    assert explain.status_code == 403
    assert search.status_code == 403


def test_foreign_draft_explain_and_search_denied(
    client, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(
        "app.services.draft_store.DraftStore.load_raw_json",
        lambda self, draft_id: {"metadata": {"owner_user_id": 99}},
    )

    explain = client.get(
        "/api/v1/commercial/drafts/d-foreign/guid-explain",
        params={"product_kind": "plate", "mark": _TYPED_PLATE},
        cookies=_manager(),
    )
    search = client.get(
        "/api/v1/commercial/drafts/d-foreign/guid-search",
        params={"product_kind": "plate", "q": "ПБ 22,8"},
        cookies=_manager(),
    )

    assert explain.status_code in (403, 404)
    assert search.status_code in (403, 404)


def test_search_by_role_and_limits(own_draft, env: Path) -> None:
    _seed_explain_cards(env)

    manager = own_draft.get(
        "/api/v1/commercial/drafts/d-1/guid-search",
        params={"product_kind": "plate", "q": "ПБ 22,8"},
        cookies=_manager(),
    )
    economist = own_draft.get(
        "/api/v1/nomenclature/guid-search",
        params={"product_kind": "plate", "q": "ПБ 22,8"},
        cookies=_econ(),
    )
    short = own_draft.get(
        "/api/v1/nomenclature/guid-search",
        params={"product_kind": "plate", "q": "ПБ"},
        cookies=_econ(),
    )

    manager_items = _by_mark(manager.json()["items"])
    assert manager_items["Плиты ПБ 22,8-12-8п"]["selectable"] is True
    assert manager_items["Плиты ПБ 22,8-10,2-8п"]["selectable"] is False
    assert all(item["selectable"] for item in economist.json()["items"])
    assert len(economist.json()["items"]) == 2
    assert short.status_code == 200
    assert short.json()["items"] == []


def test_explain_unknown_kind_is_422(client) -> None:
    response = client.get(
        "/api/v1/nomenclature/guid-explain",
        params={"product_kind": "unicorn", "mark": "X-1"},
        cookies=_econ(),
    )

    assert response.status_code == 422


def test_manager_forbidden_on_guid_replace(client) -> None:
    body = {
        "product_kind": "pile",
        "mark": "C14-40T4",
        "target_mark": "С14-40T4",
    }
    preview = client.post(
        "/api/v1/nomenclature/guid-replace/preview", json=body, cookies=_manager()
    )
    write = client.post("/api/v1/nomenclature/guid-replace", json=body, cookies=_manager())
    assert preview.status_code == 403
    assert write.status_code == 403


def test_guid_replace_preview_and_write_and_two_cards(client, env: Path) -> None:
    import json

    from app.core.settings import get_settings
    from core.guid_demand import ensure_schema as ensure_demand_schema
    from core.kp_db_schema import init_schema
    from core.kp_persistence_service import KpPersistenceService
    from core.pile_price_db import init_pile_prices_schema

    pb = str(env)
    plita = str(get_settings().plita_db_path)
    init_schema(plita)
    init_pile_prices_schema(pb)
    with sqlite3.connect(pb) as conn:
        conn.execute(
            "INSERT OR REPLACE INTO pile_prices (mark, concrete_grade, price) VALUES (?, 'B25', 10000)",
            ("С14-40T4",),
        )
        ensure_demand_schema(conn)
        upsert(conn, "bridge_pile", "С15-35T7", match_status="ambiguous")
        conn.execute(
            "INSERT INTO guid_demand (product_kind, mark, field, reason, hint, state, "
            "kp_ids_json, opened_at) VALUES ('bridge_pile', 'C15-35T7', 'guid_1c', "
            "'неоднозначный GUID (дубль 1С)', 'выберите', 'open', '[1]', '2026-11-01T00:00:00Z')"
        )
        conn.commit()
    kp_id = KpPersistenceService.save_kp_to_db(
        "12.11.2026",
        [
            {
                "line_id": "ln_p",
                "product_type": "piles",
                "product_kind": "pile",
                "mark": "C14-40T4",
                "name": "C14-40T4",
                "concrete_grade": "B25",
                "qty": 2,
                "unit_price": 100.0,
            }
        ],
        customer_name="API",
        discount_percent=10,
        product_type="piles",
        status="в архиве",
        db_path=plita,
    )
    with sqlite3.connect(pb) as conn:
        conn.execute(
            "INSERT INTO guid_demand (product_kind, mark, field, reason, hint, state, "
            "kp_ids_json, opened_at) VALUES ('pile', 'C14-40T4', 'guid_1c', 'нет GUID 1С', "
            "'заведите', 'open', ?, '2026-11-01T00:00:00Z')",
            (json.dumps([kp_id]),),
        )
        conn.commit()

    denied = client.post(
        "/api/v1/nomenclature/guid-replace/preview",
        json={
            "product_kind": "bridge_pile",
            "mark": "C15-35T7",
            "target_mark": "С15-35T7",
        },
        cookies=_econ(),
    )
    assert denied.status_code == 400
    assert "двух" in denied.json()["detail"]

    body = {
        "product_kind": "pile",
        "mark": "C14-40T4",
        "target_mark": "С14-40T4",
    }
    preview = client.post(
        "/api/v1/nomenclature/guid-replace/preview", json=body, cookies=_econ()
    )
    assert preview.status_code == 200, preview.text
    item = preview.json()["items"][0]
    assert item["kp_id"] == kp_id
    assert item["qty"] == 2
    assert item["will"]["unit_discounted"] == pytest.approx(9000)
    assert item["was"]["total"] != item["will"]["total"]
    with sqlite3.connect(plita) as conn:
        untouched = conn.execute(
            "SELECT mark FROM kp_piles WHERE kp_id = ?", (kp_id,)
        ).fetchone()
    assert untouched is not None and untouched[0] == "C14-40T4"

    written = client.post("/api/v1/nomenclature/guid-replace", json=body, cookies=_econ())
    assert written.status_code == 200, written.text
    payload = written.json()
    assert payload["written"][0]["kp_id"] == kp_id
    assert payload["skipped"] == []
    with sqlite3.connect(plita) as conn:
        mark = conn.execute(
            "SELECT mark FROM kp_piles WHERE kp_id = ?", (kp_id,)
        ).fetchone()
        actor = conn.execute(
            "SELECT changed_by, source FROM kp_mark_history"
        ).fetchone()
    assert mark is not None and mark[0] == "С14-40T4"
    assert actor == ("econ", "replace_economist")
