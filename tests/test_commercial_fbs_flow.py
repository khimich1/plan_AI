"""FBS-601: API flow for FBS commercial drafts."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest
import sqlite3
from fastapi.testclient import TestClient

from core.kp_db_schema import init_schema
from core.kp_persistence_service import KpPersistenceService

from tests.helpers.csrf import CsrfAwareTestClient

from app.core.settings import get_settings
from app.main import create_app
from tests.helpers.auth_fixtures import patch_auth_users
from app.security.session import create_session_token


def _write_sample_fbs_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 7.5, 20, 22.5, 25],
        [1, "ФБС 9.3.6-Т", 1640.75, 1731.47, 1759.90, 1788.33],
        [2, "ФБС 12.4.6-Т", 2683.65, 2848.31, 2899.91, 2951.52],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)


def _setup_fbs_env(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> Path:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    monkeypatch.setenv("DRAFTS_DIR", str(tmp_path / "drafts"))
    monkeypatch.setenv("OUTPUTS_DIR", str(tmp_path / "outputs"))
    (tmp_path / "drafts").mkdir(exist_ok=True)
    (tmp_path / "outputs").mkdir(exist_ok=True)

    xlsx = tmp_path / "fbs.xlsx"
    pb_db = tmp_path / "pb.db"
    _write_sample_fbs_xlsx(xlsx)

    from core.fbs_price_db import import_fbs_prices_from_xlsx

    import_fbs_prices_from_xlsx(str(xlsx), str(pb_db))

    import core.commercial_offer as commercial_offer
    import core.commercial_offer_xlsx as commercial_offer_xlsx
    import app.services.commercial_workflow_service as commercial_workflow_service
    import app.services.commercial_calculation_service as commercial_calculation_service

    monkeypatch.setattr(commercial_offer, "DB_PATH", str(pb_db))
    monkeypatch.setattr(commercial_offer_xlsx, "DB_PATH", str(pb_db))
    monkeypatch.setattr(commercial_workflow_service, "DB_PATH", str(pb_db))
    monkeypatch.setattr(commercial_calculation_service, "DB_PATH", str(pb_db))
    get_settings.cache_clear()
    return pb_db


@pytest.fixture()
def client(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> TestClient:
    _setup_fbs_env(monkeypatch, tmp_path)
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


def _mock_manager_lookup(monkeypatch: pytest.MonkeyPatch) -> None:
    from app.repositories.manager_repository import ManagerRepository

    def fake_get_manager(self, manager_id: int):
        if manager_id == 1:
            return {
                "id": 1,
                "fio": "Tester",
                "contact_number": "+79990001122",
                "email": "tester@test.local",
            }
        return None

    monkeypatch.setattr(ManagerRepository, "get_manager", fake_get_manager)


def test_create_fbs_draft_with_text(client: TestClient, auth_cookie: dict[str, str]) -> None:
    response = client.post(
        "/api/v1/commercial/drafts",
        data={
            "product_type": "fbs",
            "text": "ФБС 9.3.6-Т 2\nФБС 12.4.6-Т B20 1",
        },
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["metadata"]["product_type"] == "fbs"
    assert len(body["order_data"]) == 2
    assert body["order_data"][0]["product_kind"] == "fbs"
    assert body["order_data"][0]["unit_price"] is not None
    marks = {row["mark"] for row in body["order_data"]}
    assert "ФБС 9.3.6-Т" in marks
    assert body["wizard_state"]["current_step"] == "fbs"


def test_bulk_grade_applies_to_all(client: TestClient, auth_cookie: dict[str, str]) -> None:
    create = client.post(
        "/api/v1/commercial/drafts",
        data={
            "product_type": "fbs",
            "text": "ФБС 9.3.6-Т B25 2\nФБС 12.4.6-Т B25 1",
        },
    )
    assert create.status_code == 200, create.text
    draft_id = create.json()["draft_id"]

    grades = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/fbs/grades",
        json={"concrete_grade": "B7_5"},
    )
    assert grades.status_code == 200, grades.text
    body = grades.json()
    assert all(row["concrete_grade"] == "B7_5" for row in body["order_data"])
    assert not body["metadata"].get("grade_bulk_skipped_marks")


def test_fbs_calculate_and_files(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _mock_manager_lookup(monkeypatch)
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "fbs", "text": "ФБС 9.3.6-Т B25 2"},
    )
    draft_id = create.json()["draft_id"]
    assert create.json()["order_data"][0]["mark"] == "ФБС 9.3.6-Т"

    meta = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={
            "manager_id": 1,
            "client_name": "ООО ФБС",
            "discount_percent": 0,
            "conditions_mode": "standard",
        },
    )
    assert meta.status_code == 200, meta.text

    calc = client.post(f"/api/v1/commercial/drafts/{draft_id}/calculate")
    assert calc.status_code == 200, calc.text
    assert calc.json()["wizard_state"]["current_step"] == "result"

    files = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/generate-files",
        json={"file_types": ["pdf", "xlsx", "breakdown", "schema"]},
    )
    assert files.status_code == 200, files.text
    kinds = {item["kind"] for item in files.json()["files"]}
    assert kinds == {"pdf", "xlsx"}


def test_fbs_persist_mark_as_typed(tmp_path: Path) -> None:
    db_path = str(tmp_path / "plita.db")
    init_schema(db_path)
    KpPersistenceService.save_kp_to_db(
        "05.08.2026",
        [
            {
                "product_kind": "fbs",
                "name": "ФБС 9.3.6-Т",
                "mark": "ФБС 9.3.6-Т",
                "concrete_grade": "B25",
                "qty": 2,
                "unit_price": 1788.33,
            }
        ],
        customer_name="ООО ФБС",
        status="в архиве",
        product_type="fbs",
        db_path=db_path,
    )
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute("SELECT product_type FROM kp_meta WHERE kp_id = 1")
        assert cur.fetchone()[0] == "fbs"
        cur.execute("SELECT mark, concrete_grade, qty FROM kp_fbs WHERE kp_id = 1")
        assert cur.fetchone() == ("ФБС 9.3.6-Т", "B25", 2)
        cur.execute("SELECT COUNT(*) FROM kp_piles WHERE kp_id = 1")
        assert cur.fetchone()[0] == 0
        cur.execute("SELECT COUNT(*) FROM kp_bridge_piles WHERE kp_id = 1")
        assert cur.fetchone()[0] == 0
