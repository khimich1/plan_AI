"""API flow for stair-step commercial drafts."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest
import sqlite3
from fastapi.testclient import TestClient

from core.kp_db_schema import init_schema

from tests.helpers.csrf import CsrfAwareTestClient

from app.core.settings import get_settings
from app.main import create_app
from tests.helpers.auth_fixtures import patch_auth_users
from app.security.session import create_session_token


def _write_sample_step_xlsx(path: Path) -> None:
    rows = [
        [None, None, None],
        [None, None, None],
        [None, "Наименование", 15],
        [1, "Лестничные ступени ЛС11", 1409.908359678],
        [2, "Лестничные ступени ЛС14-1лев", 1815.586530576],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)


def _setup_step_env(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> Path:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    monkeypatch.setenv("DRAFTS_DIR", str(tmp_path / "drafts"))
    monkeypatch.setenv("OUTPUTS_DIR", str(tmp_path / "outputs"))
    (tmp_path / "drafts").mkdir(exist_ok=True)
    (tmp_path / "outputs").mkdir(exist_ok=True)

    step_xlsx = tmp_path / "steps.xlsx"
    pb_db = tmp_path / "pb.db"
    _write_sample_step_xlsx(step_xlsx)

    from core.step_price_db import import_step_prices_from_xlsx
    import_step_prices_from_xlsx(str(step_xlsx), str(pb_db))

    import core.commercial_offer as commercial_offer
    import core.commercial_offer_xlsx as commercial_offer_xlsx
    import app.services.commercial_workflow_service as workflow_mod
    import app.services.commercial_calculation_service as calc_mod
    import app.services.commercial_export_service as export_mod

    monkeypatch.setattr(commercial_offer, "DB_PATH", str(pb_db))
    monkeypatch.setattr(commercial_offer_xlsx, "DB_PATH", str(pb_db))
    monkeypatch.setattr(workflow_mod, "DB_PATH", str(pb_db))
    monkeypatch.setattr(calc_mod, "DB_PATH", str(pb_db))
    monkeypatch.setattr(export_mod, "DB_PATH", str(pb_db))
    get_settings.cache_clear()
    return pb_db


@pytest.fixture()
def client(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> TestClient:
    _setup_step_env(monkeypatch, tmp_path)
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


def test_create_step_draft_with_text(client: TestClient, auth_cookie: dict[str, str]) -> None:
    response = client.post(
        "/api/v1/commercial/drafts",
        data={
            "product_type": "steps",
            "text": "ЛС11 5\nЛС14-1лев 2",
        },
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["metadata"]["product_type"] == "steps"
    assert len(body["order_data"]) == 2
    assert body["order_data"][0]["product_kind"] == "step"
    assert body["order_data"][0]["unit_price"] is not None
    assert body["wizard_state"]["current_step"] == "steps"


def test_update_step_draft_replace(client: TestClient, auth_cookie: dict[str, str]) -> None:
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "steps", "text": "ЛС11 1"},
    )
    draft_id = create.json()["draft_id"]

    response = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/steps",
        data={"mode": "replace", "text": "ЛС14-1лев 3"},
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert len(body["order_data"]) == 1
    assert body["order_data"][0]["mark"] == "ЛС14-1ЛЕВ"
    assert body["order_data"][0]["qty"] == 3


def test_apply_ai_steps_endpoint(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from app.services.commercial_workflow_service import CommercialWorkflowService

    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "steps", "text": "ЛС11 2"},
    )
    draft_id = create.json()["draft_id"]

    async def fake_apply(self, draft_id_arg: str, **kwargs):
        draft = self.get_draft_details(draft_id_arg)
        draft["metadata"]["source_type"] = "ai"
        draft["metadata"]["ai_applied"] = True
        draft["metadata"]["input_text"] = "ЛС14-1лев 3"
        return draft

    monkeypatch.setattr(CommercialWorkflowService, "apply_ai_steps_instruction", fake_apply)

    response = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/steps/ai",
        data={"instruction": "замени на 14-1лев qty 3"},
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["metadata"]["source_type"] == "ai"
    assert body["metadata"]["ai_applied"] is True


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


def test_step_draft_calculate_and_generate_files(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _mock_manager_lookup(monkeypatch)
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "steps", "text": "ЛС11 2"},
    )
    draft_id = create.json()["draft_id"]

    meta = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={
            "manager_id": 1,
            "client_name": "ООО Ступени",
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


def test_step_draft_save_to_archive(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    _setup_step_env(monkeypatch, tmp_path)
    plita_db = tmp_path / "plita.db"
    init_schema(str(plita_db))
    with sqlite3.connect(str(plita_db)) as conn:
        conn.execute(
            "INSERT INTO managers (id, fio, contact_number, email) "
            "VALUES (1, 'Tester', '+79990001122', 'tester@test.local')"
        )
        conn.commit()
    monkeypatch.setenv("PLITA_DB_PATH", str(plita_db))
    get_settings.cache_clear()
    _mock_manager_lookup(monkeypatch)
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
    token = create_session_token({"id": 1, "username": "tester", "role": "admin"}, ttl_seconds=300)
    client.cookies.set("app_session", token)
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "steps", "text": "ЛС11 2"},
    )
    draft_id = create.json()["draft_id"]

    client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={
            "manager_id": 1,
            "client_name": "ООО Ступени",
            "conditions_mode": "standard",
        },
    )
    client.post(f"/api/v1/commercial/drafts/{draft_id}/calculate")

    save = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/save",
        json={"mode": "archive", "execution_terms_input": ""},
    )
    assert save.status_code == 200, save.text
    kp_id = save.json()["saved_offer"]["kp_id"]
    assert kp_id == 1

    with sqlite3.connect(str(plita_db)) as conn:
        cur = conn.cursor()
        cur.execute("SELECT product_type FROM kp_meta WHERE kp_id = ?", (kp_id,))
        assert cur.fetchone()[0] == "steps"
        cur.execute("SELECT mark FROM kp_steps WHERE kp_id = ?", (kp_id,))
        assert cur.fetchone()[0] == "ЛС11"


def test_step_xlsx_contains_mark_only() -> None:
    from core.commercial_offer_xlsx import generate_commercial_offer_xlsx

    order_data = [
        {
            "product_kind": "step",
            "name": "ЛС14-1ЛЕВ",
            "mark": "ЛС14-1ЛЕВ",
            "qty": 2,
            "unit_price": 1815.59,
        }
    ]
    xlsx = generate_commercial_offer_xlsx(
        order_data,
        offer_number="3",
        offer_date="05.08.2026",
        customer_name="Test",
        kp_db_id=3,
    )
    df = pd.read_excel(xlsx, sheet_name="КП", header=None)
    flat = " ".join(str(v) for v in df.values.flatten() if pd.notna(v))
    assert "ЛС14-1ЛЕВ" in flat
    assert "B25" not in flat


def test_order_data_from_kp_steps() -> None:
    from core.kp_order_data import order_data_from_kp_info

    kp_info = {
        "product_type": "steps",
        "discount_percent": 0,
        "steps": [
            {
                "mark": "ЛС11",
                "qty": 4,
                "unit_price": 100.0,
                "discounted_price": 100.0,
            }
        ],
    }
    order_data = order_data_from_kp_info(kp_info)
    assert len(order_data) == 1
    assert order_data[0]["product_kind"] == "step"
    assert order_data[0]["mark"] == "ЛС11"
    assert order_data[0]["qty"] == 4


def test_step_draft_rejects_plates_endpoint(
    client: TestClient,
    auth_cookie: dict[str, str],
) -> None:
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "steps", "text": "ЛС11 2"},
    )
    draft_id = create.json()["draft_id"]

    response = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/plates",
        data={"mode": "replace", "text": "ПБ 78-12-8п 1"},
    )
    assert response.status_code == 400, response.text
