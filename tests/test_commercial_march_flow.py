"""MARCH-103: API flow for march commercial drafts."""

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


def _write_sample_march_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 15, 20, 22.5, 25, "30 на граните"],
        [1, "Лестничные марши 1ЛМ 27-11-14-4", 13993.72, 14150.79, 14271.10, 14391.41, 14639.53],
        [6, "Лестничные марши ЛМ 2,8", 15819.88, 16000.91, 16139.57, 16278.24, 16564.20],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)


def _setup_march_env(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> Path:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    monkeypatch.setenv("DRAFTS_DIR", str(tmp_path / "drafts"))
    monkeypatch.setenv("OUTPUTS_DIR", str(tmp_path / "outputs"))
    (tmp_path / "drafts").mkdir(exist_ok=True)
    (tmp_path / "outputs").mkdir(exist_ok=True)

    march_xlsx = tmp_path / "marches.xlsx"
    pb_db = tmp_path / "pb.db"
    _write_sample_march_xlsx(march_xlsx)

    from core.march_price_db import import_march_prices_from_xlsx
    import_march_prices_from_xlsx(str(march_xlsx), str(pb_db))

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
    _setup_march_env(monkeypatch, tmp_path)
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


def test_create_march_draft_with_text(client: TestClient, auth_cookie: dict[str, str]) -> None:
    response = client.post(
        "/api/v1/commercial/drafts",
        data={
            "product_type": "marches",
            "text": "1ЛМ 27-11-14-4 B25 5\nЛМ 2,8 B30 2",
        },
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["metadata"]["product_type"] == "marches"
    assert len(body["order_data"]) == 2
    assert body["order_data"][0]["product_kind"] == "march"
    assert body["order_data"][0]["unit_price"] is not None
    assert body["wizard_state"]["current_step"] == "marches"
    for row in body["order_data"]:
        assert str(row.get("line_id") or "").strip()
        assert row.get("product_type") == "marches"


def test_update_march_draft_replace(client: TestClient, auth_cookie: dict[str, str]) -> None:
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "marches", "text": "1ЛМ 27-11-14-4 B25 1"},
    )
    draft_id = create.json()["draft_id"]

    response = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/marches",
        data={"mode": "replace", "text": "ЛМ 2,8 B25 3"},
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert len(body["order_data"]) == 1
    assert body["order_data"][0]["mark"] == "ЛМ 2,8"
    assert body["order_data"][0]["qty"] == 3


def test_bulk_grade_single_line_no_duplicate(
    client: TestClient,
    auth_cookie: dict[str, str],
) -> None:
    """UI regression: apply class to all must not duplicate the create row."""
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "marches", "text": "1ЛМ 27-11-14-4 B25 1"},
    )
    assert create.status_code == 200, create.text
    draft_id = create.json()["draft_id"]
    assert len(create.json()["order_data"]) == 1

    response = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/marches/grades",
        json={"concrete_grade": "B20"},
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert len(body["order_data"]) == 1
    assert body["order_data"][0]["concrete_grade"] == "B20"
    assert body["order_data"][0]["mark"] == "1ЛМ 27-11-14-4"


def test_partition_treats_untyped_legacy_mono_as_same_type() -> None:
    """Old drafts without product_type still replace instead of duplicating."""
    from app.services.commercial_workflow_service import CommercialWorkflowService

    workflow = CommercialWorkflowService()
    previous = [
        {
            "product_kind": "march",
            "mark": "1ЛМ 27-11-14-4",
            "concrete_grade": "B25",
            "qty": 1,
        }
    ]
    others, same = workflow._partition_order_by_product_type(
        previous,
        product_type="marches",
    )
    assert others == []
    assert len(same) == 1
    composed = workflow._compose_order_data_for_product_update(
        previous_order_data=previous,
        new_type_lines=[
            {
                "line_id": "new-1",
                "product_type": "marches",
                "product_kind": "march",
                "mark": "1ЛМ 27-11-14-4",
                "concrete_grade": "B20",
                "qty": 1,
            }
        ],
        product_type="marches",
        mode="replace",
        merged_cycle_text=False,
    )
    assert len(composed) == 1
    assert composed[0]["concrete_grade"] == "B20"


def test_apply_ai_marches_endpoint(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from app.services.commercial_workflow_service import CommercialWorkflowService

    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "marches", "text": "1ЛМ 27-11-14-4 B25 2"},
    )
    draft_id = create.json()["draft_id"]

    async def fake_apply(self, draft_id_arg: str, **kwargs):
        draft = self.get_draft_details(draft_id_arg)
        draft["metadata"]["source_type"] = "ai"
        draft["metadata"]["ai_applied"] = True
        draft["metadata"]["input_text"] = "ЛМ 2,8 B25 3"
        return draft

    monkeypatch.setattr(CommercialWorkflowService, "apply_ai_marches_instruction", fake_apply)

    response = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/marches/ai",
        data={"instruction": "замени на ЛМ 2,8 qty 3"},
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


def test_update_march_grades_bulk(
    client: TestClient,
    auth_cookie: dict[str, str],
) -> None:
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "marches", "text": "1ЛМ 27-11-14-4 B25 2\nЛМ 2,8 B30 1"},
    )
    draft_id = create.json()["draft_id"]

    response = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/marches/grades",
        json={"concrete_grade": "B20"},
    )
    assert response.status_code == 200, response.text
    body = response.json()
    grades = {row["concrete_grade"] for row in body["order_data"]}
    assert grades == {"B20"}
    assert all(row["unit_price"] is not None for row in body["order_data"])


def test_march_draft_calculate_and_generate_files(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _mock_manager_lookup(monkeypatch)
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "marches", "text": "1ЛМ 27-11-14-4 B25 2"},
    )
    draft_id = create.json()["draft_id"]

    meta = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={
            "manager_id": 1,
            "client_name": "ООО Марши",
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


def test_march_draft_save_to_archive(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    _setup_march_env(monkeypatch, tmp_path)
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
        data={"product_type": "marches", "text": "1ЛМ 27-11-14-4 B25 2"},
    )
    draft_id = create.json()["draft_id"]

    client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={
            "manager_id": 1,
            "client_name": "ООО Марши",
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
        assert cur.fetchone()[0] == "marches"
        cur.execute("SELECT mark FROM kp_marches WHERE kp_id = ?", (kp_id,))
        assert cur.fetchone()[0] == "1ЛМ 27-11-14-4"


def test_march_xlsx_contains_mark_and_grade() -> None:
    from core.commercial_offer_xlsx import generate_commercial_offer_xlsx

    order_data = [
        {
            "product_kind": "march",
            "name": "1ЛМ 27-11-14-4",
            "mark": "1ЛМ 27-11-14-4",
            "concrete_grade": "B25",
            "qty": 2,
            "unit_price": 14391.41,
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
    assert "1ЛМ 27-11-14-4" in flat
    assert "B25" in flat


def test_order_data_from_kp_marches() -> None:
    from core.kp_order_data import order_data_from_kp_info

    kp_info = {
        "product_type": "marches",
        "discount_percent": 0,
        "marches": [
            {
                "mark": "1ЛМ 27-11-14-4",
                "concrete_grade": "B25",
                "qty": 4,
                "unit_price": 100.0,
                "discounted_price": 100.0,
            }
        ],
    }
    order_data = order_data_from_kp_info(kp_info)
    assert len(order_data) == 1
    assert order_data[0]["product_kind"] == "march"
    assert order_data[0]["mark"] == "1ЛМ 27-11-14-4"
    assert order_data[0]["qty"] == 4


def test_march_draft_with_unknown_mark_has_null_unit_price(
    client: TestClient,
    auth_cookie: dict[str, str],
) -> None:
    response = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "marches", "text": "1ЛМ 99-99-99-9 B25 2"},
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["order_data"][0]["unit_price"] is None
    assert any("1ЛМ 99-99-99-9" in err for err in body["wizard_state"]["validation_errors"])


def test_march_draft_rejects_plates_endpoint(
    client: TestClient,
    auth_cookie: dict[str, str],
) -> None:
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "marches", "text": "1ЛМ 27-11-14-4 B25 2"},
    )
    draft_id = create.json()["draft_id"]

    response = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/plates",
        data={"mode": "replace", "text": "ПБ 78-12-8п 1"},
    )
    assert response.status_code == 400, response.text
