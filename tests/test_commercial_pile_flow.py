"""PILE-103: API flow for pile commercial drafts."""

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


def _write_sample_pile_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 15, 20, 22.5, 25, "30 на граните"],
        [None, "35 СЕЧЕНИЕ", None, None, None, None, None],
        [69, "С120.35-12", 43760.31, 44108.15, 44371.09, 44634.03, 46159.37],
        [91, "С120.35-13и", 67512.27, 67860.11, 68123.05, 68385.98, 69911.33],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)


def _setup_pile_env(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> Path:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    monkeypatch.setenv("DRAFTS_DIR", str(tmp_path / "drafts"))
    monkeypatch.setenv("OUTPUTS_DIR", str(tmp_path / "outputs"))
    (tmp_path / "drafts").mkdir(exist_ok=True)
    (tmp_path / "outputs").mkdir(exist_ok=True)

    pile_xlsx = tmp_path / "piles.xlsx"
    pb_db = tmp_path / "pb.db"
    _write_sample_pile_xlsx(pile_xlsx)

    from core.pile_price_db import import_pile_prices_from_xlsx
    import_pile_prices_from_xlsx(str(pile_xlsx), str(pb_db))

    import core.commercial_offer as commercial_offer
    import core.commercial_offer_xlsx as commercial_offer_xlsx

    monkeypatch.setattr(commercial_offer, "DB_PATH", str(pb_db))
    monkeypatch.setattr(commercial_offer_xlsx, "DB_PATH", str(pb_db))
    get_settings.cache_clear()
    return pb_db


@pytest.fixture()
def client(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> TestClient:
    _setup_pile_env(monkeypatch, tmp_path)
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


def test_create_pile_draft_with_text(client: TestClient, auth_cookie: dict[str, str]) -> None:
    response = client.post(
        "/api/v1/commercial/drafts",
        data={
            "product_type": "piles",
            "text": "С120.35-12 B25 5\nС120.35-13и B30 2",
        },
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["metadata"]["product_type"] == "piles"
    assert len(body["order_data"]) == 2
    assert body["order_data"][0]["product_kind"] == "pile"
    assert body["order_data"][0]["unit_price"] is not None
    assert body["wizard_state"]["current_step"] == "piles"


def test_update_pile_draft_replace(client: TestClient, auth_cookie: dict[str, str]) -> None:
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "piles", "text": "С120.35-12 1"},
    )
    draft_id = create.json()["draft_id"]

    response = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/piles",
        data={"mode": "replace", "text": "С120.35-13и B25 3"},
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert len(body["order_data"]) == 1
    assert body["order_data"][0]["mark"] == "С120.35-13и"
    assert body["order_data"][0]["qty"] == 3


def test_create_plate_draft_default_product_type(
    client: TestClient,
    auth_cookie: dict[str, str],
) -> None:
    response = client.post(
        "/api/v1/commercial/drafts",
        data={"text": "ПБ 78-12-8п 1"},
    )
    assert response.status_code == 200, response.text
    assert response.json()["metadata"]["product_type"] == "plates"


def test_apply_ai_piles_endpoint(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from app.domain.models.optimization_context import OptimizationContext
    from app.domain.models.plate_order import PlateOrder
    from app.services.commercial_workflow_service import CommercialWorkflowService
    from app.services.draft_store import DraftStore

    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "piles", "text": "С120.35-12 B25 2"},
    )
    draft_id = create.json()["draft_id"]

    async def fake_apply(self, draft_id_arg: str, **kwargs):
        draft = self.get_draft_details(draft_id_arg)
        draft["metadata"]["source_type"] = "ai"
        draft["metadata"]["ai_applied"] = True
        draft["metadata"]["input_text"] = "С120.35-13и B25 3"
        return draft

    monkeypatch.setattr(CommercialWorkflowService, "apply_ai_piles_instruction", fake_apply)

    response = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/piles/ai",
        data={"instruction": "замени на 13и qty 3"},
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


def test_update_pile_grades_bulk(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "piles", "text": "С120.35-12 B25 2\nС120.35-13и B30 1"},
    )
    draft_id = create.json()["draft_id"]

    response = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/piles/grades",
        json={"concrete_grade": "B20"},
    )
    assert response.status_code == 200, response.text
    body = response.json()
    grades = {row["concrete_grade"] for row in body["order_data"]}
    assert grades == {"B20"}
    assert all(row["unit_price"] is not None for row in body["order_data"])


def test_pile_draft_calculate_and_generate_files(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _mock_manager_lookup(monkeypatch)
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "piles", "text": "С120.35-12 B25 2"},
    )
    draft_id = create.json()["draft_id"]

    meta = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={
            "manager_id": 1,
            "client_name": "ООО Сваи",
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


def test_pile_draft_save_to_archive(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    _setup_pile_env(monkeypatch, tmp_path)
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
        data={"product_type": "piles", "text": "С120.35-12 B25 2"},
    )
    draft_id = create.json()["draft_id"]

    client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={
            "manager_id": 1,
            "client_name": "ООО Сваи",
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

    plita_db = tmp_path / "plita.db"
    with sqlite3.connect(str(plita_db)) as conn:
        cur = conn.cursor()
        cur.execute("SELECT product_type FROM kp_meta WHERE kp_id = ?", (kp_id,))
        assert cur.fetchone()[0] == "piles"
        cur.execute("SELECT mark FROM kp_piles WHERE kp_id = ?", (kp_id,))
        assert cur.fetchone()[0] == "С120.35-12"


def test_pile_xlsx_contains_mark_and_grade() -> None:
    from core.commercial_offer_xlsx import generate_commercial_offer_xlsx

    order_data = [
        {
            "product_kind": "pile",
            "name": "С120.35-12",
            "mark": "С120.35-12",
            "concrete_grade": "B25",
            "qty": 2,
            "unit_price": 44634.03,
        }
    ]
    xlsx = generate_commercial_offer_xlsx(
        order_data,
        offer_number="3",
        offer_date="30.07.2026",
        customer_name="Test",
        kp_db_id=3,
    )
    df = pd.read_excel(xlsx, sheet_name="КП", header=None)
    flat = " ".join(str(v) for v in df.values.flatten() if pd.notna(v))
    assert "С120.35-12" in flat
    assert "B25" in flat


def test_order_data_from_kp_piles() -> None:
    from core.kp_order_data import order_data_from_kp_info

    kp_info = {
        "product_type": "piles",
        "discount_percent": 0,
        "piles": [
            {
                "mark": "С120.35-12",
                "concrete_grade": "B25",
                "qty": 4,
                "unit_price": 100.0,
                "discounted_price": 100.0,
            }
        ],
    }
    order_data = order_data_from_kp_info(kp_info)
    assert len(order_data) == 1
    assert order_data[0]["product_kind"] == "pile"
    assert order_data[0]["mark"] == "С120.35-12"
    assert order_data[0]["qty"] == 4


def test_pile_draft_with_unknown_mark_has_null_unit_price(
    client: TestClient,
    auth_cookie: dict[str, str],
) -> None:
    response = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "piles", "text": "С120.35-99 B25 2"},
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["order_data"][0]["mark"] == "С120.35-99"
    assert body["order_data"][0]["unit_price"] is None
    assert any("С120.35-99" in err for err in body["wizard_state"]["validation_errors"])


def test_pile_draft_calculate_rejects_unknown_mark(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _mock_manager_lookup(monkeypatch)
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "piles", "text": "С120.35-99 B25 2"},
    )
    draft_id = create.json()["draft_id"]

    meta = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={
            "manager_id": 1,
            "client_name": "ООО Сваи",
            "conditions_mode": "standard",
        },
    )
    assert meta.status_code == 200, meta.text

    calc = client.post(f"/api/v1/commercial/drafts/{draft_id}/calculate")
    assert calc.status_code == 422, calc.text
    body = calc.json()
    assert body["detail"]["code"] == "unpriced_plates"
    assert "С120.35-99" in body["detail"]["details"]["positions"][0]


def test_pile_draft_rejects_plates_endpoint(
    client: TestClient,
    auth_cookie: dict[str, str],
) -> None:
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "piles", "text": "С120.35-12 B25 2"},
    )
    draft_id = create.json()["draft_id"]

    response = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/plates",
        data={"mode": "replace", "text": "ПБ 78-12-8п 1"},
    )
    assert response.status_code == 400, response.text


def test_create_pile_draft_from_image_mock_ocr(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    from core.pile_format_prompt import build_pile_parser_system_prompt

    prompt = build_pile_parser_system_prompt()
    assert "Сваи С90.30-11" in prompt
    assert "189 шт" in prompt

    fixture_png = Path(__file__).parent / "fixtures" / "pile_ocr" / "pilot_table.png"
    assert fixture_png.is_file()

    async def fake_extract(
        self,
        *,
        image_bytes: bytes,
        image_filename: str | None,
        product_type: str = "plates",
    ):
        assert product_type == "piles"
        ocr_text = "\n".join(
            [
                "Сваи 90.30-11 189",
                "Свай 110.30-13 26",
                "Свай 120.30-12 20",
            ]
        )
        return ocr_text, {
            "ocr_method": "mock-gigachat",
            "ocr_api_calls": 1,
            "ocr_verify_skipped_reason": "auto_all_checks_passed",
            "ocr_plates": [],
        }

    from app.services.commercial_draft_service import CommercialDraftService

    monkeypatch.setattr(CommercialDraftService, "extract_text_from_image", fake_extract)

    with fixture_png.open("rb") as image_file:
        response = client.post(
            "/api/v1/commercial/drafts",
            data={"product_type": "piles"},
            files={"image": ("pilot_table.png", image_file, "image/png")},
        )

    assert response.status_code == 200, response.text
    body = response.json()
    assert body["metadata"]["product_type"] == "piles"
    assert body["metadata"]["source_type"] == "image"
    assert body["metadata"]["ocr_api_calls"] == 1
    assert body["metadata"]["ocr_verify_skipped_reason"] == "auto_all_checks_passed"
    assert len(body["order_data"]) == 3
    marks = [row["mark"] for row in body["order_data"]]
    assert marks == ["С90.30-11", "С110.30-13", "С120.30-12"]
    qtys = [row["qty"] for row in body["order_data"]]
    assert qtys == [189, 26, 20]

