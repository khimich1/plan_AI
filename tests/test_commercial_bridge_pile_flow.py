"""BP-601: API flow for bridge pile commercial drafts."""

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


def _write_sample_bridge_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 25, 30],
        [1, "C8-35T1", 35695.27, 0],
        [2, "C8-35T4; C8-35В4", 49813.83, 0],
        [3, "C13-40T3", 0, 89879.61],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)


def _setup_bridge_env(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> Path:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    monkeypatch.setenv("DRAFTS_DIR", str(tmp_path / "drafts"))
    monkeypatch.setenv("OUTPUTS_DIR", str(tmp_path / "outputs"))
    (tmp_path / "drafts").mkdir(exist_ok=True)
    (tmp_path / "outputs").mkdir(exist_ok=True)

    xlsx = tmp_path / "bridge.xlsx"
    pb_db = tmp_path / "pb.db"
    _write_sample_bridge_xlsx(xlsx)

    from core.bridge_pile_price_db import import_bridge_pile_prices_from_xlsx

    import_bridge_pile_prices_from_xlsx(str(xlsx), str(pb_db))

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
    _setup_bridge_env(monkeypatch, tmp_path)
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


def test_create_bridge_pile_draft_with_text(client: TestClient, auth_cookie: dict[str, str]) -> None:
    response = client.post(
        "/api/v1/commercial/drafts",
        data={
            "product_type": "bridge_piles",
            "text": "C8-35T1 B25 2\nC8-35В4 1",
        },
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["metadata"]["product_type"] == "bridge_piles"
    assert len(body["order_data"]) == 2
    assert body["order_data"][0]["product_kind"] == "bridge_pile"
    assert body["order_data"][0]["unit_price"] is not None
    marks = {row["mark"] for row in body["order_data"]}
    assert "C8-35В4" in marks
    assert body["wizard_state"]["current_step"] == "bridge_piles"
    for row in body["order_data"]:
        assert str(row.get("line_id") or "").strip()
        assert row.get("product_type") == "bridge_piles"


def test_bulk_grade_skips_unavailable_with_warning(
    client: TestClient, auth_cookie: dict[str, str]
) -> None:
    create = client.post(
        "/api/v1/commercial/drafts",
        data={
            "product_type": "bridge_piles",
            "text": "C8-35T1 B25 1\nC13-40T3 B30 1",
        },
    )
    draft_id = create.json()["draft_id"]

    response = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/bridge-piles/grades",
        json={"concrete_grade": "B25"},
    )
    assert response.status_code == 200, response.text
    body = response.json()
    by_mark = {row["mark"]: row for row in body["order_data"]}
    assert by_mark["C8-35T1"]["concrete_grade"] == "B25"
    assert by_mark["C13-40T3"]["concrete_grade"] == "B30"
    warnings = " ".join(body["metadata"].get("warnings") or [])
    assert "C13-40T3" in warnings


def test_bridge_pile_calculate_and_generate_files(
    client: TestClient,
    auth_cookie: dict[str, str],
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    _mock_manager_lookup(monkeypatch)
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "bridge_piles", "text": "C8-35В4 B25 2"},
    )
    draft_id = create.json()["draft_id"]
    assert create.json()["order_data"][0]["mark"] == "C8-35В4"

    meta = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={
            "manager_id": 1,
            "client_name": "ООО Мост",
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


def test_bridge_pile_persist_mark_as_typed(tmp_path: Path) -> None:
    db_path = str(tmp_path / "plita.db")
    init_schema(db_path)
    KpPersistenceService.save_kp_to_db(
        "05.08.2026",
        [
            {
                "product_kind": "bridge_pile",
                "name": "C8-35В4",
                "mark": "C8-35В4",
                "concrete_grade": "B25",
                "qty": 2,
                "unit_price": 49813.83,
            }
        ],
        customer_name="ООО Мост",
        status="в архиве",
        product_type="bridge_piles",
        db_path=db_path,
    )
    with sqlite3.connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute("SELECT product_type FROM kp_meta WHERE kp_id = 1")
        assert cur.fetchone()[0] == "bridge_piles"
        cur.execute("SELECT mark, concrete_grade, qty FROM kp_bridge_piles WHERE kp_id = 1")
        assert cur.fetchone() == ("C8-35В4", "B25", 2)
        cur.execute("SELECT COUNT(*) FROM kp_piles WHERE kp_id = 1")
        assert cur.fetchone()[0] == 0


TENDER_TEXT = (
    "C14-40T4 B25 52\n"
    "C9-35T6 B25 19\n"
    "C10-35T6 B25 19\n"
    "C13-35T7 B25 19\n"
    "C11-35T6 B25 19\n"
    "C15-35T6 B25 45\n"
    "C18-40T8 B25 49"
)
REAL_CATALOG_XLSX = Path(__file__).resolve().parents[1] / "банк знаний" / "сваи вес и объем.xlsx"


def _write_tender_bridge_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 25, 30],
        [1, "C14-40T4", 1000.0, 1100.0],
        [2, "C9-35T6", 1000.0, 1100.0],
        [3, "C10-35T6", 1000.0, 1100.0],
        [4, "C13-35T7", 1000.0, 1100.0],
        [5, "C11-35T6", 1000.0, 1100.0],
        [6, "C15-35T6", 1000.0, 1100.0],
        [7, "C18-40T8", 1000.0, 1100.0],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)


def _seed_pile_catalog(plita_db: str) -> None:
    from core.pile_catalog import PileCatalogEntry, parse_pile_catalog_from_xlsx, upsert_pile_catalog

    if REAL_CATALOG_XLSX.is_file():
        entries = parse_pile_catalog_from_xlsx(str(REAL_CATALOG_XLSX), sheet="Лист1")
    else:
        entries = [
            PileCatalogEntry("С140.40", 14.0, 400, 2.26, 5650.0, 3),
            PileCatalogEntry("С90.35", 9.0, 350, 1.12, 2800.0, 7),
            PileCatalogEntry("С100.35", 10.0, 350, 1.24, 3100.0, 6),
            PileCatalogEntry("С130.35", 13.0, 350, 1.61, 4030.0, 5),
            PileCatalogEntry("С110.35", 11.0, 350, 1.37, 3430.0, 6),
            PileCatalogEntry("С150.35", 15.0, 350, 1.86, 4650.0, 4),
        ]
    upsert_pile_catalog(plita_db, entries)


def test_bridge_tender_trips_pending_c18_then_override_save_archive_patch(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """PT-901: calculate 42 pending C18 → N → 42+N → save → archive PATCH."""
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    monkeypatch.setenv("DRAFTS_DIR", str(tmp_path / "drafts"))
    monkeypatch.setenv("OUTPUTS_DIR", str(tmp_path / "outputs"))
    (tmp_path / "drafts").mkdir(exist_ok=True)
    (tmp_path / "outputs").mkdir(exist_ok=True)

    xlsx = tmp_path / "bridge.xlsx"
    pb_db = tmp_path / "pb.db"
    _write_tender_bridge_xlsx(xlsx)
    from core.bridge_pile_price_db import import_bridge_pile_prices_from_xlsx

    import_bridge_pile_prices_from_xlsx(str(xlsx), str(pb_db))

    import core.commercial_offer as commercial_offer
    import core.commercial_offer_xlsx as commercial_offer_xlsx
    import app.services.commercial_workflow_service as commercial_workflow_service
    import app.services.commercial_calculation_service as commercial_calculation_service

    monkeypatch.setattr(commercial_offer, "DB_PATH", str(pb_db))
    monkeypatch.setattr(commercial_offer_xlsx, "DB_PATH", str(pb_db))
    monkeypatch.setattr(commercial_workflow_service, "DB_PATH", str(pb_db))
    monkeypatch.setattr(commercial_calculation_service, "DB_PATH", str(pb_db))

    plita_db = tmp_path / "plita.db"
    init_schema(str(plita_db))
    _seed_pile_catalog(str(plita_db))
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
        data={"product_type": "bridge_piles", "text": TENDER_TEXT},
    )
    assert create.status_code == 200, create.text
    draft_id = create.json()["draft_id"]
    marks = {row["mark"] for row in create.json()["order_data"]}
    assert "C18-40T8" in marks
    assert "C14-40T4" in marks

    meta = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={
            "manager_id": 1,
            "client_name": "ООО Мост",
            "discount_percent": 0,
            "conditions_mode": "standard",
            "pile_logistics_cost": 1000.0,
        },
    )
    assert meta.status_code == 200, meta.text

    calc = client.post(f"/api/v1/commercial/drafts/{draft_id}/calculate")
    assert calc.status_code == 200, calc.text
    totals = calc.json()["totals"]
    pending = totals.get("pile_trip_pending_marks") or []
    assert any("C18" in str(m).upper().replace("С", "C") for m in pending)
    assert totals.get("pile_delivery_ready") is False
    assert totals.get("pile_trips") == 0
    assert totals.get("pile_delivery_total") == 0

    meta_n = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={"pile_trip_overrides": {"C18-40T8": 3}},
    )
    assert meta_n.status_code == 200, meta_n.text

    calc2 = client.post(f"/api/v1/commercial/drafts/{draft_id}/calculate")
    assert calc2.status_code == 200, calc2.text
    totals2 = calc2.json()["totals"]
    assert totals2.get("pile_delivery_ready") is True
    assert totals2.get("pile_trips") == 45
    assert totals2.get("pile_delivery_total") == pytest.approx(45000.0)

    save = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/save",
        json={"mode": "archive", "execution_terms_input": ""},
    )
    assert save.status_code == 200, save.text
    kp_id = save.json()["saved_offer"]["kp_id"]
    assert kp_id == 1

    with sqlite3.connect(str(plita_db)) as conn:
        cur = conn.cursor()
        cur.execute("SELECT pile_logistics_cost FROM KP_offers WHERE kp_id = ?", (kp_id,))
        assert cur.fetchone()[0] == pytest.approx(1000.0)
        cur.execute("SELECT pile_trip_overrides_json FROM kp_meta WHERE kp_id = ?", (kp_id,))
        raw_json = cur.fetchone()[0]
    from core.pile_trip_pricing import coerce_pile_trip_overrides

    assert coerce_pile_trip_overrides(raw_json)["C18-40T8"] == 3

    details = client.get(f"/api/v1/commercial/archive/{kp_id}")
    assert details.status_code == 200, details.text
    assert details.json()["pile_trips"] == 45

    patched = client.patch(
        f"/api/v1/commercial/archive/{kp_id}/logistics-cost",
        json={
            "logistics_cost": 0,
            "pile_logistics_cost": 2000.0,
            "pile_trip_overrides": {"C18-40T8": 5},
        },
    )
    assert patched.status_code == 200, patched.text
    body = patched.json()
    assert body["pile_trips"] == 47
    assert body["pile_logistics_cost"] == pytest.approx(2000.0)
    assert body["pile_delivery_total"] == pytest.approx(94000.0)
