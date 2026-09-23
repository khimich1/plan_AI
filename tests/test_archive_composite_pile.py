"""CP-501: archive filter + save without GUID for composite piles."""

from __future__ import annotations

from pathlib import Path

import pytest
import sqlite3

from app.core.settings import get_settings
from app.main import create_app
from app.security.session import create_session_token
from core.composite_pile_kit import import_composite_pile_kit_from_xlsx
from core.composite_pile_price_db import import_composite_pile_prices_from_xlsx
from core.guid_gate import order_lines_from_order_data
from core.kp_db_schema import init_schema
from tests.helpers import kp_db_fixtures as fx
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient

KIT = Path(__file__).parent / "fixtures" / "composite_pile_kit_sample.xlsx"
PRICE = Path(__file__).parent / "fixtures" / "composite_pile_price_sample.xlsx"


def test_order_lines_map_composite_piles_kind() -> None:
    lines = order_lines_from_order_data(
        [
            {
                "product_type": "composite_piles",
                "mark": "С60.30-ВС.1",
                "qty": 2,
            }
        ]
    )
    assert len(lines) == 1
    assert lines[0].product_kind == "composite_pile"


def test_archive_filter_and_save_without_guid(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    monkeypatch.setenv("DRAFTS_DIR", str(tmp_path / "drafts"))
    monkeypatch.setenv("OUTPUTS_DIR", str(tmp_path / "outputs"))
    (tmp_path / "drafts").mkdir(exist_ok=True)
    (tmp_path / "outputs").mkdir(exist_ok=True)

    pb_db = tmp_path / "pb.db"
    import_composite_pile_kit_from_xlsx(str(KIT), str(pb_db))
    import_composite_pile_prices_from_xlsx(str(PRICE), str(pb_db))
    import core.commercial_offer as commercial_offer
    import core.commercial_offer_xlsx as commercial_offer_xlsx
    import app.services.product_draft_config as product_draft_config

    monkeypatch.setattr(commercial_offer, "DB_PATH", str(pb_db))
    monkeypatch.setattr(commercial_offer_xlsx, "DB_PATH", str(pb_db))
    monkeypatch.setattr(product_draft_config, "DB_PATH", str(pb_db))

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

    from app.repositories.manager_repository import ManagerRepository

    monkeypatch.setattr(
        ManagerRepository,
        "get_manager",
        lambda self, manager_id: {
            "id": 1,
            "fio": "Tester",
            "contact_number": "+79990001122",
            "email": "tester@test.local",
        }
        if manager_id == 1
        else None,
    )
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
        data={"product_type": "composite_piles", "text": "С140.30-С 5"},
    )
    draft_id = create.json().get("draft_id") or create.json()["draft"]["draft_id"]
    client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={
            "manager_id": 1,
            "client_name": "ООО Архив",
            "conditions_mode": "standard",
            "counterparty_id": fx.seed_test_counterparty(str(plita_db)),
        },
    )
    assert client.post(f"/api/v1/commercial/drafts/{draft_id}/calculate").status_code == 200
    save = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/save",
        json={"mode": "archive", "execution_terms_input": ""},
    )
    assert save.status_code == 200, save.text
    kp_id = save.json()["saved_offer"]["kp_id"]

    listed = client.get(
        "/api/v1/commercial/archive",
        params={"product_type": "composite_piles", "section": "archived"},
    )
    assert listed.status_code == 200, listed.text
    payload = listed.json()
    if isinstance(payload, list):
        items = payload
    else:
        items = payload.get("items") or payload.get("offers") or []
    assert any(int(i.get("kp_id") or 0) == int(kp_id) for i in items)

    details = client.get(f"/api/v1/commercial/archive/{kp_id}")
    assert details.status_code == 200, details.text
    body = details.json()
    offer = body.get("offer") if isinstance(body, dict) and "offer" in body else body
    assert offer.get("product_type") == "composite_piles"
