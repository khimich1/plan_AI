"""Импорт листа юриста: номер занимает seq, имя само не связывает контрагента."""

from __future__ import annotations

import io
import sqlite3
from pathlib import Path

import pytest
from openpyxl import Workbook

from app.dependencies.services import get_supply_contract_service
from app.main import create_app
from app.security.session import create_session_token
from app.services.supply_contract_service import (
    SupplyContractService,
    SupplyContractValidationError,
)
from core import kp_db_schema
from core.kp_db_common import _connect
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient
from tests.test_supply_contract import (
    ADMIN,
    _buyer,
    _fresh_db,
    _seed_counterparty,
    _seed_offer,
    _service,
)

MANAGER = {"id": 8, "role": "manager", "username": "manager"}


def _sheet_bytes() -> bytes:
    book = Workbook()
    sheet = book.active
    sheet.append(
        [
            "Номер договора",
            "Дата",
            "Контрагент",
            "Менеджер",
            "Оригинал/ЭДО",
            "Наличие скана",
        ]
    )
    sheet.append(
        [
            "1027/09/26",
            "25.09.2026",
            "ООО Ромашка",
            "Пургина",
            "подписан по ЭДО",
            "получен скан",
        ]
    )
    buffer = io.BytesIO()
    book.save(buffer)
    return buffer.getvalue()


def test_import_occupies_seq_and_does_not_link_by_name(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "import.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
        _seed_offer(conn, kp_id=1, counterparty_id=client_id)
    service = _service(db_path)

    imported = service.import_lawyer_sheet(_sheet_bytes(), user=ADMIN)

    assert len(imported) == 1
    row = imported[0]
    assert row["number"] == "1027/09/26"
    assert row["counterparty_id"] is None
    assert row["imported_name"] == "ООО Ромашка"
    assert any(item["id"] == client_id for item in row["suggestions"])
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        stored = conn.execute(
            "SELECT counterparty_id, number FROM supply_contract"
        ).fetchone()
    assert stored["counterparty_id"] is None
    assert stored["number"] == "1027/09/26"

    created = service.create_for_kp(1, _buyer(), user=ADMIN)
    assert created["number"] == "1028/09/26"


def test_link_sets_counterparty_only_when_called(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path, "link.db")
    with _connect(db_path) as conn:
        conn.row_factory = sqlite3.Row
        client_id = _seed_counterparty(conn)
    service = _service(db_path)
    imported = service.import_lawyer_sheet(_sheet_bytes(), user=ADMIN)
    contract_id = int(imported[0]["id"])

    with pytest.raises(SupplyContractValidationError, match="только администратору"):
        service.import_lawyer_sheet(_sheet_bytes(), user=MANAGER)

    linked = service.link_imported_contract(contract_id, client_id, user=MANAGER)
    assert linked["counterparty_id"] == client_id
    assert linked["number"] == "1027/09/26"


def test_import_endpoint_is_admin_only(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    db_path = _fresh_db(tmp_path, "http.db")
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    from app.core.settings import get_settings

    get_settings.cache_clear()
    kp_db_schema._schema_ready.clear()
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
                "session_version": 0,
            },
            {
                "id": 2,
                "username": "manager",
                "role": "manager",
                "manager_id": None,
                "is_active": 1,
                "created_at": "2026-01-01 00:00:00",
                "session_version": 0,
            },
        ],
    )
    app = create_app()
    app.dependency_overrides[get_supply_contract_service] = lambda: SupplyContractService(
        db_path=db_path
    )
    client = CsrfAwareTestClient(app)
    payload = {"file": ("registry.xlsx", _sheet_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}

    anonymous = client.post("/api/v1/commercial/archive/supply-contracts/import", files=payload)
    assert anonymous.status_code == 401

    manager = client.post(
        "/api/v1/commercial/archive/supply-contracts/import",
        files=payload,
        cookies=_cookie(2, "manager", "manager"),
    )
    assert manager.status_code == 403

    admin = client.post(
        "/api/v1/commercial/archive/supply-contracts/import",
        files=payload,
        cookies=_cookie(1, "admin", "admin"),
    )
    assert admin.status_code == 200
    body = admin.json()
    assert body[0]["number"] == "1027/09/26"
    assert body[0]["counterparty_id"] is None
    with _connect(db_path) as conn:
        count = conn.execute("SELECT COUNT(*) FROM supply_contract").fetchone()[0]
    assert count == 1


def _cookie(user_id: int, username: str, role: str) -> dict[str, str]:
    return {
        "app_session": create_session_token(
            {"id": user_id, "username": username, "role": role},
            ttl_seconds=300,
        )
    }
