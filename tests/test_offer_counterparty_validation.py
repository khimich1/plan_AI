"""CTR-006: валидация counterparty_id на create/save и снапшот из справочника."""

from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any

import pytest
from pydantic import ValidationError

from app.repositories.counterparties_repository import CounterpartiesRepository
from app.repositories.kp_repository import KpRepository
from app.schemas.offers import CreateOfferRequest, OfferOrderItem
from app.services.commercial_workflow_service import CommercialWorkflowService
from app.services.counterparties_service import (
    CounterpartiesService,
    CounterpartyValidationError,
)
from app.services.offers_service import OffersService
from tests.helpers import kp_db_fixtures as fx
from tests.helpers.production_api_fixtures import VALID_APP_SECRET_KEY, session_cookie

_ORDER = [
    {
        "name": "ПБ 60-12-8п",
        "length_m": 6.0,
        "width_m": 1.2,
        "load_class": 800,
        "qty": 1,
        "unit_price": 1000.0,
        "weight": 500.0,
        "length_dm_raw": "60",
        "concrete_grade": "М500",
    }
]
_USER = {"id": 1, "role": "admin"}


def _seed_clients(db: str) -> dict[str, dict]:
    repo = CounterpartiesRepository(db_path=db)
    client = repo.insert(
        code_1c="00-client",
        name="РОМАШКА ООО",
        inn="7701000001",
        kpp="770101001",
        is_client=True,
        source="import",
    )
    supplier = repo.insert(
        code_1c="00-sup",
        name="ПОСТАВЩИК ООО",
        inn="9900000000",
        is_client=False,
        source="import",
    )
    inactive = repo.insert(
        code_1c="00-off",
        name="АРХИВ ООО",
        inn="8800000000",
        is_client=True,
        is_active=False,
        source="import",
    )
    return {"client": client, "supplier": supplier, "inactive": inactive}


def _payload(**overrides: Any) -> CreateOfferRequest:
    data = {
        "creation_date": "01.01.2026",
        "customer_name": "Имя из формы",
        "manager_name": "Иванов",
        "order_data": [OfferOrderItem.model_validate(_ORDER[0])],
        "save_mode": "archive",
    }
    data.update(overrides)
    return CreateOfferRequest.model_validate(data)


def test_create_offer_request_requires_counterparty_id() -> None:
    with pytest.raises(ValidationError):
        _payload()


def test_create_offer_snapshots_name_from_directory_not_form(tmp_path: Path) -> None:
    db = fx.make_iso_db(tmp_path)
    seeded = _seed_clients(db)
    service = OffersService(kp_repository=KpRepository(db_path=db))
    created = service.create_offer(
        _payload(counterparty_id=int(seeded["client"]["id"])),
        user=_USER,
    )
    kp_id = int(created["kp_id"])
    with sqlite3.connect(db) as conn:
        row = conn.execute(
            "SELECT customer_name, counterparty_id, customer_inn, customer_kpp "
            "FROM KP_offers WHERE kp_id = ?",
            (kp_id,),
        ).fetchone()
    assert row[0] == "РОМАШКА ООО"
    assert row[1] == seeded["client"]["id"]
    assert row[2] == "7701000001"
    assert row[3] == "770101001"


@pytest.mark.parametrize(
    ("seed_key", "needle"),
    [
        ("supplier", "не отмечен как клиент"),
        ("inactive", "не найден"),
    ],
)
def test_create_offer_rejects_non_client_and_inactive(
    tmp_path: Path, seed_key: str, needle: str
) -> None:
    db = fx.make_iso_db(tmp_path)
    seeded = _seed_clients(db)
    service = OffersService(kp_repository=KpRepository(db_path=db))
    with pytest.raises(CounterpartyValidationError, match=needle):
        service.create_offer(
            _payload(counterparty_id=int(seeded[seed_key]["id"])),
            user=_USER,
        )


def test_create_offer_rejects_missing_id(tmp_path: Path) -> None:
    db = fx.make_iso_db(tmp_path)
    _seed_clients(db)
    service = OffersService(kp_repository=KpRepository(db_path=db))
    with pytest.raises(CounterpartyValidationError, match="не найден"):
        service.create_offer(_payload(counterparty_id=999999), user=_USER)


def test_post_offers_without_counterparty_id_is_422(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from app.core.settings import get_settings
    from app.main import create_app
    from tests.helpers.auth_fixtures import patch_auth_users
    from tests.helpers.csrf import CsrfAwareTestClient

    db = fx.make_iso_db(tmp_path)
    monkeypatch.setenv("APP_SECRET_KEY", VALID_APP_SECRET_KEY)
    monkeypatch.setenv("PLITA_DB_PATH", db)
    get_settings.cache_clear()
    patch_auth_users(
        monkeypatch,
        [
            {
                "id": 1,
                "username": "admin",
                "role": "admin",
                "manager_id": None,
                "is_active": 1,
                "session_version": 0,
                "created_at": "2026-01-01 00:00:00",
            }
        ],
    )
    client = CsrfAwareTestClient(create_app())
    response = client.post(
        "/api/v1/offers",
        json={
            "creation_date": "01.01.2026",
            "customer_name": "Имя",
            "manager_name": "Иванов",
            "order_data": _ORDER,
        },
        cookies=session_cookie(1, "admin", "admin"),
    )
    assert response.status_code == 422


def _draft(metadata: dict[str, Any]) -> dict[str, Any]:
    return {
        "draft_id": "draft-cp",
        "order": {},
        "optimization": {},
        "order_data": [
            {
                "name": "ПБ 78-12-8п",
                "qty": 1,
                "length_m": 7.8,
                "width_m": 1.2,
                "unit_price": 1000.0,
                "concrete_grade": "М500",
            }
        ],
        "metadata": metadata,
        "saved_offer": metadata.get("saved_offer"),
    }


def _patch_save(workflow: CommercialWorkflowService, monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> dict:
    fake_xlsx = tmp_path / "kp.xlsx"
    fake_xlsx.write_bytes(b"xlsx")
    captured: dict[str, Any] = {"create": [], "update": []}
    monkeypatch.setattr(
        workflow,
        "generate_files",
        lambda *a, **k: [{"kind": "xlsx", "filename": fake_xlsx.name}],
    )
    monkeypatch.setattr(workflow.export_service, "resolve_generated_file", lambda filename: fake_xlsx)
    monkeypatch.setattr(workflow.draft_store, "update_metadata", lambda *a, **k: None)
    monkeypatch.setattr(
        workflow.calculation_service,
        "compute_totals_from_metadata",
        lambda *a, **k: {"total_with_vat": 0},
    )
    monkeypatch.setattr(
        workflow.export_service,
        "build_offer_identity_payload",
        lambda draft_id: {"offer_number": "1", "offer_date": "01.01.2026", "file_stem": "kp_1"},
    )

    def fake_create(**kwargs: Any) -> int:
        captured["create"].append(kwargs)
        return 10

    def fake_update(kp_id: int, order_data: list, **kwargs: Any) -> int:
        captured["update"].append({"kp_id": kp_id, **kwargs})
        return kp_id

    monkeypatch.setattr(workflow.kp_repository, "save_offer", fake_create)
    monkeypatch.setattr(workflow.kp_repository, "update_offer_from_order_data", fake_update)
    return captured


def test_draft_save_create_requires_valid_client(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    db = fx.make_iso_db(tmp_path)
    seeded = _seed_clients(db)
    workflow = CommercialWorkflowService()
    workflow.kp_repository = KpRepository(db_path=db)
    captured = _patch_save(workflow, monkeypatch, tmp_path)
    draft = _draft({"client_name": "Имя из формы", "manager_name": "Иванов"})
    monkeypatch.setattr(workflow, "_load_draft_or_raise", lambda _id: draft)

    with pytest.raises(CounterpartyValidationError, match="не найден"):
        workflow.save_offer("draft-cp", status="в архиве", save_mode="archive")

    draft["metadata"]["counterparty_id"] = int(seeded["client"]["id"])
    result = workflow.save_offer("draft-cp", status="в архиве", save_mode="archive")
    assert result["saved_offer"]["kp_id"] == 10
    assert captured["create"][0]["customer_name"] == "РОМАШКА ООО"
    assert captured["create"][0]["counterparty_id"] == seeded["client"]["id"]
    assert captured["create"][0]["customer_inn"] == "7701000001"


def test_draft_save_archive_update_skips_counterparty(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    db = fx.make_iso_db(tmp_path)
    workflow = CommercialWorkflowService()
    workflow.kp_repository = KpRepository(db_path=db)
    captured = _patch_save(workflow, monkeypatch, tmp_path)
    draft = _draft(
        {
            "client_name": "Старый клиент",
            "manager_name": "Иванов",
            "resume_kp_id": 55,
            "saved_offer": {"kp_id": 55, "status": "в архиве"},
        }
    )
    monkeypatch.setattr(workflow, "_load_draft_or_raise", lambda _id: draft)
    workflow.save_offer("draft-cp", status="в работе", save_mode="archive")
    assert captured["create"] == []
    assert captured["update"][0]["kp_id"] == 55
    assert "counterparty_id" not in captured["update"][0]


def test_update_draft_meta_validates_when_id_passed(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    db = fx.make_iso_db(tmp_path)
    seeded = _seed_clients(db)
    workflow = CommercialWorkflowService()
    workflow.kp_repository = KpRepository(db_path=db)
    draft = _draft({"client_name": "", "manager_name": ""})
    monkeypatch.setattr(workflow, "_load_draft_or_raise", lambda _id: draft)
    updates: dict[str, Any] = {}

    def fake_update(draft_id: str, **kwargs: Any) -> None:
        updates.update(kwargs)
        draft["metadata"].update(kwargs)

    monkeypatch.setattr(workflow.draft_store, "update_metadata", fake_update)
    monkeypatch.setattr(workflow, "_persist_wizard_step", lambda *a, **k: None)
    monkeypatch.setattr(workflow.draft_lifecycle, "get_draft_details", lambda _id: draft)

    with pytest.raises(CounterpartyValidationError, match="не найден"):
        workflow.update_draft_meta("draft-cp", counterparty_id=999999)

    workflow.update_draft_meta("draft-cp", counterparty_id=int(seeded["client"]["id"]))
    assert updates["counterparty_id"] == seeded["client"]["id"]
    assert updates["client_name"] == "РОМАШКА ООО"

    updates.clear()
    workflow.update_draft_meta("draft-cp", client_name="пока без привязки")
    assert "counterparty_id" not in updates
