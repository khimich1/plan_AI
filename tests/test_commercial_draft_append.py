"""MNA-101 / MNA-102 / MNA-103 / MNA-202 / MNA-304 / MNA-601: draft append + resume.

MNA-101: schema accepts line_id / product_type / append_batches (done).
MNA-102: plates/piles parse+calculate stamp non-empty ``line_id`` + ``product_type``.
MNA-103: append/start, cross-type type-update merge, undo-last, delete line.
MNA-202: mixed discount recomputes all lines; calculate not blocked by is_*_draft.
MNA-304: save_offer with saved_offer.kp_id → update same id; status gate.
MNA-601 (TDD RED): hydrate draft from saved KP (status «в работе» only).
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
import pytest
from fastapi.testclient import TestClient
from pydantic import ValidationError

from app.core.settings import get_settings
from app.main import create_app
from app.schemas import commercial as commercial_schemas
from app.schemas.commercial import (
    CommercialDraftDetailsResponse,
    CommercialDraftMetadata,
    CommercialOrderLine,
    CommercialSavedOffer,
)
from app.security.session import create_session_token
from app.services.commercial_workflow_service import CommercialWorkflowService
from tests.helpers.auth_fixtures import patch_auth_users
from tests.helpers.csrf import CsrfAwareTestClient


def _require_schema_export(name: str) -> type:
    cls = getattr(commercial_schemas, name, None)
    assert cls is not None, f"{name} missing from app.schemas.commercial (MNA-101)"
    return cls


def _legacy_order_line() -> dict[str, Any]:
    return {
        "name": "ПБ 78-12-8п",
        "qty": 2,
        "length_m": 7.8,
        "width_m": 1.2,
        "unit_price": 1000.0,
    }


def _mixed_order_lines() -> list[dict[str, Any]]:
    return [
        {
            "line_id": "ln_plates_1",
            "product_type": "plates",
            "append_batch_id": "batch-plates-1",
            "name": "ПБ 78-12-8п",
            "qty": 2,
            "unit_price": 1000.0,
        },
        {
            "line_id": "ln_piles_1",
            "product_type": "piles",
            "append_batch_id": "batch-piles-1",
            "mark": "С30.15-3",
            "concrete_grade": "B25",
            "qty": 12,
            "unit_price": 15200.0,
        },
    ]


def _minimal_draft_details(**overrides: Any) -> dict[str, Any]:
    base: dict[str, Any] = {
        "draft_id": "draft-append-101",
        "order": {},
        "optimization": {},
        "order_data": [_legacy_order_line()],
        "metadata": {
            "product_type": "plates",
            "client_name": "ООО Тест",
            "wide_plates_resolved": True,
            "current_step": "result",
        },
        "wizard_state": {
            "current_step": "result",
            "can_proceed_to": [],
            "next_required_action": "none",
            "validation_errors": [],
        },
        "files": [],
        "saved_offer": None,
        "totals": {"subtotal": 1000.0, "vat_amount": 200.0, "total_with_vat": 1200.0},
        "offer_identity": {
            "offer_number": "WEB_APPEND1",
            "offer_date": "12.08.2026",
            "file_stem": "kp_append1",
        },
    }
    base.update(overrides)
    return base


# --- CommercialAppendBatch ---


def test_commercial_append_batch_accepts_batch_id_product_type_and_line_ids() -> None:
    CommercialAppendBatch = _require_schema_export("CommercialAppendBatch")
    batch = CommercialAppendBatch.model_validate(
        {
            "batch_id": "batch-plates-1",
            "product_type": "plates",
            "line_ids": ["ln_plates_1", "ln_plates_2"],
        }
    )
    assert batch.batch_id == "batch-plates-1"
    assert batch.product_type == "plates"
    assert batch.line_ids == ["ln_plates_1", "ln_plates_2"]


def test_commercial_append_batch_defaults_line_ids_to_empty() -> None:
    CommercialAppendBatch = _require_schema_export("CommercialAppendBatch")
    batch = CommercialAppendBatch.model_validate(
        {"batch_id": "batch-empty", "product_type": "piles"}
    )
    assert batch.line_ids == []


def test_commercial_append_batch_rejects_invalid_product_type() -> None:
    CommercialAppendBatch = _require_schema_export("CommercialAppendBatch")
    with pytest.raises(ValidationError):
        CommercialAppendBatch.model_validate(
            {
                "batch_id": "batch-x",
                "product_type": "not-a-product",
                "line_ids": [],
            }
        )


# --- CommercialOrderLine ---


def test_commercial_order_line_accepts_line_id_product_type_append_batch_id() -> None:
    CommercialOrderLine = _require_schema_export("CommercialOrderLine")
    line = CommercialOrderLine.model_validate(
        {
            "line_id": "ln_01HZX",
            "product_type": "piles",
            "append_batch_id": "b3",
            "mark": "С30.15-3",
            "concrete_grade": "B25",
            "qty": 12,
            "unit_price": 15200.0,
        }
    )
    assert line.line_id == "ln_01HZX"
    assert line.product_type == "piles"
    assert line.append_batch_id == "b3"


def test_commercial_order_line_allows_legacy_payload_without_identity_fields() -> None:
    """Old mono lines without line_id / product_type / append_batch_id must still validate."""
    CommercialOrderLine = _require_schema_export("CommercialOrderLine")
    line = CommercialOrderLine.model_validate(_legacy_order_line())
    assert line.line_id is None
    assert line.append_batch_id is None
    assert line.product_type is None


def test_commercial_order_line_preserves_extra_product_fields() -> None:
    CommercialOrderLine = _require_schema_export("CommercialOrderLine")
    line = CommercialOrderLine.model_validate(
        {
            "line_id": "ln_fbs_1",
            "product_type": "fbs",
            "append_batch_id": "batch-fbs-1",
            "mark": "ФБС 24-4-6",
            "qty": 3,
            "unit_price": 500.0,
            "custom_flag": True,
        }
    )
    dumped = line.model_dump()
    assert dumped.get("mark") == "ФБС 24-4-6"
    assert dumped.get("custom_flag") is True


# --- CommercialDraftMetadata.append_batches / resume ---


def test_commercial_draft_metadata_accepts_append_batches() -> None:
    assert "append_batches" in CommercialDraftMetadata.model_fields, (
        "CommercialDraftMetadata.append_batches missing (MNA-101)"
    )
    meta = CommercialDraftMetadata.model_validate(
        {
            "product_type": "piles",
            "append_batches": [
                {
                    "batch_id": "batch-plates-1",
                    "product_type": "plates",
                    "line_ids": ["ln_plates_1"],
                },
                {
                    "batch_id": "batch-piles-1",
                    "product_type": "piles",
                    "line_ids": ["ln_piles_1"],
                },
            ],
        }
    )
    assert len(meta.append_batches) == 2
    assert meta.append_batches[0].batch_id == "batch-plates-1"
    assert meta.append_batches[0].product_type == "plates"
    assert meta.append_batches[0].line_ids == ["ln_plates_1"]
    assert meta.append_batches[1].batch_id == "batch-piles-1"
    # metadata.product_type = current cycle type (not "type of whole KP")
    assert meta.product_type == "piles"


def test_commercial_draft_metadata_defaults_append_batches_to_empty() -> None:
    assert "append_batches" in CommercialDraftMetadata.model_fields, (
        "CommercialDraftMetadata.append_batches missing (MNA-101)"
    )
    meta = CommercialDraftMetadata.model_validate({"product_type": "plates"})
    assert meta.append_batches == []


def test_commercial_draft_metadata_accepts_optional_resume_kp_id() -> None:
    """Archive C resume: optional resume_kp_id on metadata (alongside saved_offer.kp_id)."""
    assert "resume_kp_id" in CommercialDraftMetadata.model_fields, (
        "CommercialDraftMetadata.resume_kp_id missing (MNA-101)"
    )
    meta = CommercialDraftMetadata.model_validate(
        {
            "product_type": "plates",
            "resume_kp_id": 42,
            "append_batches": [],
        }
    )
    assert meta.resume_kp_id == 42


def test_commercial_draft_metadata_resume_kp_id_defaults_to_none() -> None:
    assert "resume_kp_id" in CommercialDraftMetadata.model_fields, (
        "CommercialDraftMetadata.resume_kp_id missing (MNA-101)"
    )
    meta = CommercialDraftMetadata.model_validate({"product_type": "steps"})
    assert meta.resume_kp_id is None


# --- Full draft details round-trip ---


def test_commercial_draft_details_accepts_mixed_append_payload() -> None:
    _require_schema_export("CommercialAppendBatch")
    _require_schema_export("CommercialOrderLine")
    assert "append_batches" in CommercialDraftMetadata.model_fields
    assert "resume_kp_id" in CommercialDraftMetadata.model_fields

    payload = _minimal_draft_details(
        order_data=_mixed_order_lines(),
        metadata={
            "product_type": "piles",
            "client_name": "ООО Микс",
            "discount_percent": 5.0,
            "wide_plates_resolved": True,
            "current_step": "result",
            "resume_kp_id": 7,
            "append_batches": [
                {
                    "batch_id": "batch-plates-1",
                    "product_type": "plates",
                    "line_ids": ["ln_plates_1"],
                },
                {
                    "batch_id": "batch-piles-1",
                    "product_type": "piles",
                    "line_ids": ["ln_piles_1"],
                },
            ],
        },
        saved_offer={
            "kp_id": 7,
            "status": "в работе",
            "mode": "database",
            "execution_terms": "",
            "saved_at": "2026-08-12T10:00:00",
        },
    )
    details = CommercialDraftDetailsResponse.model_validate(payload)

    assert details.draft_id == "draft-append-101"
    assert details.metadata.product_type == "piles"
    assert details.metadata.resume_kp_id == 7
    assert len(details.metadata.append_batches) == 2
    assert details.metadata.append_batches[0].line_ids == ["ln_plates_1"]

    assert isinstance(details.order_data[0], CommercialOrderLine)
    assert isinstance(details.order_data[1], CommercialOrderLine)
    assert details.order_data[0].line_id == "ln_plates_1"
    assert details.order_data[1].product_type == "piles"
    assert details.order_data[0].append_batch_id == "batch-plates-1"

    # Existing resume pattern: saved_offer.kp_id (Q1=C)
    assert details.saved_offer is not None
    assert details.saved_offer.kp_id == 7
    assert details.saved_offer.status == "в работе"


def test_commercial_draft_details_accepts_legacy_payload_without_append_fields() -> None:
    """Legacy mono draft without line_id / append_batches / resume_kp_id must still validate."""
    assert "append_batches" in CommercialDraftMetadata.model_fields, (
        "CommercialDraftMetadata.append_batches missing (MNA-101)"
    )
    assert "resume_kp_id" in CommercialDraftMetadata.model_fields, (
        "CommercialDraftMetadata.resume_kp_id missing (MNA-101)"
    )
    details = CommercialDraftDetailsResponse.model_validate(_minimal_draft_details())
    assert details.draft_id == "draft-append-101"
    assert details.metadata.product_type == "plates"
    assert details.metadata.append_batches == []
    assert details.metadata.resume_kp_id is None
    assert details.saved_offer is None
    assert len(details.order_data) == 1
    assert isinstance(details.order_data[0], CommercialOrderLine)
    assert details.order_data[0].line_id is None
    assert details.order_data[0].product_type is None
    assert details.order_data[0].append_batch_id is None


def test_commercial_draft_details_rejects_invalid_product_type_on_order_line() -> None:
    payload = _minimal_draft_details(
        order_data=[
            {
                "line_id": "ln_bad",
                "product_type": "not-a-product",
                "name": "X",
                "qty": 1,
            }
        ]
    )
    with pytest.raises(ValidationError):
        CommercialDraftDetailsResponse.model_validate(payload)


def test_commercial_saved_offer_still_carries_kp_id_for_resume() -> None:
    """Follow existing schema: resume of saved KP uses saved_offer.kp_id."""
    saved = CommercialSavedOffer.model_validate(
        {"kp_id": 99, "status": "в работе", "mode": "database"}
    )
    assert saved.kp_id == 99


# --- MNA-102: stamp line_id + product_type on calculate / input paths ---


def _assert_order_lines_have_identity(
    order_data: list[dict[str, Any]],
    *,
    product_type: str,
) -> list[str]:
    assert order_data, "expected non-empty order_data"
    line_ids: list[str] = []
    for idx, line in enumerate(order_data):
        line_id = str(line.get("line_id") or "").strip()
        assert line_id, f"order_data[{idx}] missing non-empty line_id: {line!r}"
        assert line.get("product_type") == product_type, (
            f"order_data[{idx}] product_type={line.get('product_type')!r}, "
            f"expected {product_type!r}"
        )
        line_ids.append(line_id)
    assert len(line_ids) == len(set(line_ids)), f"line_id must be unique: {line_ids}"
    return line_ids


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


def _write_sample_plate_xlsx(path: Path) -> None:
    """Minimal plate price book covering MNA-102 sample lines (ПБ 60/78-12-8п)."""
    df = pd.DataFrame(
        {
            "Unnamed: 0": ["ПБ 60-12", "ПБ 78-12"],
            "6 нагрузка": [5000, 6000],
            "8 нагрузка": [5000, 6000],
            "10 нагрузка": [5000, 6000],
            "12 нагрузка": [5000, 6000],
        }
    )
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="14.07.2026", index=False, startrow=1)


def _write_sample_fbs_xlsx(path: Path) -> None:
    rows = [
        [None, "Наименование", 7.5, 20, 22.5, 25],
        [1, "ФБС 9.3.6-Т", 1640.75, 1731.47, 1759.90, 1788.33],
        [2, "ФБС 12.4.6-Т", 2683.65, 2848.31, 2899.91, 2951.52],
    ]
    df = pd.DataFrame(rows)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Прайс", index=False, header=False)


def _seed_prays_plity(pb_db: Path) -> None:
    """Ensure nomenclature lookup has a table (empty is fine; avoids OperationalError)."""
    import sqlite3

    with sqlite3.connect(str(pb_db)) as conn:
        conn.execute(
            'CREATE TABLE IF NOT EXISTS prays_plity ('
            '"Уникальный идентификатор (Номенклатура)" TEXT, '
            '"Товар" TEXT)'
        )
        conn.commit()


def _setup_draft_env(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    monkeypatch.setenv("APP_SECRET_KEY", "test-secret-key-for-pytest-must-be-32-chars-min")
    monkeypatch.setenv("DRAFTS_DIR", str(tmp_path / "drafts"))
    monkeypatch.setenv("OUTPUTS_DIR", str(tmp_path / "outputs"))
    (tmp_path / "drafts").mkdir(exist_ok=True)
    (tmp_path / "outputs").mkdir(exist_ok=True)
    get_settings.cache_clear()


def _patch_commercial_db_path(monkeypatch: pytest.MonkeyPatch, pb_db: Path) -> None:
    import app.services.commercial_calculation_service as calc_mod
    import app.services.commercial_export_service as export_mod
    import app.services.commercial_workflow_service as workflow_mod
    import core.commercial_offer as commercial_offer
    import core.commercial_offer_xlsx as commercial_offer_xlsx
    import core.kp_db_nomenclature as kp_db_nomenclature

    monkeypatch.setattr(commercial_offer, "DB_PATH", str(pb_db))
    monkeypatch.setattr(commercial_offer_xlsx, "DB_PATH", str(pb_db))
    monkeypatch.setattr(workflow_mod, "DB_PATH", str(pb_db))
    monkeypatch.setattr(calc_mod, "DB_PATH", str(pb_db))
    monkeypatch.setattr(export_mod, "DB_PATH", str(pb_db))
    monkeypatch.setattr(kp_db_nomenclature, "_PB_DB_PATH", str(pb_db))


def _patch_price_db_path(monkeypatch: pytest.MonkeyPatch, pb_db: Path) -> None:
    """Redirect PRICE_DB_PATH / PB_DB_PATH bindings captured at import (ILP, procurement, etc.)."""
    import core.config_and_data as config_and_data
    import core.db_config as db_config
    import core.optimization.ilp_model as ilp_model
    import core.price_db as price_db
    import core.project_paths as project_paths
    import core.visualization as visualization
    from viz_modules.layout_sequence import deps as layout_deps
    from viz_modules.procurement import adapters_default

    pb_path = Path(pb_db)
    pb_str = str(pb_path)

    monkeypatch.setenv("PB_DB_PATH", pb_str)
    monkeypatch.setenv("PRICE_DB_PATH", pb_str)

    # Path-typed module bindings (from core.project_paths import PRICE_DB_PATH)
    monkeypatch.setattr(project_paths, "PRICE_DB_PATH", pb_path)
    monkeypatch.setattr(config_and_data, "PRICE_DB_PATH", pb_path)
    monkeypatch.setattr(ilp_model, "PRICE_DB_PATH", pb_path)
    monkeypatch.setattr(adapters_default, "PRICE_DB_PATH", pb_path)
    monkeypatch.setattr(layout_deps, "PRICE_DB_PATH", pb_path)
    monkeypatch.setattr(visualization, "PRICE_DB_PATH", pb_path)

    # str defaults / settings snapshots
    monkeypatch.setattr(price_db, "DEFAULT_DB", pb_str)
    monkeypatch.setattr(db_config, "PB_DB_PATH", pb_str)


def _setup_pile_price_env(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    _setup_draft_env(monkeypatch, tmp_path)
    pile_xlsx = tmp_path / "piles.xlsx"
    pb_db = tmp_path / "pb.db"
    _write_sample_pile_xlsx(pile_xlsx)

    from core.pile_price_db import import_pile_prices_from_xlsx

    import_pile_prices_from_xlsx(str(pile_xlsx), str(pb_db))
    _patch_commercial_db_path(monkeypatch, pb_db)
    get_settings.cache_clear()


def _setup_plate_price_env(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> Path:
    """Isolate plates pb.db (prices + prays_plity) under tmp_path — writable for SQLite journals."""
    _setup_draft_env(monkeypatch, tmp_path)
    plate_xlsx = tmp_path / "plates.xlsx"
    pb_db = tmp_path / "pb.db"
    plita_db = tmp_path / "plita.db"
    _write_sample_plate_xlsx(plate_xlsx)

    from core.price_db import import_from_xlsx

    import_from_xlsx(str(plate_xlsx), str(pb_db), preferred_sheet="14.07.2026")
    _seed_prays_plity(pb_db)

    # PlateParserService resolves pb as sibling of plita.db
    monkeypatch.setenv("PLITA_DB_PATH", str(plita_db))
    _patch_commercial_db_path(monkeypatch, pb_db)
    _patch_price_db_path(monkeypatch, pb_db)
    get_settings.cache_clear()
    return pb_db


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


def _auth_client(monkeypatch: pytest.MonkeyPatch) -> TestClient:
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
    return client


def _patch_client_meta(
    client: TestClient,
    draft_id: str,
    *,
    client_name: str,
    discount_percent: float = 0.0,
) -> None:
    meta = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={
            "manager_id": 1,
            "client_name": client_name,
            "discount_percent": discount_percent,
            "conditions_mode": "standard",
        },
    )
    assert meta.status_code == 200, meta.text


def _setup_mixed_price_env(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> Path:
    """Plates + piles + FBS prices in one pb.db for cross-type append API tests (MNA-103)."""
    _setup_draft_env(monkeypatch, tmp_path)
    plate_xlsx = tmp_path / "plates.xlsx"
    pile_xlsx = tmp_path / "piles.xlsx"
    fbs_xlsx = tmp_path / "fbs.xlsx"
    pb_db = tmp_path / "pb.db"
    plita_db = tmp_path / "plita.db"
    _write_sample_plate_xlsx(plate_xlsx)
    _write_sample_pile_xlsx(pile_xlsx)
    _write_sample_fbs_xlsx(fbs_xlsx)

    from core.fbs_price_db import import_fbs_prices_from_xlsx
    from core.pile_price_db import import_pile_prices_from_xlsx
    from core.price_db import import_from_xlsx

    import_from_xlsx(str(plate_xlsx), str(pb_db), preferred_sheet="14.07.2026")
    import_pile_prices_from_xlsx(str(pile_xlsx), str(pb_db))
    import_fbs_prices_from_xlsx(str(fbs_xlsx), str(pb_db))
    _seed_prays_plity(pb_db)

    monkeypatch.setenv("PLITA_DB_PATH", str(plita_db))
    _patch_commercial_db_path(monkeypatch, pb_db)
    _patch_price_db_path(monkeypatch, pb_db)
    get_settings.cache_clear()
    return pb_db


@pytest.fixture()
def plates_client(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> TestClient:
    _setup_plate_price_env(monkeypatch, tmp_path)
    return _auth_client(monkeypatch)


@pytest.fixture()
def piles_client(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> TestClient:
    _setup_pile_price_env(monkeypatch, tmp_path)
    return _auth_client(monkeypatch)


@pytest.fixture()
def mixed_client(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> TestClient:
    """Auth client with both plate and pile prices (MNA-103 cross-type append)."""
    _setup_mixed_price_env(monkeypatch, tmp_path)
    return _auth_client(monkeypatch)


def test_plates_parse_stamps_nonempty_line_id_and_product_type(
    plates_client: TestClient,
) -> None:
    """MNA-102 input path: after plates parse/create every line has line_id + product_type."""
    response = plates_client.post(
        "/api/v1/commercial/drafts",
        data={"text": "ПБ 78-12-8п 2\nПБ 60-12-8п 1"},
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["metadata"]["product_type"] == "plates"
    _assert_order_lines_have_identity(body["order_data"], product_type="plates")


def test_plates_calculate_stamps_nonempty_line_id_and_product_type(
    plates_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-102 calculate path: after plates calculate every line has line_id + product_type."""
    _mock_manager_lookup(monkeypatch)
    create = plates_client.post(
        "/api/v1/commercial/drafts",
        data={"text": "ПБ 78-12-8п 2\nПБ 60-12-8п 1"},
    )
    assert create.status_code == 200, create.text
    draft_id = create.json()["draft_id"]
    _patch_client_meta(plates_client, draft_id, client_name="ООО Плиты MNA-102")

    calc = plates_client.post(f"/api/v1/commercial/drafts/{draft_id}/calculate")
    assert calc.status_code == 200, calc.text
    body = calc.json()
    assert body["wizard_state"]["current_step"] == "result"
    _assert_order_lines_have_identity(body["order_data"], product_type="plates")


def test_plates_line_id_stable_across_recalculate(
    plates_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-102: unchanged plate lines keep the same line_id across recalculate."""
    _mock_manager_lookup(monkeypatch)
    create = plates_client.post(
        "/api/v1/commercial/drafts",
        data={"text": "ПБ 78-12-8п 2\nПБ 60-12-8п 1"},
    )
    assert create.status_code == 200, create.text
    draft_id = create.json()["draft_id"]
    _patch_client_meta(plates_client, draft_id, client_name="ООО Плиты Stable")

    first = plates_client.post(f"/api/v1/commercial/drafts/{draft_id}/calculate")
    assert first.status_code == 200, first.text
    ids_first = _assert_order_lines_have_identity(
        first.json()["order_data"], product_type="plates"
    )

    second = plates_client.post(f"/api/v1/commercial/drafts/{draft_id}/calculate")
    assert second.status_code == 200, second.text
    ids_second = _assert_order_lines_have_identity(
        second.json()["order_data"], product_type="plates"
    )
    assert ids_second == ids_first


def test_plates_line_id_stable_across_identical_replace(
    plates_client: TestClient,
) -> None:
    """MNA-102 prefer-stable: replace with same plate text keeps line_ids for unchanged lines."""
    create = plates_client.post(
        "/api/v1/commercial/drafts",
        data={"text": "ПБ 78-12-8п 2\nПБ 60-12-8п 1"},
    )
    assert create.status_code == 200, create.text
    draft_id = create.json()["draft_id"]
    ids_before = _assert_order_lines_have_identity(
        create.json()["order_data"], product_type="plates"
    )

    replace = plates_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/plates",
        data={"mode": "replace", "text": "ПБ 78-12-8п 2\nПБ 60-12-8п 1"},
    )
    assert replace.status_code == 200, replace.text
    ids_after = _assert_order_lines_have_identity(
        replace.json()["order_data"], product_type="plates"
    )
    assert ids_after == ids_before


def test_piles_parse_stamps_nonempty_line_id_and_product_type(
    piles_client: TestClient,
) -> None:
    """MNA-102 input path: after piles parse/create every line has line_id + product_type."""
    response = piles_client.post(
        "/api/v1/commercial/drafts",
        data={
            "product_type": "piles",
            "text": "С120.35-12 B25 5\nС120.35-13и B30 2",
        },
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["metadata"]["product_type"] == "piles"
    _assert_order_lines_have_identity(body["order_data"], product_type="piles")


def test_piles_calculate_stamps_nonempty_line_id_and_product_type(
    piles_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-102 calculate path: after piles calculate every line has line_id + product_type."""
    _mock_manager_lookup(monkeypatch)
    create = piles_client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "piles", "text": "С120.35-12 B25 2\nС120.35-13и B25 1"},
    )
    assert create.status_code == 200, create.text
    draft_id = create.json()["draft_id"]
    _patch_client_meta(piles_client, draft_id, client_name="ООО Сваи MNA-102")

    calc = piles_client.post(f"/api/v1/commercial/drafts/{draft_id}/calculate")
    assert calc.status_code == 200, calc.text
    body = calc.json()
    assert body["wizard_state"]["current_step"] == "result"
    _assert_order_lines_have_identity(body["order_data"], product_type="piles")


def test_piles_line_id_stable_across_recalculate(
    piles_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-102: unchanged pile lines keep the same line_id across recalculate."""
    _mock_manager_lookup(monkeypatch)
    create = piles_client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "piles", "text": "С120.35-12 B25 2\nС120.35-13и B25 1"},
    )
    assert create.status_code == 200, create.text
    draft_id = create.json()["draft_id"]
    _patch_client_meta(piles_client, draft_id, client_name="ООО Сваи Stable")

    first = piles_client.post(f"/api/v1/commercial/drafts/{draft_id}/calculate")
    assert first.status_code == 200, first.text
    ids_first = _assert_order_lines_have_identity(
        first.json()["order_data"], product_type="piles"
    )

    second = piles_client.post(f"/api/v1/commercial/drafts/{draft_id}/calculate")
    assert second.status_code == 200, second.text
    ids_second = _assert_order_lines_have_identity(
        second.json()["order_data"], product_type="piles"
    )
    assert ids_second == ids_first


def test_piles_line_id_stable_across_identical_replace(
    piles_client: TestClient,
) -> None:
    """MNA-102 prefer-stable: replace with same pile text keeps line_ids for unchanged lines."""
    create = piles_client.post(
        "/api/v1/commercial/drafts",
        data={
            "product_type": "piles",
            "text": "С120.35-12 B25 5\nС120.35-13и B30 2",
        },
    )
    assert create.status_code == 200, create.text
    draft_id = create.json()["draft_id"]
    ids_before = _assert_order_lines_have_identity(
        create.json()["order_data"], product_type="piles"
    )

    replace = piles_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/piles",
        data={"mode": "replace", "text": "С120.35-12 B25 5\nС120.35-13и B30 2"},
    )
    assert replace.status_code == 200, replace.text
    ids_after = _assert_order_lines_have_identity(
        replace.json()["order_data"], product_type="piles"
    )
    assert ids_after == ids_before


# --- MNA-103: append/start, cross-type merge, undo-last, delete line ---
#
# Assumed contracts (minimal surface; commit via existing type update + calculate):
#   POST   /api/v1/commercial/drafts/{id}/append/start
#          JSON {"product_type": "<ProductType>"}
#          → set metadata.product_type; clear cycle input (input_text); keep header;
#            keep prior order_data; seal prior cycle into append_batches if needed
#   PATCH  /api/v1/commercial/drafts/{id}/piles|plates|…  mode=append
#          → merge new cycle lines onto prior order_data (do not wipe other types)
#   POST   /api/v1/commercial/drafts/{id}/calculate
#          → price all lines; record current cycle in metadata.append_batches
#   POST   /api/v1/commercial/drafts/{id}/append/undo-last
#          → drop last append_batches entry and its lines from order_data
#   DELETE /api/v1/commercial/drafts/{id}/lines/{line_id}
#          → remove that line only; update append_batches.line_ids

_PLATES_APPEND_TEXT = "ПБ 78-12-8п 2\nПБ 60-12-8п 1"
_PLATES_APPEND_TEXT_2 = "ПБ 78-12-8п 1"
_PILES_APPEND_TEXT = "С120.35-12 B25 2\nС120.35-13и B25 1"
_FBS_APPEND_TEXT = "ФБС 9.3.6-Т B25 2"
_APPEND_DISCOUNT = 6.0
_APPEND_CLIENT = "ООО Микс MNA-103"


def _assert_append_batches_cover_order(
    order_data: list[dict[str, Any]],
    append_batches: list[dict[str, Any]],
) -> None:
    assert append_batches, "expected metadata.append_batches after append cycles"
    covered: list[str] = []
    for batch in append_batches:
        assert str(batch.get("batch_id") or "").strip(), f"batch missing batch_id: {batch!r}"
        assert batch.get("product_type") in {
            "plates",
            "piles",
            "steps",
            "marches",
            "bridge_piles",
            "fbs",
        }, batch
        line_ids = list(batch.get("line_ids") or [])
        assert line_ids, f"batch {batch.get('batch_id')!r} has empty line_ids"
        covered.extend(line_ids)

    order_ids = [str(line.get("line_id") or "") for line in order_data]
    assert covered == order_ids, (
        f"append_batches line_ids must match chronological order_data ids;\n"
        f"covered={covered!r}\norder_ids={order_ids!r}"
    )
    for line in order_data:
        assert str(line.get("append_batch_id") or "").strip(), (
            f"line missing append_batch_id: {line!r}"
        )


def _create_plates_result_draft(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
    *,
    client_name: str = _APPEND_CLIENT,
    discount_percent: float = _APPEND_DISCOUNT,
    plates_text: str = _PLATES_APPEND_TEXT,
) -> tuple[str, list[dict[str, Any]]]:
    """Create plates draft → meta → calculate; return (draft_id, plate order_data)."""
    _mock_manager_lookup(monkeypatch)
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"text": plates_text},
    )
    assert create.status_code == 200, create.text
    draft_id = create.json()["draft_id"]
    _patch_client_meta(
        client,
        draft_id,
        client_name=client_name,
        discount_percent=discount_percent,
    )
    calc = client.post(f"/api/v1/commercial/drafts/{draft_id}/calculate")
    assert calc.status_code == 200, calc.text
    body = calc.json()
    assert body["wizard_state"]["current_step"] == "result"
    plates = list(body["order_data"] or [])
    _assert_order_lines_have_identity(plates, product_type="plates")
    return draft_id, plates


def _append_piles_cycle(
    client: TestClient,
    draft_id: str,
    *,
    piles_text: str = _PILES_APPEND_TEXT,
) -> dict[str, Any]:
    """append/start(piles) → PATCH piles mode=append → calculate; return final body."""
    start = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/append/start",
        json={"product_type": "piles"},
    )
    assert start.status_code == 200, start.text
    start_body = start.json()
    assert start_body["metadata"]["product_type"] == "piles"

    update = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/piles",
        data={"mode": "append", "text": piles_text},
    )
    assert update.status_code == 200, update.text

    calc = client.post(f"/api/v1/commercial/drafts/{draft_id}/calculate")
    assert calc.status_code == 200, calc.text
    body = calc.json()
    assert body["wizard_state"]["current_step"] == "result"
    return body


def _append_product_cycle(
    client: TestClient,
    draft_id: str,
    *,
    product_type: str,
    text: str,
    patch_path: str | None = None,
) -> dict[str, Any]:
    """append/start(type) → PATCH type mode=append → calculate; return final body."""
    endpoint = patch_path or product_type
    start = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/append/start",
        json={"product_type": product_type},
    )
    assert start.status_code == 200, start.text
    assert start.json()["metadata"]["product_type"] == product_type

    update = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/{endpoint}",
        data={"mode": "append", "text": text},
    )
    assert update.status_code == 200, update.text

    calc = client.post(f"/api/v1/commercial/drafts/{draft_id}/calculate")
    assert calc.status_code == 200, calc.text
    body = calc.json()
    assert body["wizard_state"]["current_step"] == "result"
    return body


def _build_plates_then_piles_draft(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> tuple[str, dict[str, Any], list[dict[str, Any]], list[dict[str, Any]]]:
    """Full plates→piles append flow. Returns draft_id, body, plate_lines, pile_lines."""
    draft_id, plates_before = _create_plates_result_draft(client, monkeypatch)
    body = _append_piles_cycle(client, draft_id)
    order_data = list(body["order_data"] or [])
    plate_lines = [ln for ln in order_data if ln.get("product_type") == "plates"]
    pile_lines = [ln for ln in order_data if ln.get("product_type") == "piles"]
    assert len(plate_lines) == len(plates_before)
    return draft_id, body, plate_lines, pile_lines


def test_append_start_sets_product_type_clears_cycle_input_keeps_header(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-103: POST append/start switches cycle type, clears input, keeps sticky header."""
    draft_id, plates = _create_plates_result_draft(mixed_client, monkeypatch)
    plate_ids = [str(ln["line_id"]) for ln in plates]

    before = mixed_client.get(f"/api/v1/commercial/drafts/{draft_id}")
    assert before.status_code == 200, before.text
    meta_before = before.json()["metadata"]
    assert meta_before["client_name"] == _APPEND_CLIENT
    assert float(meta_before["discount_percent"]) == _APPEND_DISCOUNT
    assert meta_before.get("manager_id") == 1
    assert str(meta_before.get("input_text") or "").strip(), (
        "precondition: plates cycle should have non-empty input_text before append/start"
    )

    start = mixed_client.post(
        f"/api/v1/commercial/drafts/{draft_id}/append/start",
        json={"product_type": "piles"},
    )
    assert start.status_code == 200, start.text
    body = start.json()
    meta = body["metadata"]

    assert meta["product_type"] == "piles"
    assert meta["client_name"] == _APPEND_CLIENT
    assert float(meta["discount_percent"]) == _APPEND_DISCOUNT
    assert meta.get("manager_id") == 1
    assert str(meta.get("input_text") or "").strip() == ""
    assert [str(ln.get("line_id")) for ln in body["order_data"]] == plate_ids
    _assert_order_lines_have_identity(body["order_data"], product_type="plates")


def test_plates_then_piles_append_merges_chronologically_with_one_discount(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-103: plates then piles → len sum; chronological order; one discount in meta."""
    draft_id, body, plate_lines, pile_lines = _build_plates_then_piles_draft(
        mixed_client, monkeypatch
    )
    order_data = list(body["order_data"] or [])

    assert len(plate_lines) >= 1
    assert len(pile_lines) >= 1
    assert len(order_data) == len(plate_lines) + len(pile_lines)

    # Chronological: all plates from first cycle, then piles from second.
    types = [str(ln.get("product_type")) for ln in order_data]
    assert types == (["plates"] * len(plate_lines)) + (["piles"] * len(pile_lines))

    plate_ids = _assert_order_lines_have_identity(plate_lines, product_type="plates")
    pile_ids = _assert_order_lines_have_identity(pile_lines, product_type="piles")
    assert len(set(plate_ids) & set(pile_ids)) == 0

    meta = body["metadata"]
    assert float(meta["discount_percent"]) == _APPEND_DISCOUNT
    assert meta["client_name"] == _APPEND_CLIENT
    # Exactly one sticky discount field (not per-batch discounts).
    assert "discount_percent" in meta
    assert len(meta.get("append_batches") or []) == 2
    _assert_append_batches_cover_order(order_data, meta["append_batches"])
    assert meta["append_batches"][0]["product_type"] == "plates"
    assert meta["append_batches"][1]["product_type"] == "piles"
    assert meta["append_batches"][0]["line_ids"] == plate_ids
    assert meta["append_batches"][1]["line_ids"] == pile_ids

    # Draft still reachable after merge.
    again = mixed_client.get(f"/api/v1/commercial/drafts/{draft_id}")
    assert again.status_code == 200, again.text
    assert len(again.json()["order_data"]) == len(order_data)


def test_append_undo_last_removes_only_last_batch(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-103: POST append/undo-last drops last batch lines; prior batch remains."""
    draft_id, body, plate_lines, pile_lines = _build_plates_then_piles_draft(
        mixed_client, monkeypatch
    )
    plate_ids = [str(ln["line_id"]) for ln in plate_lines]
    pile_ids = [str(ln["line_id"]) for ln in pile_lines]
    assert len(body["metadata"]["append_batches"]) == 2

    undo = mixed_client.post(
        f"/api/v1/commercial/drafts/{draft_id}/append/undo-last",
    )
    assert undo.status_code == 200, undo.text
    undone = undo.json()
    order_data = list(undone["order_data"] or [])
    remaining_ids = [str(ln.get("line_id")) for ln in order_data]

    assert remaining_ids == plate_ids
    assert not any(pid in remaining_ids for pid in pile_ids)
    _assert_order_lines_have_identity(order_data, product_type="plates")

    batches = list(undone["metadata"].get("append_batches") or [])
    assert len(batches) == 1
    assert batches[0]["product_type"] == "plates"
    assert batches[0]["line_ids"] == plate_ids
    assert float(undone["metadata"]["discount_percent"]) == _APPEND_DISCOUNT
    assert undone["metadata"]["client_name"] == _APPEND_CLIENT


def test_delete_line_removes_one_line_id(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-103: DELETE /lines/{line_id} removes exactly that line."""
    draft_id, body, plate_lines, pile_lines = _build_plates_then_piles_draft(
        mixed_client, monkeypatch
    )
    order_before = list(body["order_data"] or [])
    assert len(order_before) >= 3
    target = plate_lines[0]
    target_id = str(target["line_id"])
    kept_ids = [
        str(ln["line_id"]) for ln in order_before if str(ln["line_id"]) != target_id
    ]

    delete = mixed_client.delete(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{target_id}",
    )
    assert delete.status_code == 200, delete.text
    after = delete.json()
    order_after = list(after["order_data"] or [])
    after_ids = [str(ln.get("line_id")) for ln in order_after]

    assert target_id not in after_ids
    assert after_ids == kept_ids
    assert len(order_after) == len(order_before) - 1

    # Piles from second batch still present; sticky header intact.
    assert any(ln.get("product_type") == "piles" for ln in order_after)
    assert float(after["metadata"]["discount_percent"]) == _APPEND_DISCOUNT
    assert after["metadata"]["client_name"] == _APPEND_CLIENT

    # append_batches no longer list the deleted line_id; empty batches removed.
    for batch in after["metadata"].get("append_batches") or []:
        assert target_id not in list(batch.get("line_ids") or [])
        assert list(batch.get("line_ids") or []), "empty append_batches must be dropped"


def test_plates_piles_plates_reentry_keeps_chronological_product_types(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-103: plates → piles → plates again keeps chronological type sequence."""
    draft_id, plates_before = _create_plates_result_draft(mixed_client, monkeypatch)
    n_plates_1 = len(plates_before)

    piles_body = _append_piles_cycle(mixed_client, draft_id)
    n_piles = sum(1 for ln in piles_body["order_data"] if ln.get("product_type") == "piles")
    assert n_piles >= 1

    again = _append_product_cycle(
        mixed_client,
        draft_id,
        product_type="plates",
        text=_PLATES_APPEND_TEXT_2,
    )
    order_data = list(again["order_data"] or [])
    types = [str(ln.get("product_type")) for ln in order_data]
    n_plates_2 = sum(1 for t in types if t == "plates") - n_plates_1
    assert n_plates_2 >= 1

    expected = (
        (["plates"] * n_plates_1)
        + (["piles"] * n_piles)
        + (["plates"] * n_plates_2)
    )
    assert types == expected, (
        f"re-entry must append chronologically, not regroup by type;\n"
        f"got={types!r}\nexpected={expected!r}"
    )
    assert len(again["metadata"].get("append_batches") or []) == 3
    _assert_append_batches_cover_order(order_data, again["metadata"]["append_batches"])
    assert [b["product_type"] for b in again["metadata"]["append_batches"]] == [
        "plates",
        "piles",
        "plates",
    ]


def test_second_ocr_append_after_type_reentry_keeps_sealed_first_plates(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-103: plates→piles→plates; second mode=append before calculate keeps sealed first plates.

    Regression: append+merged_cycle_text must not drop sealed earlier same-type batches.
    """
    draft_id, plates_before = _create_plates_result_draft(mixed_client, monkeypatch)
    first_plate_ids = [str(ln["line_id"]) for ln in plates_before]
    n_plates_1 = len(plates_before)
    assert n_plates_1 >= 1

    piles_body = _append_piles_cycle(mixed_client, draft_id)
    n_piles = sum(
        1 for ln in piles_body["order_data"] if ln.get("product_type") == "piles"
    )
    assert n_piles >= 1

    start = mixed_client.post(
        f"/api/v1/commercial/drafts/{draft_id}/append/start",
        json={"product_type": "plates"},
    )
    assert start.status_code == 200, start.text

    first_append = mixed_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/plates",
        data={"mode": "append", "text": _PLATES_APPEND_TEXT_2},
    )
    assert first_append.status_code == 200, first_append.text
    after_first = list(first_append.json()["order_data"] or [])
    after_first_ids = [str(ln.get("line_id")) for ln in after_first]
    assert after_first_ids[:n_plates_1] == first_plate_ids

    second_append = mixed_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/plates",
        data={"mode": "append", "text": "ПБ 60-12-8п 2"},
    )
    assert second_append.status_code == 200, second_append.text
    order_data = list(second_append.json()["order_data"] or [])
    order_ids = [str(ln.get("line_id")) for ln in order_data]
    types = [str(ln.get("product_type")) for ln in order_data]

    assert order_ids[:n_plates_1] == first_plate_ids, (
        "sealed first-cycle plates must survive second OCR append after type re-entry"
    )
    n_plates_cycle = sum(1 for t in types if t == "plates") - n_plates_1
    assert n_plates_cycle >= 1
    expected = (
        (["plates"] * n_plates_1)
        + (["piles"] * n_piles)
        + (["plates"] * n_plates_cycle)
    )
    assert types == expected, (
        f"chronological type sequence must be preserved;\n"
        f"got={types!r}\nexpected={expected!r}"
    )


def test_plates_then_fbs_append_merges_without_wiping_plates(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-103: plates → fbs mode=append keeps prior plates (cross-type beyond piles)."""
    draft_id, plates_before = _create_plates_result_draft(mixed_client, monkeypatch)
    plate_ids = [str(ln["line_id"]) for ln in plates_before]

    body = _append_product_cycle(
        mixed_client,
        draft_id,
        product_type="fbs",
        text=_FBS_APPEND_TEXT,
    )
    order_data = list(body["order_data"] or [])
    plate_lines = [ln for ln in order_data if ln.get("product_type") == "plates"]
    fbs_lines = [ln for ln in order_data if ln.get("product_type") == "fbs"]

    assert [str(ln["line_id"]) for ln in plate_lines] == plate_ids
    assert len(fbs_lines) >= 1
    types = [str(ln.get("product_type")) for ln in order_data]
    assert types == (["plates"] * len(plate_lines)) + (["fbs"] * len(fbs_lines))

    _assert_order_lines_have_identity(fbs_lines, product_type="fbs")
    assert len(body["metadata"].get("append_batches") or []) == 2
    _assert_append_batches_cover_order(order_data, body["metadata"]["append_batches"])
    assert body["metadata"]["append_batches"][0]["product_type"] == "plates"
    assert body["metadata"]["append_batches"][1]["product_type"] == "fbs"
    assert float(body["metadata"]["discount_percent"]) == _APPEND_DISCOUNT


def test_delete_missing_line_returns_russian_404_without_echoing_id(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-103: missing line_id → 404 with Russian detail, no echoed id."""
    draft_id, _plates = _create_plates_result_draft(mixed_client, monkeypatch)
    missing_id = "ln_does_not_exist_mna103"
    response = mixed_client.delete(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{missing_id}",
    )
    assert response.status_code == 404, response.text
    detail = response.json().get("detail", "")
    assert detail == "Строка не найдена."
    assert missing_id not in detail


# --- MNA-202: mixed discount + calculate validation ---


def _products_list_total(order_data: list[dict[str, Any]]) -> float:
    total = 0.0
    for line in order_data:
        unit = line.get("unit_price")
        if unit is None:
            continue
        total += float(unit) * float(line.get("qty") or 0)
    return total


def test_mixed_calculate_succeeds_when_cycle_product_type_is_piles(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-202: calculate_draft must not treat whole draft as piles-only via is_*_draft.

    After append/start(piles) metadata.product_type is the cycle type; mixed order still
    calculates when prerequisites are met.
    """
    draft_id, body, plate_lines, pile_lines = _build_plates_then_piles_draft(
        mixed_client, monkeypatch
    )
    assert len(plate_lines) >= 1
    assert len(pile_lines) >= 1
    assert body["metadata"]["product_type"] == "piles"
    assert body["wizard_state"]["current_step"] == "result"

    # Explicit recalculate while cycle type remains piles.
    recalc = mixed_client.post(f"/api/v1/commercial/drafts/{draft_id}/calculate")
    assert recalc.status_code == 200, recalc.text
    again = recalc.json()
    assert again["wizard_state"]["current_step"] == "result"
    assert len(again["order_data"]) == len(plate_lines) + len(pile_lines)
    types = [str(ln.get("product_type")) for ln in again["order_data"]]
    assert "plates" in types and "piles" in types


def test_mixed_discount_change_recomputes_totals_for_all_lines(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-202: one sticky discount_percent recomputes products for plates + piles."""
    draft_id, body, plate_lines, pile_lines = _build_plates_then_piles_draft(
        mixed_client, monkeypatch
    )
    assert len(plate_lines) >= 1 and len(pile_lines) >= 1

    # Start from 0% / no trip cost so the delta is unambiguous.
    zero = mixed_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={"discount_percent": 0.0, "logistics_cost": 0.0},
    )
    assert zero.status_code == 200, zero.text
    at_zero = zero.json()
    list_products = _products_list_total(at_zero["order_data"])
    assert list_products > 0.0
    assert float(at_zero["totals"]["total_with_vat"]) == pytest.approx(list_products)

    plate_list = _products_list_total(
        [ln for ln in at_zero["order_data"] if ln.get("product_type") == "plates"]
    )
    pile_list = _products_list_total(
        [ln for ln in at_zero["order_data"] if ln.get("product_type") == "piles"]
    )
    assert plate_list > 0.0 and pile_list > 0.0
    assert plate_list + pile_list == pytest.approx(list_products)

    discount_percent = 10.0
    updated = mixed_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={"discount_percent": discount_percent},
    )
    assert updated.status_code == 200, updated.text
    at_discount = updated.json()
    assert float(at_discount["metadata"]["discount_percent"]) == discount_percent

    expected = list_products * (1.0 - discount_percent / 100.0)
    assert float(at_discount["totals"]["total_with_vat"]) == pytest.approx(expected)

    # Savings must include both plate and pile contributions (not plates-only).
    savings = list_products - float(at_discount["totals"]["total_with_vat"])
    assert savings == pytest.approx(list_products * discount_percent / 100.0)
    assert savings > plate_list * discount_percent / 100.0


def test_mixed_calculate_blocked_by_unresolved_wide_plates_even_if_cycle_is_piles(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-202: unresolved wide plates still block calculate when cycle type is piles."""
    from app.services.draft_store import DraftStore

    draft_id, plates_before = _create_plates_result_draft(mixed_client, monkeypatch)
    assert len(plates_before) >= 1

    start = mixed_client.post(
        f"/api/v1/commercial/drafts/{draft_id}/append/start",
        json={"product_type": "piles"},
    )
    assert start.status_code == 200, start.text
    assert start.json()["metadata"]["product_type"] == "piles"
    assert any(
        ln.get("product_type") == "plates" for ln in start.json()["order_data"]
    )

    # Inject unresolved wide-plate gate while cycle is piles and plates remain in order.
    store = DraftStore()
    store.update_metadata(
        draft_id,
        wide_plate_lines=[{"id": "wide-mna202", "line": "ПБ 59-15-8п 1", "qty": 1}],
        wide_plates_resolved=False,
    )

    calc = mixed_client.post(f"/api/v1/commercial/drafts/{draft_id}/calculate")
    # Validation errors surface as generic 400 (MSG_VALIDATION); must not succeed.
    assert calc.status_code == 400, calc.text
    assert calc.json().get("detail") == "Проверьте введённые данные."


# --- MNA-304: save_offer update branch (saved_offer.kp_id) + status gate ----------


def _draft_with_saved_kp(
    *,
    kp_id: int,
    status: str = "в работе",
    draft_id: str = "draft-mna304",
) -> dict[str, Any]:
    """Minimal draft payload as returned by draft_store for resume/append save."""
    saved = {
        "kp_id": kp_id,
        "status": status,
        "mode": "database",
        "execution_terms": "",
        "saved_at": "2026-08-12T12:00:00",
    }
    return {
        "draft_id": draft_id,
        "order": {},
        "optimization": {},
        "order_data": [
            {
                "line_id": "ln_plate_resume",
                "product_type": "plates",
                "name": "ПБ 78-12-8п",
                "qty": 2,
                "length_m": 7.8,
                "width_m": 1.2,
                "unit_price": 1000.0,
                "concrete_grade": "М500",
            },
            {
                "line_id": "ln_pile_append",
                "product_type": "piles",
                "mark": "С120.35-12",
                "concrete_grade": "B25",
                "qty": 3,
                "unit_price": 40000.0,
            },
        ],
        "metadata": {
            "product_type": "piles",
            "client_name": "ООО Resume",
            "manager_name": "Иван",
            "discount_percent": 5.0,
            "logistics_cost": 0.0,
            "delivery_conditions": "",
            "payment_conditions": "",
            "wide_plates_resolved": True,
            "current_step": "result",
            "owner_user_id": 2,
            "saved_offer": saved,
            "resume_kp_id": kp_id,
        },
        "wizard_state": {
            "current_step": "result",
            "can_proceed_to": [],
            "next_required_action": "none",
            "validation_errors": [],
        },
        "files": [],
        "saved_offer": saved,
        "totals": {"subtotal": 1000.0, "vat_amount": 200.0, "total_with_vat": 1200.0},
        "offer_identity": {
            "offer_number": str(kp_id),
            "offer_date": "12.08.2026",
            "file_stem": f"kp_{kp_id}",
        },
    }


def test_save_offer_with_saved_kp_id_updates_same_id_not_create(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-304: when saved_offer.kp_id is set, save updates that KP (no new create)."""
    workflow = CommercialWorkflowService()
    existing_kp_id = 55
    draft = _draft_with_saved_kp(kp_id=existing_kp_id)
    fake_xlsx = tmp_path / "kp-append.xlsx"
    fake_xlsx.write_bytes(b"xlsx")

    generate_calls: list[Any] = []
    create_calls: list[dict[str, Any]] = []
    update_calls: list[dict[str, Any]] = []

    monkeypatch.setattr(workflow, "_load_draft_or_raise", lambda _draft_id: draft)
    monkeypatch.setattr(
        workflow,
        "generate_files",
        lambda draft_id, file_types=None: (
            generate_calls.append({"draft_id": draft_id, "file_types": file_types})
            or [{"kind": "xlsx", "filename": fake_xlsx.name}]
        ),
    )
    monkeypatch.setattr(
        workflow.export_service,
        "resolve_generated_file",
        lambda filename: fake_xlsx,
    )
    monkeypatch.setattr(workflow.draft_store, "update_metadata", lambda *a, **k: None)

    def fake_create(**kwargs: Any) -> int:
        create_calls.append(kwargs)
        return 999

    def fake_update(kp_id: int, order_data: list, **kwargs: Any) -> int:
        update_calls.append({"kp_id": kp_id, "order_data": order_data, **kwargs})
        return kp_id

    monkeypatch.setattr(workflow.kp_repository, "save_offer", fake_create)
    # Implementer may wire via repository method and/or offers_write.
    monkeypatch.setattr(
        workflow.kp_repository,
        "update_offer_from_order_data",
        fake_update,
        raising=False,
    )
    from core.kp import offers_write

    monkeypatch.setattr(
        offers_write,
        "update_kp_from_order_data",
        fake_update,
        raising=False,
    )
    from core.kp_persistence_service import KpPersistenceService

    monkeypatch.setattr(
        KpPersistenceService,
        "update_kp_from_order_data",
        fake_update,
        raising=False,
    )

    result = workflow.save_offer(
        "draft-mna304",
        execution_terms="",
        status="в работе",
        save_mode="database",
    )

    assert result["saved_offer"]["kp_id"] == existing_kp_id
    assert create_calls == [], "must not INSERT a new KP when saved_offer.kp_id is set"
    assert update_calls, "must call update_kp_from_order_data / update_offer_from_order_data"
    assert update_calls[0]["kp_id"] == existing_kp_id
    assert generate_calls, "R1: files must be regenerated on append save"
    assert result["saved_offer"]["status"] == "в работе"


def test_save_offer_with_saved_kp_id_rejects_when_status_not_in_progress(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-304 / R2: resume save rejected when saved KP status ≠ «в работе»."""
    workflow = CommercialWorkflowService()
    draft = _draft_with_saved_kp(kp_id=88, status="выполнено")
    fake_xlsx = tmp_path / "kp-blocked.xlsx"
    fake_xlsx.write_bytes(b"x")

    monkeypatch.setattr(workflow, "_load_draft_or_raise", lambda _draft_id: draft)
    monkeypatch.setattr(
        workflow,
        "generate_files",
        lambda _draft_id, file_types=None: [{"kind": "xlsx", "filename": fake_xlsx.name}],
    )
    monkeypatch.setattr(
        workflow.export_service,
        "resolve_generated_file",
        lambda filename: fake_xlsx,
    )
    monkeypatch.setattr(workflow.draft_store, "update_metadata", lambda *a, **k: None)
    monkeypatch.setattr(
        workflow.kp_repository,
        "save_offer",
        lambda **kwargs: 999,
    )

    with pytest.raises(ValueError, match="в работе"):
        workflow.save_offer(
            "draft-mna304",
            execution_terms="",
            status="в работе",
            save_mode="database",
        )


# --- MNA-601: hydrate draft from saved KP (Q1=C, status «в работе» only) ----------
#
# Assumed contracts:
#   Workflow: CommercialWorkflowService.hydrate_draft_from_saved_kp(kp_id, *, owner_user_id)
#             → CommercialDraftDetails-shaped dict
#             → order_data + header from KP; saved_offer.kp_id / resume_kp_id bound
#             → rejects when kp_meta.status ≠ «в работе» (ValueError)
#   HTTP:     POST /api/v1/commercial/archive/{kp_id}/resume  (see test_archive_endpoints)


def _kp_raw_for_hydrate(
    *,
    kp_id: int = 42,
    status: str = "в работе",
) -> dict[str, Any]:
    """Minimal KP raw dict as returned by get_kp_by_id / archive repository."""
    return {
        "kp_id": kp_id,
        "status": status,
        "customer_name": "ООО Resume Hydrate",
        "manager_name": "Иван Иванов",
        "discount_percent": 7.5,
        "logistics_cost": 18600.0,
        "delivery_conditions": "Доставка авто",
        "payment_conditions": "100% предоплата",
        "execution_terms": "",
        "creation_date": "12.08.2026",
        "product_type": "mixed",
        "plates": [
            {
                "line_id": "ln_plate_h1",
                "name": "ПБ 78-12-8п",
                "qty": 2,
                "length": 7.8,
                "width": 1.2,
                "unit_price": 1000.0,
                "concrete_grade": "М500",
                "position_number": 1,
            }
        ],
        "piles": [
            {
                "line_id": "ln_pile_h1",
                "mark": "С120.35-12",
                "concrete_grade": "B25",
                "qty": 3,
                "unit_price": 40000.0,
                "position_number": 2,
            }
        ],
    }


def test_hydrate_draft_from_saved_kp_loads_order_data_and_header(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-601: hydrate builds draft with KP order_data + sticky header fields."""
    assert hasattr(CommercialWorkflowService, "hydrate_draft_from_saved_kp"), (
        "CommercialWorkflowService.hydrate_draft_from_saved_kp missing (MNA-601)"
    )

    monkeypatch.setenv("DRAFTS_DIR", str(tmp_path / "drafts"))
    monkeypatch.setenv("OUTPUTS_DIR", str(tmp_path / "outputs"))
    get_settings.cache_clear()

    kp_id = 42
    kp_raw = _kp_raw_for_hydrate(kp_id=kp_id, status="в работе")

    from core.kp import offers_read

    monkeypatch.setattr(offers_read, "get_kp_by_id", lambda _id, db_path=None: kp_raw)
    monkeypatch.setattr(
        "app.services.commercial_workflow_service.get_kp_by_id",
        lambda _id, db_path=None: kp_raw,
        raising=False,
    )

    workflow = CommercialWorkflowService()
    # Prefer repository get_offer if implementer wires via KpRepository.
    monkeypatch.setattr(
        workflow.kp_repository,
        "get_offer",
        lambda _id: kp_raw,
        raising=False,
    )
    monkeypatch.setattr(
        workflow.manager_repository,
        "list_managers",
        lambda: [
            {
                "id": 1,
                "fio": "Иван Иванов",
                "contact_number": "+79990001122",
                "email": "ivan@test.local",
            }
        ],
    )

    result = workflow.hydrate_draft_from_saved_kp(kp_id, owner_user_id=1)

    assert result.get("draft_id"), "hydrate must return a new draft_id"
    order_data = list(result.get("order_data") or [])
    assert len(order_data) >= 2, f"expected plates+piles lines, got {order_data!r}"
    product_types = {str(line.get("product_type") or "") for line in order_data}
    assert "plates" in product_types
    assert "piles" in product_types

    metadata = result.get("metadata") or {}
    assert metadata.get("client_name") == "ООО Resume Hydrate"
    assert metadata.get("manager_name") == "Иван Иванов"
    assert float(metadata.get("discount_percent") or 0) == pytest.approx(7.5)
    assert float(metadata.get("logistics_cost") or 0) == pytest.approx(18600.0)
    assert metadata.get("delivery_conditions") == "Доставка авто"
    assert metadata.get("payment_conditions") == "100% предоплата"


def test_hydrate_draft_from_saved_kp_binds_saved_offer_and_resume_kp_id(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-601 / Q1=C: draft is bound to same kp_id via saved_offer + resume_kp_id."""
    assert hasattr(CommercialWorkflowService, "hydrate_draft_from_saved_kp"), (
        "CommercialWorkflowService.hydrate_draft_from_saved_kp missing (MNA-601)"
    )

    monkeypatch.setenv("DRAFTS_DIR", str(tmp_path / "drafts"))
    monkeypatch.setenv("OUTPUTS_DIR", str(tmp_path / "outputs"))
    get_settings.cache_clear()

    kp_id = 77
    kp_raw = _kp_raw_for_hydrate(kp_id=kp_id, status="в работе")

    from core.kp import offers_read

    monkeypatch.setattr(offers_read, "get_kp_by_id", lambda _id, db_path=None: kp_raw)
    monkeypatch.setattr(
        "app.services.commercial_workflow_service.get_kp_by_id",
        lambda _id, db_path=None: kp_raw,
        raising=False,
    )

    workflow = CommercialWorkflowService()
    monkeypatch.setattr(
        workflow.kp_repository,
        "get_offer",
        lambda _id: kp_raw,
        raising=False,
    )
    monkeypatch.setattr(
        workflow.manager_repository,
        "list_managers",
        lambda: [
            {
                "id": 1,
                "fio": "Иван Иванов",
                "contact_number": "+79990001122",
                "email": "ivan@test.local",
            }
        ],
    )

    result = workflow.hydrate_draft_from_saved_kp(kp_id, owner_user_id=1)

    saved = result.get("saved_offer")
    assert saved is not None, "saved_offer must be set after hydrate"
    assert saved.get("kp_id") == kp_id
    assert saved.get("status") == "в работе"

    metadata = result.get("metadata") or {}
    assert metadata.get("resume_kp_id") == kp_id
    # Ready for append loop (picker next); client sticky / skip.
    assert (result.get("wizard_state") or {}).get("current_step") == "result"


def test_hydrate_draft_from_saved_kp_rejects_when_status_not_in_progress(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-601 / R2: hydrate rejected when KP status ≠ «в работе»."""
    assert hasattr(CommercialWorkflowService, "hydrate_draft_from_saved_kp"), (
        "CommercialWorkflowService.hydrate_draft_from_saved_kp missing (MNA-601)"
    )

    monkeypatch.setenv("DRAFTS_DIR", str(tmp_path / "drafts"))
    monkeypatch.setenv("OUTPUTS_DIR", str(tmp_path / "outputs"))
    get_settings.cache_clear()

    kp_id = 88
    kp_raw = _kp_raw_for_hydrate(kp_id=kp_id, status="выполнено")

    from core.kp import offers_read

    monkeypatch.setattr(offers_read, "get_kp_by_id", lambda _id, db_path=None: kp_raw)
    monkeypatch.setattr(
        "app.services.commercial_workflow_service.get_kp_by_id",
        lambda _id, db_path=None: kp_raw,
        raising=False,
    )

    workflow = CommercialWorkflowService()
    monkeypatch.setattr(
        workflow.kp_repository,
        "get_offer",
        lambda _id: kp_raw,
        raising=False,
    )

    with pytest.raises(ValueError, match="в работе"):
        workflow.hydrate_draft_from_saved_kp(kp_id, owner_user_id=1)


@pytest.mark.parametrize(
    "blocked_status",
    ["На СГП", "отклонено", "в ожидании", "в архиве", "выполнено"],
)
def test_hydrate_draft_from_saved_kp_rejects_all_non_in_progress_statuses(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    blocked_status: str,
) -> None:
    """MNA-601 / R2: only «в работе» may hydrate; СГП and others blocked."""
    assert hasattr(CommercialWorkflowService, "hydrate_draft_from_saved_kp"), (
        "CommercialWorkflowService.hydrate_draft_from_saved_kp missing (MNA-601)"
    )

    monkeypatch.setenv("DRAFTS_DIR", str(tmp_path / "drafts"))
    monkeypatch.setenv("OUTPUTS_DIR", str(tmp_path / "outputs"))
    get_settings.cache_clear()

    kp_raw = _kp_raw_for_hydrate(kp_id=91, status=blocked_status)
    from core.kp import offers_read

    monkeypatch.setattr(offers_read, "get_kp_by_id", lambda _id, db_path=None: kp_raw)
    monkeypatch.setattr(
        "app.services.commercial_workflow_service.get_kp_by_id",
        lambda _id, db_path=None: kp_raw,
        raising=False,
    )

    workflow = CommercialWorkflowService()
    monkeypatch.setattr(
        workflow.kp_repository,
        "get_offer",
        lambda _id: kp_raw,
        raising=False,
    )

    with pytest.raises(ValueError, match="в работе"):
        workflow.hydrate_draft_from_saved_kp(91, owner_user_id=1)
