"""MNA-702: E2E multi-append flow + mono regression gate (SC-1..SC-9).

Locks the contract end-to-end by composing existing capabilities:
- draft append cycles (MNA-103) from ``test_commercial_draft_append``
- unified / mono export (MNA-401/402) from ``test_commercial_export_mixed``
- save / resume / update same ``kp_id`` (MNA-304 / MNA-601)
- PB-only logistics (MNA-001 / MNA-201)
- archive multi badges (MNA-602)
- production mixed-with-plates (MNA-701)

Flow covered: plates→piles→plates create; append to saved; export; undo/delete;
PB logistics; mono unchanged.
"""

from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any

import pandas as pd
import pytest
from fastapi.testclient import TestClient

from app.core.settings import get_settings
from app.repositories.kp_repository import KpRepository
from app.schemas.commercial import WizardStepId
from app.services.commercial_calculation_service import CommercialCalculationService
from app.services.commercial_wizard_step_service import CommercialWizardStepService
from app.services.draft_store import DraftStore
from app.services.production_service import ProductionService
from core.cargo_delivery_pricing import (
    delivery_service_charge_rub,
    total_order_cargo_weight_kg,
)
from core.commercial_line_format import format_line_name
from core.commercial_offer import is_unified_commercial_document
from core.kp_db_schema import init_schema
from tests.test_commercial_draft_append import (
    _APPEND_CLIENT,
    _APPEND_DISCOUNT,
    _PILES_APPEND_TEXT,
    _PLATES_APPEND_TEXT_2,
    _append_piles_cycle,
    _append_product_cycle,
    _assert_append_batches_cover_order,
    _assert_order_lines_have_identity,
    _auth_client,
    _create_plates_result_draft,
    _mock_manager_lookup,
    _products_list_total,
    _setup_mixed_price_env,
    _setup_pile_price_env,
    _setup_plate_price_env,
)
from tests.test_commercial_export_mixed import (
    MONO_PLATES_HEADERS,
    MONO_PILES_HEADERS,
    UNIFIED_HEADERS,
    _body_column,
    _plate,
    _pile,
    _xlsx_header_and_body,
)

# --- fixtures / local helpers -------------------------------------------------


@pytest.fixture()
def flow_client(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> TestClient:
    """Auth client with plates+piles prices and initialized plita.db for save/resume."""
    _setup_mixed_price_env(monkeypatch, tmp_path)
    plita_db = tmp_path / "plita.db"
    init_schema(str(plita_db))
    with sqlite3.connect(str(plita_db)) as conn:
        conn.execute(
            "INSERT INTO managers (id, fio, contact_number, email) "
            "VALUES (1, 'Tester', '+79990001122', 'tester@test.local')"
        )
        conn.commit()
    get_settings.cache_clear()
    return _auth_client(monkeypatch)


def _build_plates_piles_plates(
    client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> tuple[str, dict[str, Any]]:
    """plates → piles → plates again; return (draft_id, final body)."""
    draft_id, plates_before = _create_plates_result_draft(client, monkeypatch)
    n_plates_1 = len(plates_before)
    assert n_plates_1 >= 1

    piles_body = _append_piles_cycle(client, draft_id)
    n_piles = sum(
        1 for ln in piles_body["order_data"] if ln.get("product_type") == "piles"
    )
    assert n_piles >= 1

    again = _append_product_cycle(
        client,
        draft_id,
        product_type="plates",
        text=_PLATES_APPEND_TEXT_2,
    )
    types = [str(ln.get("product_type")) for ln in again["order_data"]]
    n_plates_2 = sum(1 for t in types if t == "plates") - n_plates_1
    assert n_plates_2 >= 1
    expected = (
        (["plates"] * n_plates_1)
        + (["piles"] * n_piles)
        + (["plates"] * n_plates_2)
    )
    assert types == expected
    assert len(again["metadata"].get("append_batches") or []) == 3
    _assert_append_batches_cover_order(
        again["order_data"], again["metadata"]["append_batches"]
    )
    return draft_id, again


def _save_draft_kp(client: TestClient, draft_id: str) -> int:
    """Persist draft as KP with status «в работе» (production path)."""
    save = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/save",
        json={"mode": "database", "execution_terms_input": "14 дней"},
    )
    assert save.status_code == 200, save.text
    kp_id = int(save.json()["saved_offer"]["kp_id"])
    assert kp_id >= 1
    assert save.json()["saved_offer"]["status"] == "в работе"
    return kp_id


def _save_draft_kp_archive(client: TestClient, draft_id: str) -> int:
    """Persist draft as KP with status «в архиве» (resume/update gate)."""
    save = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/save",
        json={"mode": "archive", "execution_terms_input": "14 дней"},
    )
    assert save.status_code == 200, save.text
    kp_id = int(save.json()["saved_offer"]["kp_id"])
    assert kp_id >= 1
    assert save.json()["saved_offer"]["status"] == "в архиве"
    return kp_id


def _wizard_step_service() -> CommercialWizardStepService:
    """Mirror test_commercial_wizard_step_service construction (required kwargs)."""
    return CommercialWizardStepService(
        calculation_service=CommercialCalculationService(),
        draft_store=DraftStore(),
    )


def _order_with_resolvable_pile_weights(
    order_data: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Stamp length_m/width_m on piles so resolve_kp_line_weight_kg yields >0.

    Pile commercial lines carry mark/qty only; cargo helper uses the plate formula
    (length_m × width_m). Dims here make unfiltered kg > plates-only for SC-5.
    """
    out: list[dict[str, Any]] = []
    for ln in order_data:
        item = dict(ln)
        if item.get("product_type") == "piles":
            if not float(item.get("length_m") or 0):
                item["length_m"] = 12.0
            if not float(item.get("width_m") or 0):
                item["width_m"] = 0.35
        out.append(item)
    return out


def _plita_db_path() -> Path:
    return Path(get_settings().plita_db_path)


def _meta_product_type(kp_id: int) -> str:
    with sqlite3.connect(str(_plita_db_path())) as conn:
        cur = conn.cursor()
        cur.execute("SELECT product_type FROM kp_meta WHERE kp_id = ?", (kp_id,))
        row = cur.fetchone()
    assert row is not None, f"kp_meta missing for kp_id={kp_id}"
    return str(row[0])


def _generate_xlsx_bytes(client: TestClient, draft_id: str) -> bytes:
    files = client.post(
        f"/api/v1/commercial/drafts/{draft_id}/generate-files",
        json={"file_types": ["xlsx", "pdf"]},
    )
    assert files.status_code == 200, files.text
    kinds = {item["kind"] for item in files.json()["files"]}
    assert "xlsx" in kinds
    assert "pdf" in kinds
    xlsx_name = next(
        item["filename"] for item in files.json()["files"] if item["kind"] == "xlsx"
    )
    download = client.get(
        f"/api/v1/commercial/files/{xlsx_name}",
        params={"draft_id": draft_id},
    )
    assert download.status_code == 200, download.text
    assert download.content[:2] == b"PK"
    return download.content


def _xlsx_headers_and_type_column(content: bytes) -> tuple[list[str], list[str], list[str]]:
    from io import BytesIO

    df = pd.read_excel(BytesIO(content), sheet_name="КП", header=None)
    header_idx = None
    headers: list[str] = []
    for i, row in df.iterrows():
        vals = [str(v).strip() for v in row.tolist() if pd.notna(v)]
        if vals and vals[0] == "№":
            header_idx = int(i)
            raw = [v if pd.notna(v) else "" for v in row.tolist()]
            while raw and raw[-1] == "":
                raw.pop()
            headers = [str(c).strip() if c != "" else "" for c in raw]
            break
    assert header_idx is not None, "table header row with № not found in XLSX"
    body = df.iloc[header_idx + 1 :].copy()
    body.columns = range(len(body.columns))
    names = _body_column(body, headers, "Наименование")
    types = _body_column(body, headers, "Тип") if "Тип" in headers else []
    return headers, names, types


# --- SC-1: plates→piles→plates → one kp_id ------------------------------------


def test_sc1_plates_piles_plates_create_saves_one_mixed_kp(
    flow_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """SC-1: ≥2 cycles (plates→piles→plates) save to a single mixed kp_id."""
    _mock_manager_lookup(monkeypatch)
    draft_id, body = _build_plates_piles_plates(flow_client, monkeypatch)
    assert float(body["metadata"]["discount_percent"]) == _APPEND_DISCOUNT
    assert body["metadata"]["client_name"] == _APPEND_CLIENT

    kp_id = _save_draft_kp(flow_client, draft_id)
    assert _meta_product_type(kp_id) == "mixed"

    with sqlite3.connect(str(_plita_db_path())) as conn:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM kp_plates WHERE kp_id = ?", (kp_id,))
        n_plates = int(cur.fetchone()[0])
        cur.execute("SELECT COUNT(*) FROM kp_piles WHERE kp_id = ?", (kp_id,))
        n_piles = int(cur.fetchone()[0])
        cur.execute("SELECT COUNT(*) FROM KP_offers WHERE kp_id = ?", (kp_id,))
        assert int(cur.fetchone()[0]) == 1

    assert n_plates >= 2
    assert n_piles >= 1
    order_plates = sum(
        1 for ln in body["order_data"] if ln.get("product_type") == "plates"
    )
    order_piles = sum(
        1 for ln in body["order_data"] if ln.get("product_type") == "piles"
    )
    assert n_plates == order_plates
    assert n_piles == order_piles


# --- SC-2: unified export order / type / grade / one discount -----------------


def test_sc2_unified_export_after_multi_append(
    flow_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """SC-2: multi/append → unified XLSX; chronological type; grade in name; one discount."""
    _mock_manager_lookup(monkeypatch)
    draft_id, body = _build_plates_piles_plates(flow_client, monkeypatch)
    order_data = list(body["order_data"] or [])
    batches = list(body["metadata"].get("append_batches") or [])

    assert is_unified_commercial_document(order_data, append_batches=batches) is True

    content = _generate_xlsx_bytes(flow_client, draft_id)
    headers, names, types = _xlsx_headers_and_type_column(content)
    assert headers == UNIFIED_HEADERS

    # Drop delivery row if present; product rows follow append order.
    product_types = [t for t in types if t]
    assert product_types[0] == "Плиты"
    assert "Сваи" in product_types
    assert product_types[-1] == "Плиты"
    # No regrouping: first contiguous plates block, then piles, then plates again.
    first_pile = product_types.index("Сваи")
    assert all(t == "Плиты" for t in product_types[:first_pile])
    assert "Плиты" in product_types[first_pile + 1 :]

    pile_lines = [ln for ln in order_data if ln.get("product_type") == "piles"]
    assert pile_lines
    expected_pile_name = format_line_name(pile_lines[0])
    assert any(expected_pile_name in n or "(B25)" in n for n in names), names

    assert float(body["metadata"]["discount_percent"]) == _APPEND_DISCOUNT
    list_total = _products_list_total(order_data)
    discounted = list_total * (1.0 - _APPEND_DISCOUNT / 100.0)
    assert float(body["totals"]["total_with_vat"]) == pytest.approx(discounted)


# --- SC-3: skip client on cycle ≥2 --------------------------------------------


def test_sc3_second_cycle_skips_client_step(
    flow_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """SC-3: after first result, append/start → wizard omits client (proceed to result)."""
    draft_id, _plates = _create_plates_result_draft(flow_client, monkeypatch)

    start = flow_client.post(
        f"/api/v1/commercial/drafts/{draft_id}/append/start",
        json={"product_type": "piles"},
    )
    assert start.status_code == 200, start.text
    body = start.json()
    meta = body["metadata"]
    assert meta["client_name"] == _APPEND_CLIENT
    assert meta["product_type"] == "piles"
    assert body["wizard_state"]["current_step"] == "piles"

    wizard = _wizard_step_service()
    assert wizard.should_skip_client_step(meta) is True
    order = wizard.wizard_step_order(meta)
    assert WizardStepId.client not in order
    assert body["wizard_state"]["can_proceed_to"] == ["result"]


def test_pile_grades_bulk_only_touches_current_cycle_not_sealed_plates(
    flow_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Append piles after plates: bulk grade must not wipe sealed plate lines."""
    draft_id, plate_lines = _create_plates_result_draft(flow_client, monkeypatch)
    sealed_plate_ids = {str(ln["line_id"]) for ln in plate_lines}

    start = flow_client.post(
        f"/api/v1/commercial/drafts/{draft_id}/append/start",
        json={"product_type": "piles"},
    )
    assert start.status_code == 200, start.text

    piles = flow_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/piles",
        data={"mode": "append", "text": _PILES_APPEND_TEXT},
    )
    assert piles.status_code == 200, piles.text
    before = piles.json()
    assert sealed_plate_ids.issubset({str(ln["line_id"]) for ln in before["order_data"]})

    grades = flow_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/piles/grades",
        json={"concrete_grade": "B20"},
    )
    assert grades.status_code == 200, grades.text
    after = grades.json()
    after_ids = {str(ln["line_id"]) for ln in after["order_data"]}
    assert sealed_plate_ids.issubset(after_ids)
    plate_rows = [ln for ln in after["order_data"] if ln.get("product_type") == "plates"]
    pile_rows = [ln for ln in after["order_data"] if ln.get("product_type") == "piles"]
    assert plate_rows
    assert pile_rows
    assert all(ln.get("concrete_grade") == "B20" for ln in pile_rows)
    assert all(str(ln.get("append_batch_id") or "").strip() for ln in plate_rows)


# --- SC-4: undo last batch + delete line --------------------------------------


def test_sc4_undo_last_and_delete_line_in_multi_flow(
    flow_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """SC-4: undo last batch and delete by line_id leave draft on result."""
    draft_id, body = _build_plates_piles_plates(flow_client, monkeypatch)
    before_ids = [str(ln["line_id"]) for ln in body["order_data"]]
    last_batch = body["metadata"]["append_batches"][-1]
    last_ids = {str(x) for x in last_batch["line_ids"]}
    assert last_batch["product_type"] == "plates"
    assert last_ids

    undo = flow_client.post(
        f"/api/v1/commercial/drafts/{draft_id}/append/undo-last",
    )
    assert undo.status_code == 200, undo.text
    after_undo = undo.json()
    assert after_undo["wizard_state"]["current_step"] == "result"
    remaining = [str(ln["line_id"]) for ln in after_undo["order_data"]]
    assert last_ids.isdisjoint(set(remaining))
    assert len(after_undo["metadata"]["append_batches"]) == 2
    assert set(remaining) == set(before_ids) - last_ids

    target = remaining[-1]
    deleted = flow_client.delete(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{target}",
    )
    assert deleted.status_code == 200, deleted.text
    after_del = deleted.json()
    assert after_del["wizard_state"]["current_step"] == "result"
    ids_after = [str(ln["line_id"]) for ln in after_del["order_data"]]
    assert target not in ids_after
    for batch in after_del["metadata"].get("append_batches") or []:
        assert target not in list(batch.get("line_ids") or [])


# --- SC-5: PB-only logistics --------------------------------------------------


def test_sc5_mixed_logistics_uses_plates_cargo_only(
    flow_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """SC-5: trip cost on mixed draft → delivery from plates kg only."""
    _mock_manager_lookup(monkeypatch)
    draft_id, _plates = _create_plates_result_draft(flow_client, monkeypatch)
    body = _append_piles_cycle(flow_client, draft_id, piles_text=_PILES_APPEND_TEXT)
    order_data = list(body["order_data"] or [])
    assert any(ln.get("product_type") == "plates" for ln in order_data)
    assert any(ln.get("product_type") == "piles" for ln in order_data)

    trip = 5000.0
    zero = flow_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={"discount_percent": 0.0, "logistics_cost": 0.0},
    )
    assert zero.status_code == 200, zero.text
    products = float(zero.json()["totals"]["total_with_vat"])
    assert products == pytest.approx(_products_list_total(order_data))

    with_trip = flow_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={"logistics_cost": trip},
    )
    assert with_trip.status_code == 200, with_trip.text
    # API delivery uses real draft lines (plates-only); stamp pile dims only for
    # the unfiltered vs plates-only cargo sanity (piles lack length_m/width_m).
    plates_kg = total_order_cargo_weight_kg(order_data, product_types={"plates"})
    weighted = _order_with_resolvable_pile_weights(order_data)
    all_kg = total_order_cargo_weight_kg(weighted)
    assert plates_kg > 0.0
    assert all_kg > plates_kg
    expected_delivery = delivery_service_charge_rub(trip, plates_kg)
    assert expected_delivery == trip  # one trip for sample qty
    assert float(with_trip.json()["totals"]["total_with_vat"]) == pytest.approx(
        products + expected_delivery
    )
    # Sanity: wrong (all-lines) delivery would differ when piles add mass.
    wrong = delivery_service_charge_rub(trip, all_kg)
    if wrong != expected_delivery:
        assert float(with_trip.json()["totals"]["total_with_vat"]) != pytest.approx(
            products + wrong
        )


def test_sc5_piles_only_no_delivery_despite_trip_cost(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """SC-5: no plates → delivery stays 0 even if logistics_cost set."""
    _setup_pile_price_env(monkeypatch, tmp_path)
    client = _auth_client(monkeypatch)
    _mock_manager_lookup(monkeypatch)

    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "piles", "text": _PILES_APPEND_TEXT},
    )
    assert create.status_code == 200, create.text
    draft_id = create.json()["draft_id"]
    meta = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={
            "manager_id": 1,
            "client_name": "ООО Только сваи",
            "discount_percent": 0,
            "conditions_mode": "standard",
            "logistics_cost": 9000.0,
        },
    )
    assert meta.status_code == 200, meta.text
    calc = client.post(f"/api/v1/commercial/drafts/{draft_id}/calculate")
    assert calc.status_code == 200, calc.text
    body = calc.json()
    order_data = list(body["order_data"] or [])
    assert all(ln.get("product_type") == "piles" for ln in order_data)
    assert total_order_cargo_weight_kg(order_data, product_types={"plates"}) == 0.0
    assert float(body["totals"]["total_with_vat"]) == pytest.approx(
        _products_list_total(order_data)
    )

    headers, xbody = _xlsx_header_and_body(order_data, logistics_cost=9000.0)
    assert headers == MONO_PILES_HEADERS
    names = _body_column(xbody, headers, "Наименование")
    assert not any("доставк" in n.lower() for n in names), names


# --- SC-6: append to saved KP (Q1=C) same kp_id -------------------------------


def test_sc6_append_to_saved_kp_keeps_same_kp_id(
    flow_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """SC-6: archive plates KP → resume → append piles → archive-save same kp_id."""
    _mock_manager_lookup(monkeypatch)
    draft_id, plates = _create_plates_result_draft(flow_client, monkeypatch)
    plate_ids = _assert_order_lines_have_identity(plates, product_type="plates")
    kp_id = _save_draft_kp_archive(flow_client, draft_id)
    assert _meta_product_type(kp_id) == "plates"

    resume = flow_client.post(f"/api/v1/commercial/archive/{kp_id}/resume")
    assert resume.status_code == 200, resume.text
    resumed = resume.json()
    resume_draft_id = resumed["draft_id"]
    assert resumed["saved_offer"]["kp_id"] == kp_id
    assert resumed["saved_offer"]["status"] == "в архиве"
    assert resumed["metadata"]["resume_kp_id"] == kp_id
    assert resumed["wizard_state"]["current_step"] == "result"
    wizard = _wizard_step_service()
    assert wizard.should_skip_client_step(resumed["metadata"]) is True

    appended = _append_piles_cycle(flow_client, resume_draft_id)
    order_data = list(appended["order_data"] or [])
    assert [str(ln["line_id"]) for ln in order_data if ln.get("product_type") == "plates"][
        : len(plate_ids)
    ] == plate_ids
    assert any(ln.get("product_type") == "piles" for ln in order_data)

    save2 = flow_client.post(
        f"/api/v1/commercial/drafts/{resume_draft_id}/save",
        json={"mode": "archive", "execution_terms_input": "14 дней"},
    )
    assert save2.status_code == 200, save2.text
    assert int(save2.json()["saved_offer"]["kp_id"]) == kp_id
    assert save2.json()["saved_offer"]["status"] == "в архиве"
    assert _meta_product_type(kp_id) == "mixed"

    with sqlite3.connect(str(_plita_db_path())) as conn:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM KP_offers")
        assert int(cur.fetchone()[0]) == 1
        cur.execute("SELECT COUNT(*) FROM kp_piles WHERE kp_id = ?", (kp_id,))
        assert int(cur.fetchone()[0]) >= 1
        cur.execute("SELECT status FROM kp_meta WHERE kp_id = ?", (kp_id,))
        assert str(cur.fetchone()[0]) == "в архиве"


def test_sc6_resume_blocked_when_status_in_work(
    flow_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """SC-6 / archive-edit: resume blocked for status «в работе» (production)."""
    _mock_manager_lookup(monkeypatch)
    draft_id, _plates = _create_plates_result_draft(flow_client, monkeypatch)
    kp_id = _save_draft_kp(flow_client, draft_id)

    resume = flow_client.post(f"/api/v1/commercial/archive/{kp_id}/resume")
    assert resume.status_code in (400, 409), resume.text


def test_sc6_resume_blocked_when_status_not_archived(
    flow_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """SC-6: resume/append only for status «в архиве»."""
    _mock_manager_lookup(monkeypatch)
    draft_id, _plates = _create_plates_result_draft(flow_client, monkeypatch)
    kp_id = _save_draft_kp_archive(flow_client, draft_id)

    with sqlite3.connect(str(_plita_db_path())) as conn:
        conn.execute(
            "UPDATE kp_meta SET status = ? WHERE kp_id = ?",
            ("выполнено", kp_id),
        )
        conn.commit()

    resume = flow_client.post(f"/api/v1/commercial/archive/{kp_id}/resume")
    assert resume.status_code in (400, 409), resume.text


# --- SC-7: archive multi badges -----------------------------------------------


def test_sc7_archive_list_exposes_multiple_product_type_badges(
    flow_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """SC-7 / Q3: after mixed save, archive list row carries N product_types badges."""
    _mock_manager_lookup(monkeypatch)
    draft_id, _body = _build_plates_piles_plates(flow_client, monkeypatch)
    kp_id = _save_draft_kp(flow_client, draft_id)
    assert _meta_product_type(kp_id) == "mixed"

    # «в работе» lands in in_production section (see offers_read grouping).
    listing = flow_client.get(
        "/api/v1/commercial/archive",
        params={"section": "in_production"},
    )
    assert listing.status_code == 200, listing.text
    rows = listing.json()
    row = next(r for r in rows if int(r["kp_id"]) == kp_id)
    assert "product_types" in row
    assert set(row["product_types"]) == {"plates", "piles"}
    assert len(row["product_types"]) == 2


# --- SC-8: mono without append — no regression --------------------------------


def test_sc8_mono_plates_export_unchanged(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """SC-8 / R3: mono one-shot plates keep classic headers (not unified)."""
    _setup_plate_price_env(monkeypatch, tmp_path)
    client = _auth_client(monkeypatch)
    _mock_manager_lookup(monkeypatch)

    draft_id, plates = _create_plates_result_draft(
        client,
        monkeypatch,
        client_name="ООО Mono",
        discount_percent=0.0,
    )
    assert len(plates) >= 1
    assert is_unified_commercial_document(plates) is False

    content = _generate_xlsx_bytes(client, draft_id)
    headers, names, types = _xlsx_headers_and_type_column(content)
    assert headers == MONO_PLATES_HEADERS
    assert "Тип" not in headers
    assert types == []
    assert any("ПБ" in n for n in names)


def test_sc8_mono_domain_helpers_still_classic() -> None:
    """SC-8 characterization: domain helpers agree mono ≠ unified."""
    order = [_plate(line_id="a"), _plate(line_id="b", name="ПБ 56-6-8п", mark="ПБ 56-6-8п")]
    assert is_unified_commercial_document(order) is False
    headers, _body = _xlsx_header_and_body(order)
    assert headers == MONO_PLATES_HEADERS

    mixed = [_plate(line_id="p1"), _pile(line_id="s1")]
    assert is_unified_commercial_document(mixed) is True
    headers_m, _ = _xlsx_header_and_body(mixed)
    assert headers_m == UNIFIED_HEADERS


# --- SC-9: production only plates (mixed-with-plates OK) ----------------------


def test_sc9_production_includes_saved_mixed_with_plates_only(
    flow_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """SC-9: saved mixed KP is a production candidate; plate rows only in candidates."""
    _mock_manager_lookup(monkeypatch)
    draft_id, body = _build_plates_piles_plates(flow_client, monkeypatch)
    kp_id = _save_draft_kp(flow_client, draft_id)
    assert _meta_product_type(kp_id) == "mixed"

    db_path = str(_plita_db_path())
    repo = KpRepository(db_path=db_path)
    items = repo.list_kps_in_production()
    by_id = {int(item["kp_id"]): item for item in items}
    assert kp_id in by_id, "mixed-with-plates must appear in production candidates"

    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "UPDATE kp_plates SET status = 'в производстве' WHERE kp_id = ?",
            (kp_id,),
        )
        conn.commit()
        cur = conn.cursor()
        cur.execute("SELECT id FROM kp_plates WHERE kp_id = ? ORDER BY id", (kp_id,))
        plate_ids = [int(r[0]) for r in cur.fetchall()]
        cur.execute("SELECT COUNT(*) FROM kp_piles WHERE kp_id = ?", (kp_id,))
        assert int(cur.fetchone()[0]) >= 1

    service = ProductionService(kp_repository=repo)
    candidates = service.list_kp_candidates()
    item = next(c for c in candidates["items"] if int(c["kp_id"]) == kp_id)
    plates = item["plates"]
    assert {int(p["id"]) for p in plates} == set(plate_ids)
    order_pile_marks = {
        str(ln.get("mark") or ln.get("name") or "")
        for ln in body["order_data"]
        if ln.get("product_type") == "piles"
    }
    for p in plates:
        assert p.get("plate_name") not in order_pile_marks


# --- Integrated smoke: create → export → undo → logistics → mono gate ---------


def test_e2e_multi_append_create_export_undo_logistics_mono_gate(
    flow_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """MNA-702 smoke: full create path + export + undo + PB logistics + mono gate."""
    _mock_manager_lookup(monkeypatch)

    # Create multi
    draft_id, body = _build_plates_piles_plates(flow_client, monkeypatch)
    assert len(body["metadata"]["append_batches"]) == 3

    # Export unified
    content = _generate_xlsx_bytes(flow_client, draft_id)
    headers, _names, types = _xlsx_headers_and_type_column(content)
    assert headers == UNIFIED_HEADERS
    assert "Плиты" in types and "Сваи" in types

    # Undo last plates batch → still result
    undo = flow_client.post(
        f"/api/v1/commercial/drafts/{draft_id}/append/undo-last",
    )
    assert undo.status_code == 200, undo.text
    assert undo.json()["wizard_state"]["current_step"] == "result"
    assert len(undo.json()["metadata"]["append_batches"]) == 2

    # PB logistics still applies on remaining mixed (plates+piles)
    trip = 1000.0
    patched = flow_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/meta",
        json={"discount_percent": 0.0, "logistics_cost": trip},
    )
    assert patched.status_code == 200, patched.text
    order_data = list(patched.json()["order_data"] or [])
    plates_kg = total_order_cargo_weight_kg(order_data, product_types={"plates"})
    expected = _products_list_total(order_data) + delivery_service_charge_rub(
        trip, plates_kg
    )
    assert float(patched.json()["totals"]["total_with_vat"]) == pytest.approx(expected)

    # Save one kp_id
    kp_id = _save_draft_kp(flow_client, draft_id)
    assert _meta_product_type(kp_id) == "mixed"

    # Mono regression gate (domain, no I/O)
    mono = [_plate(line_id="mono1")]
    assert is_unified_commercial_document(mono) is False
    mono_headers, _ = _xlsx_header_and_body(mono)
    assert mono_headers == MONO_PLATES_HEADERS
