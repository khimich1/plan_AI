"""PATCH draft line oneoff unit_price (contractual price on this draft only)."""

from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any

import pytest
from fastapi.testclient import TestClient

from app.services.commercial_calculation_service import CommercialCalculationService
from tests.test_commercial_draft_append import (
    _auth_client,
    _patch_price_db_path,
    _setup_pile_price_env,
    _setup_mixed_price_env,
)


def _line_by_id(order_data: list[dict[str, Any]], line_id: str) -> dict[str, Any]:
    for line in order_data:
        if str(line.get("line_id") or "") == line_id:
            return line
    raise AssertionError(f"line {line_id!r} not in order_data")


def _pile_prices_count(db_path: Path) -> int:
    conn = sqlite3.connect(db_path)
    try:
        return int(conn.execute("SELECT COUNT(*) FROM pile_prices").fetchone()[0])
    finally:
        conn.close()


@pytest.fixture()
def piles_client_db(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> tuple[TestClient, Path]:
    _setup_pile_price_env(monkeypatch, tmp_path)
    pb_db = tmp_path / "pb.db"
    _patch_price_db_path(monkeypatch, pb_db)
    return _auth_client(monkeypatch), pb_db


@pytest.fixture()
def mixed_client(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> TestClient:
    _setup_mixed_price_env(monkeypatch, tmp_path)
    return _auth_client(monkeypatch)


def _create_unpriced_pile_draft(client: TestClient) -> tuple[str, dict[str, Any]]:
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "piles", "text": "С110.30-6 B25 2"},
    )
    assert create.status_code == 200, create.text
    body = create.json()
    order = list(body.get("order_data") or [])
    assert order, body
    line = order[0]
    assert line.get("unit_price") is None
    return str(body["draft_id"]), line


def test_oneoff_on_unpriced_pile_sets_flag_and_clears_gate(
    piles_client_db: tuple[TestClient, Path],
) -> None:
    client, pb_db = piles_client_db
    count_before = _pile_prices_count(pb_db)
    draft_id, line = _create_unpriced_pile_draft(client)
    line_id = str(line["line_id"])

    response = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{line_id}",
        json={"unit_price": 27585.43},
    )
    assert response.status_code == 200, response.text
    updated = _line_by_id(list(response.json()["order_data"] or []), line_id)
    assert float(updated["unit_price"]) == pytest.approx(27585.43)
    assert updated.get("price_source") == "oneoff"
    assert float(updated["line_total"]) == pytest.approx(27585.43 * 2)
    labels = CommercialCalculationService().unpriced_position_labels(
        list(response.json()["order_data"] or [])
    )
    assert labels == []
    errors = list(response.json().get("wizard_state", {}).get("validation_errors") or [])
    assert not any("Нет цен" in str(item) for item in errors)
    assert _pile_prices_count(pb_db) == count_before


def test_oneoff_edit_then_clear_null(
    piles_client_db: tuple[TestClient, Path],
) -> None:
    client, _pb_db = piles_client_db
    draft_id, line = _create_unpriced_pile_draft(client)
    line_id = str(line["line_id"])

    first = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{line_id}",
        json={"unit_price": 1000},
    )
    assert first.status_code == 200, first.text
    second = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{line_id}",
        json={"unit_price": 2000},
    )
    assert second.status_code == 200, second.text
    edited = _line_by_id(list(second.json()["order_data"] or []), line_id)
    assert float(edited["unit_price"]) == pytest.approx(2000)
    assert edited.get("price_source") == "oneoff"

    cleared = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{line_id}",
        json={"unit_price": None},
    )
    assert cleared.status_code == 200, cleared.text
    empty = _line_by_id(list(cleared.json()["order_data"] or []), line_id)
    assert empty.get("unit_price") is None
    assert empty.get("price_source") not in {"oneoff"}
    labels = CommercialCalculationService().unpriced_position_labels(
        list(cleared.json()["order_data"] or [])
    )
    assert labels


def test_qty_omit_does_not_clear_oneoff(
    piles_client_db: tuple[TestClient, Path],
) -> None:
    client, _pb_db = piles_client_db
    draft_id, line = _create_unpriced_pile_draft(client)
    line_id = str(line["line_id"])
    client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{line_id}",
        json={"unit_price": 1500},
    )
    qty_only = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{line_id}",
        json={"qty": 4},
    )
    assert qty_only.status_code == 200, qty_only.text
    updated = _line_by_id(list(qty_only.json()["order_data"] or []), line_id)
    assert int(updated["qty"]) == 4
    assert float(updated["unit_price"]) == pytest.approx(1500)
    assert updated.get("price_source") == "oneoff"
    assert float(updated["line_total"]) == pytest.approx(6000)


def test_qty_and_oneoff_together(
    piles_client_db: tuple[TestClient, Path],
) -> None:
    client, _pb_db = piles_client_db
    draft_id, line = _create_unpriced_pile_draft(client)
    line_id = str(line["line_id"])
    response = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{line_id}",
        json={"qty": 3, "unit_price": 1100},
    )
    assert response.status_code == 200, response.text
    updated = _line_by_id(list(response.json()["order_data"] or []), line_id)
    assert int(updated["qty"]) == 3
    assert float(updated["unit_price"]) == pytest.approx(1100)
    assert updated.get("price_source") == "oneoff"
    assert float(updated["line_total"]) == pytest.approx(3300)


def test_catalog_priced_line_rejects_oneoff(
    piles_client_db: tuple[TestClient, Path],
) -> None:
    client, pb_db = piles_client_db
    count_before = _pile_prices_count(pb_db)
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "piles", "text": "С120.35-12 B25 2"},
    )
    assert create.status_code == 200, create.text
    line = list(create.json()["order_data"] or [])[0]
    assert line.get("unit_price") is not None
    line_id = str(line["line_id"])
    draft_id = str(create.json()["draft_id"])

    response = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{line_id}",
        json={"unit_price": 1},
    )
    assert response.status_code == 400, response.text
    assert "Цена из прайса" in str(response.json().get("detail") or "")
    after = client.get(f"/api/v1/commercial/drafts/{draft_id}").json()
    unchanged = _line_by_id(list(after["order_data"] or []), line_id)
    assert float(unchanged["unit_price"]) == pytest.approx(float(line["unit_price"]))
    assert unchanged.get("price_source") != "oneoff"
    assert _pile_prices_count(pb_db) == count_before


def test_oneoff_with_source_text_rejected(
    piles_client_db: tuple[TestClient, Path],
) -> None:
    client, _pb_db = piles_client_db
    draft_id, line = _create_unpriced_pile_draft(client)
    line_id = str(line["line_id"])
    response = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{line_id}",
        json={"source_text": "С120.35-12 B25 2", "unit_price": 1000},
    )
    assert response.status_code == 400, response.text
    after = client.get(f"/api/v1/commercial/drafts/{draft_id}").json()
    unchanged = _line_by_id(list(after["order_data"] or []), line_id)
    assert unchanged.get("unit_price") is None


def test_oneoff_on_sealed_line_rejected(mixed_client: TestClient) -> None:
    create = mixed_client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "piles", "text": "С110.30-6 B25 2"},
    )
    assert create.status_code == 200, create.text
    draft_id = str(create.json()["draft_id"])
    pile_line = list(create.json()["order_data"] or [])[0]
    line_id = str(pile_line["line_id"])

    start = mixed_client.post(
        f"/api/v1/commercial/drafts/{draft_id}/append/start",
        json={"product_type": "plates"},
    )
    assert start.status_code == 200, start.text
    sealed = _line_by_id(list(start.json()["order_data"] or []), line_id)
    assert str(sealed.get("append_batch_id") or "").strip()

    response = mixed_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{line_id}",
        json={"unit_price": 900},
    )
    assert response.status_code == 400, response.text
    after = mixed_client.get(f"/api/v1/commercial/drafts/{draft_id}").json()
    unchanged = _line_by_id(list(after["order_data"] or []), line_id)
    assert unchanged.get("unit_price") is None


def test_oneoff_zero_and_over_cap_rejected(
    piles_client_db: tuple[TestClient, Path],
) -> None:
    client, _pb_db = piles_client_db
    draft_id, line = _create_unpriced_pile_draft(client)
    line_id = str(line["line_id"])

    zero = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{line_id}",
        json={"unit_price": 0},
    )
    assert zero.status_code == 400, zero.text

    huge = client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{line_id}",
        json={"unit_price": 10_000_001},
    )
    assert huge.status_code == 400, huge.text
    after = client.get(f"/api/v1/commercial/drafts/{draft_id}").json()
    assert _line_by_id(list(after["order_data"] or []), line_id).get("unit_price") is None


def test_clear_null_on_catalog_line_rejected(
    piles_client_db: tuple[TestClient, Path],
) -> None:
    client, _pb_db = piles_client_db
    create = client.post(
        "/api/v1/commercial/drafts",
        data={"product_type": "piles", "text": "С120.35-12 B25 2"},
    )
    assert create.status_code == 200, create.text
    line = list(create.json()["order_data"] or [])[0]
    response = client.patch(
        f"/api/v1/commercial/drafts/{create.json()['draft_id']}/lines/{line['line_id']}",
        json={"unit_price": None},
    )
    assert response.status_code == 400, response.text
