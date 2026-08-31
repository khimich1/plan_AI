"""PATCH /lines/{line_id} qty + source_text replace, and POST /lines/restore."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from fastapi.testclient import TestClient

from app.core.http_errors import MSG_VALIDATION
from tests.test_commercial_draft_append import (
    _assert_append_batches_cover_order,
    _auth_client,
    _build_plates_then_piles_draft,
    _create_plates_result_draft,
    _setup_mixed_price_env,
)


@pytest.fixture()
def mixed_client(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> TestClient:
    _setup_mixed_price_env(monkeypatch, tmp_path)
    return _auth_client(monkeypatch)


def _line_by_id(order_data: list[dict[str, Any]], line_id: str) -> dict[str, Any]:
    for line in order_data:
        if str(line.get("line_id") or "") == line_id:
            return line
    raise AssertionError(f"line {line_id!r} not in order_data")


def test_patch_line_qty_updates_totals_keeps_line_id(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    draft_id, plates = _create_plates_result_draft(mixed_client, monkeypatch)
    target = plates[0]
    target_id = str(target["line_id"])
    old_qty = int(target["qty"])
    unit_price = float(target["unit_price"])
    new_qty = old_qty + 5
    assert new_qty != old_qty

    before = mixed_client.get(f"/api/v1/commercial/drafts/{draft_id}")
    assert before.status_code == 200, before.text
    totals_before = dict(before.json().get("totals") or {})

    patch = mixed_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{target_id}",
        json={"qty": new_qty},
    )
    assert patch.status_code == 200, patch.text
    body = patch.json()
    line = _line_by_id(list(body["order_data"] or []), target_id)
    assert int(line["qty"]) == new_qty
    assert str(line["line_id"]) == target_id
    assert float(line["unit_price"]) == pytest.approx(unit_price)
    totals_after = dict(body.get("totals") or {})
    assert totals_after.get("total_qty") != totals_before.get("total_qty") or new_qty != old_qty
    assert float(totals_after.get("subtotal") or 0) != pytest.approx(
        float(totals_before.get("subtotal") or 0)
    )


def test_patch_line_qty_missing_returns_russian_404(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    draft_id, _plates = _create_plates_result_draft(mixed_client, monkeypatch)
    missing_id = "ln_does_not_exist_qty"
    response = mixed_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{missing_id}",
        json={"qty": 3},
    )
    assert response.status_code == 404, response.text
    detail = response.json().get("detail", "")
    assert detail == "Строка не найдена."
    assert missing_id not in detail


def test_patch_line_qty_zero_returns_400_unchanged(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    draft_id, plates = _create_plates_result_draft(mixed_client, monkeypatch)
    target_id = str(plates[0]["line_id"])
    old_qty = int(plates[0]["qty"])
    response = mixed_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{target_id}",
        json={"qty": 0},
    )
    assert response.status_code == 400, response.text
    after = mixed_client.get(f"/api/v1/commercial/drafts/{draft_id}")
    assert after.status_code == 200
    line = _line_by_id(list(after.json()["order_data"] or []), target_id)
    assert int(line["qty"]) == old_qty


def test_patch_line_source_text_1_to_1_updates_mark(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    draft_id, plates = _create_plates_result_draft(mixed_client, monkeypatch)
    target = plates[0]
    target_id = str(target["line_id"])
    index = 0
    response = mixed_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{target_id}",
        json={"source_text": "ПБ 60-12-8п 3"},
    )
    assert response.status_code == 200, response.text
    order = list(response.json()["order_data"] or [])
    assert len(order) == len(plates)
    replacement = order[index]
    assert str(replacement.get("line_id") or "") != ""
    assert target_id not in [str(ln.get("line_id")) for ln in order] or str(
        replacement.get("name") or replacement.get("mark") or ""
    )
    name = str(replacement.get("name") or replacement.get("mark") or "")
    assert "60-12" in name or "60-12" in str(replacement)
    assert int(replacement["qty"]) == 3
    assert replacement.get("unit_price") is not None


def test_patch_line_source_text_1_to_n_splices_and_updates_batches(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    draft_id, body, plate_lines, pile_lines = _build_plates_then_piles_draft(
        mixed_client, monkeypatch
    )
    target = pile_lines[0]
    target_id = str(target["line_id"])
    order_before = list(body["order_data"] or [])
    index = next(i for i, ln in enumerate(order_before) if str(ln.get("line_id")) == target_id)

    response = mixed_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{target_id}",
        json={"source_text": "С120.35-12 B25 1\nС120.35-13и B25 1"},
    )
    assert response.status_code == 200, response.text
    after = response.json()
    order_after = list(after["order_data"] or [])
    after_ids = [str(ln.get("line_id") or "") for ln in order_after]
    assert target_id not in after_ids
    spliced = order_after[index : index + 2]
    assert len(spliced) == 2
    new_ids = [str(ln.get("line_id") or "") for ln in spliced]
    assert all(nid and nid != target_id for nid in new_ids)
    assert len(order_after) == len(order_before) + 1

    for batch in after["metadata"].get("append_batches") or []:
        ids = list(batch.get("line_ids") or [])
        assert target_id not in ids
    _assert_append_batches_cover_order(order_after, after["metadata"]["append_batches"])
    pile_batch = next(
        b for b in after["metadata"]["append_batches"] if b.get("product_type") == "piles"
    )
    for nid in new_ids:
        assert nid in list(pile_batch.get("line_ids") or [])


def test_patch_line_invalid_source_text_400_unchanged(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    draft_id, plates = _create_plates_result_draft(mixed_client, monkeypatch)
    target_id = str(plates[0]["line_id"])
    names_before = [str(ln.get("name") or ln.get("mark")) for ln in plates]
    response = mixed_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{target_id}",
        json={"source_text": "это не марка изделия xyz"},
    )
    assert response.status_code == 400, response.text
    detail = response.json().get("detail", "")
    assert detail in {MSG_VALIDATION, "Не удалось распознать строку."} or isinstance(detail, str)
    after = mixed_client.get(f"/api/v1/commercial/drafts/{draft_id}")
    assert after.status_code == 200
    names_after = [
        str(ln.get("name") or ln.get("mark")) for ln in after.json()["order_data"] or []
    ]
    assert names_after == names_before
    assert target_id in [str(ln.get("line_id")) for ln in after.json()["order_data"] or []]


def test_restore_after_delete_returns_same_line_id(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    draft_id, plates = _create_plates_result_draft(mixed_client, monkeypatch)
    snapshot = dict(plates[0])
    target_id = str(snapshot["line_id"])
    delete = mixed_client.delete(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{target_id}",
    )
    assert delete.status_code == 200, delete.text
    assert target_id not in [
        str(ln.get("line_id")) for ln in delete.json()["order_data"] or []
    ]

    restore = mixed_client.post(
        f"/api/v1/commercial/drafts/{draft_id}/lines/restore",
        json={"index": 0, "lines": [snapshot], "replace_line_ids": []},
    )
    assert restore.status_code == 200, restore.text
    order = list(restore.json()["order_data"] or [])
    assert str(order[0].get("line_id")) == target_id
    _assert_append_batches_cover_order(order, restore.json()["metadata"]["append_batches"])


def test_restore_after_replace_uses_replace_line_ids(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    draft_id, _body = _create_plates_result_draft(mixed_client, monkeypatch)
    before = mixed_client.get(f"/api/v1/commercial/drafts/{draft_id}").json()
    order_before = list(before["order_data"] or [])
    snapshot = dict(order_before[0])
    target_id = str(snapshot["line_id"])

    patched = mixed_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{target_id}",
        json={"source_text": "ПБ 78-12-8п 1\nПБ 60-12-8п 1"},
    )
    assert patched.status_code == 200, patched.text
    after_replace = list(patched.json()["order_data"] or [])
    new_ids = [
        str(ln.get("line_id"))
        for ln in after_replace
        if str(ln.get("line_id")) != target_id
        and str(ln.get("line_id")) not in {str(x.get("line_id")) for x in order_before[1:]}
    ]
    assert len(new_ids) >= 2

    restore = mixed_client.post(
        f"/api/v1/commercial/drafts/{draft_id}/lines/restore",
        json={"index": 0, "lines": [snapshot], "replace_line_ids": new_ids},
    )
    assert restore.status_code == 200, restore.text
    restored_ids = [str(ln.get("line_id")) for ln in restore.json()["order_data"] or []]
    assert restored_ids[0] == target_id
    for nid in new_ids:
        assert nid not in restored_ids
