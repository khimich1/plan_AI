"""PATCH /lines/{line_id} qty + source_text replace, and POST /lines/restore."""

from __future__ import annotations

import json
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


def test_patch_line_qty_invalidates_breakdown_tables(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """After line edit, cached breakdown must be cleared (not served stale)."""
    from app.services.draft_store import DraftStore

    draft_id, plates = _create_plates_result_draft(mixed_client, monkeypatch)
    target_id = str(plates[0]["line_id"])
    old_qty = int(plates[0]["qty"])

    before = mixed_client.get(f"/api/v1/commercial/drafts/{draft_id}")
    assert before.status_code == 200
    assert int(before.json()["metadata"].get("breakdown_tables_count") or 0) > 0

    stored_before = DraftStore().load_raw_json(draft_id) or {}
    tables_before = list((stored_before.get("metadata") or {}).get("breakdown_tables") or [])
    assert len(tables_before) > 0

    patch = mixed_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{target_id}",
        json={"qty": old_qty + 3},
    )
    assert patch.status_code == 200, patch.text
    assert int(patch.json()["metadata"].get("breakdown_tables_count") or 0) == 0

    stored_after = DraftStore().load_raw_json(draft_id) or {}
    meta_after = dict(stored_after.get("metadata") or {})
    assert list(meta_after.get("breakdown_tables") or []) == []
    assert int(meta_after.get("breakdown_tables_count") or 0) == 0


def test_get_breakdown_regenerates_after_qty_patch(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """GET breakdown after invalidate rebuilds tables from current composition (not stale cache)."""
    draft_id, plates = _create_plates_result_draft(mixed_client, monkeypatch)
    target_id = str(plates[0]["line_id"])
    old_qty = int(plates[0]["qty"])

    patch = mixed_client.patch(
        f"/api/v1/commercial/drafts/{draft_id}/lines/{target_id}",
        json={"qty": old_qty + 7},
    )
    assert patch.status_code == 200, patch.text
    assert int(patch.json()["metadata"].get("breakdown_tables_count") or 0) == 0

    response = mixed_client.get(f"/api/v1/commercial/drafts/{draft_id}/breakdown")
    assert response.status_code == 200, response.text
    items = list(response.json().get("items") or [])
    assert len(items) > 0

    details = mixed_client.get(f"/api/v1/commercial/drafts/{draft_id}")
    assert details.status_code == 200
    assert int(details.json()["metadata"].get("breakdown_tables_count") or 0) == len(items)


def test_delete_line_invalidates_breakdown_tables(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from app.services.draft_store import DraftStore

    draft_id, plates = _create_plates_result_draft(mixed_client, monkeypatch)
    target_id = str(plates[0]["line_id"])
    assert int(
        mixed_client.get(f"/api/v1/commercial/drafts/{draft_id}").json()["metadata"].get(
            "breakdown_tables_count"
        )
        or 0
    ) > 0

    delete = mixed_client.delete(f"/api/v1/commercial/drafts/{draft_id}/lines/{target_id}")
    assert delete.status_code == 200, delete.text
    assert int(delete.json()["metadata"].get("breakdown_tables_count") or 0) == 0

    stored = DraftStore().load_raw_json(draft_id) or {}
    meta = dict(stored.get("metadata") or {})
    assert list(meta.get("breakdown_tables") or []) == []


def test_undo_last_append_invalidates_breakdown_tables(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from app.services.draft_store import DraftStore

    draft_id, body, _plate_lines, _pile_lines = _build_plates_then_piles_draft(
        mixed_client, monkeypatch
    )
    # Ensure tables exist (plates calculate may leave them; force via GET if empty).
    if int(body["metadata"].get("breakdown_tables_count") or 0) == 0:
        regen = mixed_client.get(f"/api/v1/commercial/drafts/{draft_id}/breakdown")
        assert regen.status_code == 200, regen.text
        assert len(list(regen.json().get("items") or [])) > 0
        body = mixed_client.get(f"/api/v1/commercial/drafts/{draft_id}").json()
    assert int(body["metadata"].get("breakdown_tables_count") or 0) > 0

    undo = mixed_client.post(f"/api/v1/commercial/drafts/{draft_id}/append/undo-last")
    assert undo.status_code == 200, undo.text
    assert int(undo.json()["metadata"].get("breakdown_tables_count") or 0) == 0

    stored = DraftStore().load_raw_json(draft_id) or {}
    assert list((stored.get("metadata") or {}).get("breakdown_tables") or []) == []


def test_restore_line_invalidates_breakdown_tables(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from app.services.draft_store import DraftStore

    draft_id, plates = _create_plates_result_draft(mixed_client, monkeypatch)
    snapshot = dict(plates[0])
    target_id = str(snapshot["line_id"])

    delete = mixed_client.delete(f"/api/v1/commercial/drafts/{draft_id}/lines/{target_id}")
    assert delete.status_code == 200, delete.text

    regen = mixed_client.get(f"/api/v1/commercial/drafts/{draft_id}/breakdown")
    assert regen.status_code == 200, regen.text
    assert len(list(regen.json().get("items") or [])) > 0
    details = mixed_client.get(f"/api/v1/commercial/drafts/{draft_id}").json()
    assert int(details["metadata"].get("breakdown_tables_count") or 0) > 0

    restore = mixed_client.post(
        f"/api/v1/commercial/drafts/{draft_id}/lines/restore",
        json={"index": 0, "lines": [snapshot], "replace_line_ids": []},
    )
    assert restore.status_code == 200, restore.text
    assert int(restore.json()["metadata"].get("breakdown_tables_count") or 0) == 0

    stored = DraftStore().load_raw_json(draft_id) or {}
    assert list((stored.get("metadata") or {}).get("breakdown_tables") or []) == []


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


_OLD_2280 = "Плиты ПБ 22,80-12-8п"
_NEW_228 = "Плиты ПБ 22,8-12-8п"
_OLD_WIDTH = "Плиты ПБ 29-5,25-8п"
_NEW_WIDTH = "Плиты ПБ 29-5,3-8п"


def _draft_file(draft_id: str) -> Path:
    from app.core.settings import get_settings

    return Path(get_settings().drafts_dir) / f"{draft_id}.json"


def _retarget_plates(draft_id: str, name: str, *, oneoff: bool = False) -> list[int]:
    path = _draft_file(draft_id)
    data = json.loads(path.read_text(encoding="utf-8"))
    qtys: list[int] = []
    for index, line in enumerate(data["order_data"]):
        if str(line.get("product_type") or "plates") != "plates":
            continue
        qty = 2 if index == 0 else 4
        line["name"] = name
        line["source_text"] = f"{name} старый текст"
        line["qty"] = qty
        line["concrete_grade"] = "М500"
        if oneoff:
            line["price_source"] = "oneoff"
        qtys.append(qty)
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return qtys


def _seed_one_plate_card(name: str, guid: str, *, duplicate: bool = False) -> None:
    import sqlite3

    from app.core.settings import get_settings
    from core.config_and_data import parse_name_to_sizes
    from core.kp_db_nomenclature import load_class_from_plate_name
    from core.price_db import length_m_to_price_length_dm

    pb = str(get_settings().pb_db_path)
    length_m, _width = parse_name_to_sizes(name)
    load_class = load_class_from_plate_name(name)
    assert length_m is not None and load_class is not None
    with sqlite3.connect(pb) as conn:
        conn.execute(
            "INSERT OR REPLACE INTO prices (length_dm, load_code, price) VALUES (?, ?, 10000)",
            (length_m_to_price_length_dm(length_m), load_class // 100),
        )
        conn.execute(
            'INSERT INTO prays_plity ("Уникальный идентификатор (Номенклатура)", "Товар") '
            "VALUES (?, ?)",
            (guid, name),
        )
        if duplicate:
            conn.execute(
                'INSERT INTO prays_plity ("Уникальный идентификатор (Номенклатура)", "Товар") '
                "VALUES (?, ?)",
                (guid + "-b", name),
            )


def test_manager_replaces_every_line_and_does_not_touch_the_queue(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """NG-301: 22,80 → 22,8 во всех строках, скидка черновика, очередь молчит."""
    import app.services.guid_queue_sync as queue_sync
    from app.services.commercial_workflow_service import CommercialWorkflowService
    from app.services.promise_service import PromiseService

    def fail_queue(*_args: object, **_kwargs: object) -> None:
        raise AssertionError("очередь")

    def fail_promise(*_args: object, **_kwargs: object) -> None:
        raise AssertionError("обещание")

    monkeypatch.setattr(queue_sync, "sync_guid_queue_with_kp", fail_queue)
    monkeypatch.setattr(PromiseService, "recalc_on_composition_change", fail_promise)
    draft_id, _plates = _create_plates_result_draft(mixed_client, monkeypatch)
    qtys = _retarget_plates(draft_id, _OLD_2280)
    _seed_one_plate_card(_NEW_228, "guid-228")
    before = json.loads(_draft_file(draft_id).read_text(encoding="utf-8"))
    assert int(before["metadata"].get("breakdown_tables_count") or 0) > 0

    body = CommercialWorkflowService().replace_guid_mark(
        draft_id,
        product_kind="plate",
        mark=_OLD_2280,
        target_mark=_NEW_228,
    )

    lines = [line for line in body["order_data"] if line.get("name") == _NEW_228]
    assert len(lines) == len(qtys)
    assert [int(line["qty"]) for line in lines] == qtys
    assert all(line["source_text"] == _NEW_228 for line in lines)
    assert all(line["nomenclature_id"] == "guid-228" for line in lines)
    assert all(int(line["load_class"]) == 800 for line in lines)
    assert all(float(line["unit_price"]) == pytest.approx(10000) for line in lines)
    assert all(float(line["discounted_price"]) == pytest.approx(9400) for line in lines)
    assert body["metadata"].get("breakdown_tables_count") == 0
    from core.guid_gate import OrderLine, check_invoice_guids
    import sqlite3
    from app.core.settings import get_settings

    conn = sqlite3.connect(str(get_settings().pb_db_path))
    try:
        report = check_invoice_guids([OrderLine("plate", _NEW_228)], conn)
    finally:
        conn.close()
    assert not report.missing


def test_manager_replace_rewrites_width_weight_and_refuses_bad_targets(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from app.services.commercial_workflow_service import CommercialWorkflowService
    from core.guid_replace import GuidReplaceError
    from core.kp_plate_weight import resolve_kp_line_weight_kg

    draft_id, _plates = _create_plates_result_draft(mixed_client, monkeypatch)
    _retarget_plates(draft_id, _OLD_WIDTH)
    _seed_one_plate_card(_NEW_WIDTH, "guid-53")
    body = CommercialWorkflowService().replace_guid_mark(
        draft_id,
        product_kind="plate",
        mark=_OLD_WIDTH,
        target_mark=_NEW_WIDTH,
    )
    line = next(item for item in body["order_data"] if item["name"] == _NEW_WIDTH)
    _unit, weight = resolve_kp_line_weight_kg(
        {"length_m": line["length_m"], "width_m": 0.53, "qty": line["qty"]}
    )
    assert float(line["width_m"]) == pytest.approx(0.53)
    assert float(line["weight"]) == pytest.approx(weight)
    assert line["nomenclature_id"] == "guid-53"

    blocked = _create_plates_result_draft(mixed_client, monkeypatch)[0]
    _retarget_plates(blocked, _OLD_2280)
    _seed_one_plate_card(_NEW_228, "guid-two", duplicate=True)
    with pytest.raises(GuidReplaceError, match="одним номером"):
        CommercialWorkflowService().replace_guid_mark(
            blocked,
            product_kind="plate",
            mark=_OLD_2280,
            target_mark=_NEW_228,
        )

    oneoff = _create_plates_result_draft(mixed_client, monkeypatch)[0]
    _retarget_plates(oneoff, _OLD_2280, oneoff=True)
    with pytest.raises(GuidReplaceError, match="(?i)договорн"):
        CommercialWorkflowService().replace_guid_mark(
            oneoff,
            product_kind="plate",
            mark=_OLD_2280,
            target_mark=_NEW_228,
        )
    stored = json.loads(_draft_file(oneoff).read_text(encoding="utf-8"))
    assert stored["order_data"][0]["name"] == _OLD_2280


def test_draft_guid_replace_api_previews_without_writing_and_hides_foreign_drafts(
    mixed_client: TestClient,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from app.security.session import create_session_token
    from tests.helpers.auth_fixtures import patch_auth_users

    draft_id, _plates = _create_plates_result_draft(mixed_client, monkeypatch)
    _retarget_plates(draft_id, _OLD_2280)
    _seed_one_plate_card(_NEW_228, "guid-api")
    body = {
        "product_kind": "plate",
        "mark": _OLD_2280,
        "target_mark": _NEW_228,
    }
    preview = mixed_client.post(
        f"/api/v1/commercial/drafts/{draft_id}/guid-replace/preview",
        json=body,
    )
    assert preview.status_code == 200, preview.text
    item = preview.json()["items"][0]
    assert item["qty"] == 6
    assert item["will"]["unit_discounted"] == pytest.approx(9400)
    stored = json.loads(_draft_file(draft_id).read_text(encoding="utf-8"))
    assert stored["order_data"][0]["name"] == _OLD_2280

    written = mixed_client.post(
        f"/api/v1/commercial/drafts/{draft_id}/guid-replace",
        json=body,
    )
    assert written.status_code == 200, written.text
    assert all(
        line["name"] == _NEW_228
        for line in written.json()["order_data"]
        if str(line.get("product_type") or "plates") == "plates"
    )

    denied_id, _ = _create_plates_result_draft(mixed_client, monkeypatch)
    _retarget_plates(denied_id, _OLD_2280)
    _seed_one_plate_card(_NEW_228, "guid-dup", duplicate=True)
    denied = mixed_client.post(
        f"/api/v1/commercial/drafts/{denied_id}/guid-replace",
        json=body,
    )
    assert denied.status_code == 400
    assert "одним номером" in denied.text

    patch_auth_users(
        monkeypatch,
        [
            {
                "id": 2,
                "username": "other",
                "role": "manager",
                "manager_id": None,
                "is_active": 1,
                "created_at": "2026-01-01 00:00:00",
            }
        ],
    )
    token = create_session_token(
        {"id": 2, "username": "other", "role": "manager"}, ttl_seconds=300
    )
    mixed_client.cookies.set("app_session", token)
    foreign = mixed_client.post(
        f"/api/v1/commercial/drafts/{draft_id}/guid-replace/preview",
        json=body,
    )
    assert foreign.status_code == 403

