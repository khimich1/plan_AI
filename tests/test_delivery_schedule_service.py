"""T4/T5: CRUD DeliveryScheduleService + живой светофор на get/replace."""

from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any

import pytest
from fastapi import HTTPException
from pydantic import ValidationError

from app.domain.enums import KpStatus
from app.schemas.delivery_schedule import (
    BatchIn,
    BatchItemIn,
    DeliverySchedulePut,
)
from app.security.offer_access import FORBIDDEN_OFFER_DETAIL
from app.services.delivery_schedule_service import (
    DeliveryScheduleNotFoundError,
    DeliveryScheduleService,
    DeliveryScheduleValidationError,
)
from core import kp_db_schema
from core.kp_db_common import _connect

ADMIN = {"id": 1, "role": "admin"}
MANAGER = {"id": 10, "role": "manager"}
MANAGER_B = {"id": 20, "role": "manager"}

_TRAFFIC_STATUSES = frozenset({"green", "yellow", "red"})
_TODAY = "2026-08-07"


def _fresh_db(tmp_path: Path, name: str = "plita.db") -> str:
    db_path = str(tmp_path / name)
    kp_db_schema._schema_ready.clear()
    kp_db_schema.ensure_schema(db_path)
    return db_path


def _seed_kp(
    db_path: str,
    *,
    kp_id: int = 1,
    status: str = KpStatus.IN_WORK.value,
    owner_user_id: int | None = 1,
    plate_qty: int = 10,
    plate_name: str = "ПБ 60-12-8п",
    length_m: float | None = None,
    width_m: float | None = None,
    load_class: int | None = None,
) -> int:
    """КП + meta + одна позиция. Возвращает plate_id."""
    with _connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO KP_offers (kp_id, creation_date, customer_name) "
            "VALUES (?, '2026-08-07', 'ООО Тест')",
            (kp_id,),
        )
        cur.execute(
            "INSERT INTO kp_meta (kp_id, status, owner_user_id, product_type) "
            "VALUES (?, ?, ?, 'plates')",
            (kp_id, status, owner_user_id),
        )
        cur.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, qty,
                length_m, width_m, load_class
            )
            VALUES (?, 1, ?, ?, ?, ?, ?)
            """,
            (kp_id, plate_name, plate_qty, length_m, width_m, load_class),
        )
        plate_id = int(cur.lastrowid)
        conn.commit()
    return plate_id


def _seed_completed_plates(
    db_path: str,
    *,
    kp_id: int,
    plate_name: str,
    qty: int,
    length_m: float | None = 6.0,
    width_m: float | None = 1.2,
    load_class: int | None = 800,
) -> None:
    """Минимальный seed СГП для produced через KpReadinessService."""
    with _connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO completed_plates (
                kp_id, plate_name, length_m, width_m, load_class,
                qty, completed_date, production_day, plan_id
            ) VALUES (?, ?, ?, ?, ?, ?, '07.08.2026', 1, 'plan-t5')
            """,
            (kp_id, plate_name, length_m, width_m, load_class, qty),
        )
        conn.commit()


def _patch_empty_occupancy(monkeypatch: pytest.MonkeyPatch) -> None:
    """days_info={} → все даты свободны (R2-дефолт в check_batches)."""
    monkeypatch.setattr(
        "app.services.delivery_schedule_service.get_global_calendar_info",
        lambda: {"days_info": {}},
    )


def _batch(
    *,
    plate_id: int,
    qty: int = 3,
    name: str = "1 этаж",
    deliver_from: str = "2026-09-01",
    deliver_to: str = "2026-09-10",
    produce_by: str = "2026-08-25",
    sort_order: int = 0,
) -> BatchIn:
    return BatchIn(
        name=name,
        deliver_from=deliver_from,
        deliver_to=deliver_to,
        produce_by=produce_by,
        items=[BatchItemIn(plate_id=plate_id, qty=qty)],
        sort_order=sort_order,
    )


def _payload(
    plate_id: int,
    *,
    qty: int = 3,
    invoice_number: str | None = "СЧ-101",
    contract_number: str | None = "Д-5",
    batches: list[BatchIn] | None = None,
) -> DeliverySchedulePut:
    return DeliverySchedulePut(
        invoice_number=invoice_number,
        contract_number=contract_number,
        batches=batches if batches is not None else [_batch(plate_id=plate_id, qty=qty)],
    )


def _batch_content(view: Any) -> list[dict[str, Any]]:
    """Логическое содержимое партий (без id — replace пересоздаёт строки)."""
    result: list[dict[str, Any]] = []
    for batch in view.batches:
        result.append(
            {
                "name": batch.name,
                "deliver_from": batch.deliver_from,
                "deliver_to": batch.deliver_to,
                "produce_by": batch.produce_by,
                "sort_order": batch.sort_order,
                "items": [
                    {"plate_id": item.plate_id, "qty": item.qty}
                    for item in batch.items
                ],
            }
        )
    return result


def test_get_raises_not_found_when_no_schedule(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    _seed_kp(db_path, kp_id=1)
    service = DeliveryScheduleService(db_path=db_path)

    with pytest.raises(DeliveryScheduleNotFoundError, match="не найден"):
        service.get(1, user=ADMIN)


def test_get_raises_not_found_when_kp_missing(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    service = DeliveryScheduleService(db_path=db_path)

    with pytest.raises(DeliveryScheduleNotFoundError, match="не найдено"):
        service.get(999, user=ADMIN)


def test_get_foreign_manager_forbidden(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    plate_id = _seed_kp(db_path, kp_id=1, owner_user_id=MANAGER["id"])
    service = DeliveryScheduleService(db_path=db_path)
    service.replace(1, _payload(plate_id, qty=1), MANAGER)

    with pytest.raises(HTTPException) as exc_info:
        service.get(1, user=MANAGER_B)
    assert exc_info.value.status_code == 403
    assert exc_info.value.detail == FORBIDDEN_OFFER_DETAIL


def test_import_draft_foreign_manager_forbidden(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    _seed_kp(db_path, kp_id=1, owner_user_id=MANAGER["id"])
    service = DeliveryScheduleService(db_path=db_path)

    with pytest.raises(HTTPException) as exc_info:
        service.import_draft(1, b"PK\x03\x04", user=MANAGER_B)
    assert exc_info.value.status_code == 403
    assert exc_info.value.detail == FORBIDDEN_OFFER_DETAIL


def test_replace_creates_schedule_and_is_idempotent_by_content(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    plate_id = _seed_kp(db_path, kp_id=1, plate_qty=10)
    service = DeliveryScheduleService(db_path=db_path)
    payload = _payload(plate_id, qty=4, invoice_number="СЧ-200")

    first = service.replace(1, payload, ADMIN)
    assert first.kp_id == 1
    assert first.status == "draft"
    assert first.invoice_number == "СЧ-200"
    assert first.contract_number == "Д-5"
    assert len(first.batches) == 1
    assert first.batches[0].name == "1 этаж"
    assert first.batches[0].items[0].plate_id == plate_id
    assert first.batches[0].items[0].qty == 4
    assert first.batches[0].items[0].plate_name == "ПБ 60-12-8п"

    second = service.replace(1, payload, ADMIN)
    assert second.kp_id == first.kp_id
    assert second.id == first.id  # upsert header — тот же schedule_id
    assert second.invoice_number == first.invoice_number
    assert second.contract_number == first.contract_number
    assert second.status == first.status
    assert _batch_content(second) == _batch_content(first)

    with _connect(db_path) as conn:
        n_schedules = conn.execute(
            "SELECT COUNT(*) FROM delivery_schedule WHERE kp_id = 1"
        ).fetchone()[0]
        n_batches = conn.execute("SELECT COUNT(*) FROM delivery_batch").fetchone()[0]
    assert n_schedules == 1
    assert n_batches == 1


def test_replace_rejects_qty_sum_exceeding_kp_plate(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    plate_id = _seed_kp(db_path, kp_id=1, plate_qty=5)
    service = DeliveryScheduleService(db_path=db_path)
    payload = DeliverySchedulePut(
        invoice_number="СЧ-1",
        batches=[
            _batch(plate_id=plate_id, qty=3, name="A", sort_order=0),
            _batch(plate_id=plate_id, qty=3, name="B", sort_order=1),
        ],
    )

    with pytest.raises(DeliveryScheduleValidationError, match="превышает"):
        service.replace(1, payload, ADMIN)


def test_batch_rejects_deliver_from_after_deliver_to() -> None:
    with pytest.raises(ValidationError):
        BatchIn(
            name="партия",
            deliver_from="2026-09-15",
            deliver_to="2026-09-01",
            produce_by="2026-08-25",
            items=[BatchItemIn(plate_id=1, qty=1)],
        )


@pytest.mark.parametrize("bad_qty", [0, -1])
def test_batch_item_rejects_qty_less_than_one(bad_qty: int) -> None:
    with pytest.raises(ValidationError):
        BatchItemIn(plate_id=1, qty=bad_qty)


@pytest.mark.parametrize(
    "status",
    [KpStatus.ARCHIVED.value, KpStatus.DONE.value],
)
def test_replace_rejects_kp_status_not_editable(tmp_path: Path, status: str) -> None:
    db_path = _fresh_db(tmp_path, f"status_{status}.db")
    plate_id = _seed_kp(db_path, kp_id=1, status=status)
    service = DeliveryScheduleService(db_path=db_path)

    with pytest.raises(DeliveryScheduleValidationError, match="в работе"):
        service.replace(1, _payload(plate_id), ADMIN)


def test_replace_allows_on_sgp_status(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    plate_id = _seed_kp(db_path, kp_id=1, status=KpStatus.ON_SGP.value)
    service = DeliveryScheduleService(db_path=db_path)

    view = service.replace(1, _payload(plate_id, qty=2), ADMIN)
    assert view.kp_id == 1
    assert len(view.batches) == 1


def test_replace_updates_invoice_number(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    plate_id = _seed_kp(db_path, kp_id=1)
    service = DeliveryScheduleService(db_path=db_path)

    first = service.replace(
        1, _payload(plate_id, invoice_number="СЧ-OLD", qty=2), ADMIN
    )
    assert first.invoice_number == "СЧ-OLD"

    updated = service.replace(
        1, _payload(plate_id, invoice_number="СЧ-NEW", qty=2), ADMIN
    )
    assert updated.invoice_number == "СЧ-NEW"
    assert updated.id == first.id

    loaded = service.get(1, user=ADMIN)
    assert loaded.invoice_number == "СЧ-NEW"


def test_replace_rejects_foreign_plate_id(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    plate_a = _seed_kp(db_path, kp_id=1, plate_qty=10)
    plate_b = _seed_kp(db_path, kp_id=2, plate_qty=10, plate_name="ПБ 72-12-8п")
    service = DeliveryScheduleService(db_path=db_path)

    # КП №1, но plate_id принадлежит КП №2
    payload = _payload(plate_b, qty=1)
    with pytest.raises(DeliveryScheduleValidationError, match="не принадлежит"):
        service.replace(1, payload, ADMIN)

    # sanity: свой plate_id проходит
    ok = service.replace(1, _payload(plate_a, qty=1), ADMIN)
    assert ok.batches[0].items[0].plate_id == plate_a


def test_replace_not_found_for_missing_kp(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    service = DeliveryScheduleService(db_path=db_path)
    payload = DeliverySchedulePut(
        batches=[
            BatchIn(
                name="x",
                deliver_from="2026-09-01",
                deliver_to="2026-09-10",
                produce_by="2026-08-25",
                items=[],
            )
        ]
    )

    with pytest.raises(DeliveryScheduleNotFoundError, match="не найдено"):
        service.replace(999, payload, ADMIN)


def test_manager_can_replace_own_kp(tmp_path: Path) -> None:
    db_path = _fresh_db(tmp_path)
    plate_id = _seed_kp(db_path, kp_id=1, owner_user_id=MANAGER["id"])
    service = DeliveryScheduleService(db_path=db_path)

    view = service.replace(1, _payload(plate_id, qty=1), MANAGER)
    assert view.kp_id == 1
    assert view.batches[0].items[0].qty == 1


# ---------------------------------------------------------------------------
# T5: живой светофор (enrich на get / replace)
# ---------------------------------------------------------------------------


def test_replace_and_get_set_traffic_status_green_when_capacity_free(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Пустая occupancy + далёкий produce_by → status green (и ∈ {g,y,r})."""
    _patch_empty_occupancy(monkeypatch)
    db_path = _fresh_db(tmp_path)
    plate_id = _seed_kp(
        db_path,
        kp_id=1,
        plate_qty=10,
        length_m=6.0,
        width_m=1.2,
        load_class=800,
    )
    service = DeliveryScheduleService(db_path=db_path)
    service._today_override = _TODAY

    payload = DeliverySchedulePut(
        invoice_number="СЧ-TL",
        batches=[
            _batch(
                plate_id=plate_id,
                qty=2,
                produce_by="2026-12-01",
                deliver_from="2026-12-05",
                deliver_to="2026-12-15",
            )
        ],
    )
    replaced = service.replace(1, payload, ADMIN)
    assert replaced.batches[0].status in _TRAFFIC_STATUSES
    assert replaced.batches[0].status == "green"

    loaded = service.get(1, user=ADMIN)
    assert loaded.batches[0].status in _TRAFFIC_STATUSES
    assert loaded.batches[0].status == "green"
    assert loaded.traffic_light_degraded is False


def test_get_marks_changed_when_kp_plate_qty_drops_below_batch(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """R4: qty позиции КП < qty в партии → get не падает, changed=True."""
    _patch_empty_occupancy(monkeypatch)
    db_path = _fresh_db(tmp_path)
    plate_id = _seed_kp(
        db_path,
        kp_id=1,
        plate_qty=10,
        length_m=6.0,
        width_m=1.2,
        load_class=800,
    )
    service = DeliveryScheduleService(db_path=db_path)
    service._today_override = _TODAY

    service.replace(1, _payload(plate_id, qty=5), ADMIN)

    with _connect(db_path) as conn:
        conn.execute("UPDATE kp_plates SET qty = 2 WHERE id = ?", (plate_id,))
        conn.commit()

    view = service.get(1, user=ADMIN)
    assert view.batches[0].changed is True
    assert view.batches[0].items[0].changed is True
    assert view.batches[0].status in _TRAFFIC_STATUSES


def test_get_green_when_completed_plates_cover_batch(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """produced закрывает партию → green, ready_date=None (tracks_needed=0)."""
    _patch_empty_occupancy(monkeypatch)
    plate_name = "ПБ 60-12-8п"
    db_path = _fresh_db(tmp_path)
    plate_id = _seed_kp(
        db_path,
        kp_id=1,
        plate_qty=10,
        plate_name=plate_name,
        length_m=6.0,
        width_m=1.2,
        load_class=800,
    )
    service = DeliveryScheduleService(db_path=db_path)
    service._today_override = _TODAY

    batch_qty = 3
    service.replace(1, _payload(plate_id, qty=batch_qty), ADMIN)
    _seed_completed_plates(
        db_path,
        kp_id=1,
        plate_name=plate_name,
        qty=batch_qty,
        length_m=6.0,
        width_m=1.2,
        load_class=800,
    )

    view = service.get(1, user=ADMIN)
    assert view.batches[0].status == "green"
    # Полностью закрытая партия: tracks_needed=0 → ready_date/hint пустые.
    assert view.batches[0].ready_date is None
    assert view.batches[0].hint is None


def test_produced_splits_on_sgp_across_same_identity_plates(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Две строки одной марки: on_sgp не дублируется на каждый plate_id."""
    _patch_empty_occupancy(monkeypatch)
    plate_name = "ПБ 60-12-8п"
    db_path = _fresh_db(tmp_path)
    plate_a = _seed_kp(
        db_path,
        kp_id=1,
        plate_qty=10,
        plate_name=plate_name,
        length_m=6.0,
        width_m=1.2,
        load_class=800,
    )
    with _connect(db_path) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO kp_plates (
                kp_id, position_number, plate_name, qty,
                length_m, width_m, load_class
            )
            VALUES (1, 2, ?, 10, 6.0, 1.2, 800)
            """,
            (plate_name,),
        )
        plate_b = int(cur.lastrowid)
        conn.commit()

    service = DeliveryScheduleService(db_path=db_path)
    plates = service._load_plates_meta(1)
    assert plate_a in plates and plate_b in plates

    _seed_completed_plates(
        db_path,
        kp_id=1,
        plate_name=plate_name,
        qty=10,
        length_m=6.0,
        width_m=1.2,
        load_class=800,
    )

    produced = service._load_produced_by_plate_id(1, plates)
    assert produced[plate_a] == 5
    assert produced[plate_b] == 5
    assert produced[plate_a] + produced[plate_b] == 10


def test_traffic_light_degraded_when_calendar_raises(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Ожидаемый сбой calendar → degraded=True, status=None; TypeError не глотаем."""
    _patch_empty_occupancy(monkeypatch)
    db_path = _fresh_db(tmp_path)
    plate_id = _seed_kp(
        db_path,
        kp_id=1,
        plate_qty=10,
        length_m=6.0,
        width_m=1.2,
        load_class=800,
    )
    service = DeliveryScheduleService(db_path=db_path)
    service._today_override = _TODAY
    service.replace(1, _payload(plate_id, qty=2), ADMIN)

    monkeypatch.setattr(
        "app.services.delivery_schedule_service.get_global_calendar_info",
        lambda: (_ for _ in ()).throw(RuntimeError("calendar down")),
    )
    view = service.get(1, user=ADMIN)
    assert view.traffic_light_degraded is True
    assert view.batches[0].status is None
    assert view.batches[0].ready_date is None
    assert view.batches[0].hint is None

    monkeypatch.setattr(
        "app.services.delivery_schedule_service.get_global_calendar_info",
        lambda: {"days_info": {}},
    )
    monkeypatch.setattr(
        "app.services.delivery_schedule_service.check_batches",
        lambda **_kwargs: (_ for _ in ()).throw(TypeError("bug in check")),
    )
    with pytest.raises(TypeError, match="bug in check"):
        service.get(1, user=ADMIN)
