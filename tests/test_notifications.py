"""GPS-022: уведомление kp_guid_missing при архиве КП."""

from __future__ import annotations

import json
import logging
import sqlite3
from pathlib import Path

import pytest

from app.repositories.auth_repository import AuthRepository
from app.repositories.promise_repository import PromiseRepository
from app.services.kp_guid_notify import NOTIFICATION_KIND, notify_kp_guid_missing
from core.guid_demand import list_open
from core.nomenclature_guid import ensure_schema, upsert
from tests.helpers import kp_db_fixtures as fx

_PWD = "StrongPassword123!"
_PILE_MARK = "С70.35-9у"
_READY_MARK = "С30.30-3"


def _init_pb(path: Path) -> None:
    conn = sqlite3.connect(path)
    try:
        ensure_schema(conn)
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS prays_plity (
                "Уникальный идентификатор (Номенклатура)" TEXT,
                "Товар" TEXT
            )
            """
        )
        conn.commit()
    finally:
        conn.close()


def _seed_missing_pile(pb: Path) -> None:
    conn = sqlite3.connect(pb)
    try:
        upsert(conn, "pile", "С70.35-9", match_status="missing")
        conn.commit()
    finally:
        conn.close()


def _seed_ready_pile(pb: Path) -> None:
    conn = sqlite3.connect(pb)
    try:
        upsert(
            conn,
            "pile",
            _READY_MARK,
            guid_1c="aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
            match_status="auto",
        )
        conn.commit()
    finally:
        conn.close()


def _make_economist(plita: str, username: str = "econ", *, active: bool = True) -> dict:
    repo = AuthRepository(db_path=plita)
    user, _created = repo.create_or_update_user(
        username=username,
        password=_PWD,
        role="economist",
        is_active=active,
    )
    return user


def _order_missing() -> list[dict]:
    return [{"product_type": "piles", "mark": _PILE_MARK, "qty": 2}]


def _order_ready() -> list[dict]:
    return [{"product_type": "piles", "mark": _READY_MARK, "qty": 1}]


def test_archive_with_missing_guid_notifies_each_economist(tmp_path: Path) -> None:
    plita = fx.make_iso_db(tmp_path)
    pb = tmp_path / "pb.db"
    _init_pb(pb)
    _seed_missing_pile(pb)
    first = _make_economist(plita, "econ_a")
    second = _make_economist(plita, "econ_b")
    _make_economist(plita, "econ_off", active=False)

    created = notify_kp_guid_missing(
        kp_id=12,
        seq=12,
        order_data=_order_missing(),
        pb_db_path=str(pb),
        plita_db_path=plita,
    )
    assert created == 2

    repo = PromiseRepository(db_path=plita)
    for user in (first, second):
        rows = repo.list_notifications(user_id=int(user["id"]), kind=NOTIFICATION_KIND)
        assert len(rows) == 1
        payload = json.loads(rows[0]["payload_json"])
        assert payload["kp_id"] == 12
        assert payload["seq"] == 12
        assert payload["marks"] == [_PILE_MARK]
        assert payload["link"] == "/prices"

    inactive = AuthRepository(db_path=plita).get_user_by_username("econ_off")
    assert inactive is not None
    assert repo.list_notifications(user_id=int(inactive["id"])) == []


def test_archive_without_holes_is_silent(tmp_path: Path) -> None:
    plita = fx.make_iso_db(tmp_path)
    pb = tmp_path / "pb.db"
    _init_pb(pb)
    _seed_ready_pile(pb)
    user = _make_economist(plita)

    created = notify_kp_guid_missing(
        kp_id=3,
        seq=3,
        order_data=_order_ready(),
        pb_db_path=str(pb),
        plita_db_path=plita,
    )
    assert created == 0
    repo = PromiseRepository(db_path=plita)
    assert repo.list_notifications(user_id=int(user["id"])) == []


def test_repeat_archive_does_not_duplicate(tmp_path: Path) -> None:
    plita = fx.make_iso_db(tmp_path)
    pb = tmp_path / "pb.db"
    _init_pb(pb)
    _seed_missing_pile(pb)
    user = _make_economist(plita)

    kwargs = dict(
        kp_id=7,
        seq=7,
        order_data=_order_missing(),
        pb_db_path=str(pb),
        plita_db_path=plita,
    )
    assert notify_kp_guid_missing(**kwargs) == 1
    assert notify_kp_guid_missing(**kwargs) == 0

    repo = PromiseRepository(db_path=plita)
    rows = repo.list_notifications(user_id=int(user["id"]), kind=NOTIFICATION_KIND)
    assert len(rows) == 1


def test_no_economist_logs_warning(tmp_path: Path, caplog: pytest.LogCaptureFixture) -> None:
    plita = fx.make_iso_db(tmp_path)
    pb = tmp_path / "pb.db"
    _init_pb(pb)
    _seed_missing_pile(pb)
    AuthRepository(db_path=plita).create_or_update_user(
        username="manager",
        password=_PWD,
        role="manager",
    )

    with caplog.at_level(logging.WARNING, logger="app.services.kp_guid_notify"):
        created = notify_kp_guid_missing(
            kp_id=1,
            seq=1,
            order_data=_order_missing(),
            pb_db_path=str(pb),
            plita_db_path=plita,
        )
    assert created == 0
    assert any("экономист" in rec.message.lower() for rec in caplog.records)
    conn = sqlite3.connect(pb)
    try:
        rows = list_open(conn)
    finally:
        conn.close()
    assert len(rows) == 1
    assert rows[0].mark == _PILE_MARK
    assert list(rows[0].kp_ids) == [1]


def test_archive_without_holes_does_not_write_demand(tmp_path: Path) -> None:
    plita = fx.make_iso_db(tmp_path)
    pb = tmp_path / "pb.db"
    _init_pb(pb)
    _seed_ready_pile(pb)

    created = notify_kp_guid_missing(
        kp_id=3,
        seq=3,
        order_data=_order_ready(),
        pb_db_path=str(pb),
        plita_db_path=plita,
    )
    assert created == 0
    conn = sqlite3.connect(pb)
    try:
        assert list_open(conn) == []
    finally:
        conn.close()


def test_repeat_archive_notifies_only_new_mark(tmp_path: Path) -> None:
    """A5: дедуп по паре (КП, марка) — новая марка на том же КП приходит отдельно."""
    plita = fx.make_iso_db(tmp_path)
    pb = tmp_path / "pb.db"
    _init_pb(pb)
    _seed_missing_pile(pb)
    conn = sqlite3.connect(pb)
    try:
        upsert(conn, "pile", "С30.30-3", match_status="missing")
        conn.commit()
    finally:
        conn.close()
    user = _make_economist(plita)

    first = notify_kp_guid_missing(
        kp_id=9,
        seq=9,
        order_data=_order_missing(),
        pb_db_path=str(pb),
        plita_db_path=plita,
    )
    second = notify_kp_guid_missing(
        kp_id=9,
        seq=9,
        order_data=_order_missing() + [{"product_type": "piles", "mark": "С30.30-3", "qty": 1}],
        pb_db_path=str(pb),
        plita_db_path=plita,
    )
    assert first == 1
    assert second == 1

    repo = PromiseRepository(db_path=plita)
    rows = repo.list_notifications(user_id=int(user["id"]), kind=NOTIFICATION_KIND)
    marks = [json.loads(row["payload_json"])["marks"][0] for row in rows]
    assert sorted(marks) == ["С30.30-3", _PILE_MARK]

    conn = sqlite3.connect(pb)
    try:
        demand = {(row.mark, tuple(row.kp_ids)) for row in list_open(conn)}
    finally:
        conn.close()
    assert demand == {(_PILE_MARK, (9,)), ("С30.30-3", (9,))}
