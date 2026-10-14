"""Разбор otvet_*.json: номер счёта на КП и одно оповещение владельцу."""

from __future__ import annotations

import json
import sqlite3
from pathlib import Path

_UID_ORDER = "f335beda-be5f-11f1-9009-04d9f5387845"
_UID_KP = "f335beca-be5f-11f1-9009-04d9f5387845"
_REPLY = {
    "УИД_ЗК": _UID_ORDER,
    "Номер_ЗК": "703",
    "УИД_КП": _UID_KP,
    "Номер_КП": "46",
    "Описание": "Документ успешно сформирован",
}


def _seed(tmp_path: Path, *, kp_id: int = 46, owner_user_id: int | None = 7) -> str:
    from core.kp_db_schema import init_schema

    db_path = str(tmp_path / "plita.db")
    init_schema(db_path)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            INSERT INTO KP_offers (kp_id, creation_date, customer_name)
            VALUES (?, '02.10.2026', 'КАПИТАЛСТРОЙ ООО')
            """,
            (kp_id,),
        )
        conn.execute(
            "INSERT INTO kp_meta (kp_id, status, owner_user_id) VALUES (?, 'на согласовании', ?)",
            (kp_id, owner_user_id),
        )
    return db_path


def _write_reply(directory: Path, name: str = "otvet_190064.json", body: dict | None = None) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / name
    path.write_text(json.dumps(body if body is not None else _REPLY, ensure_ascii=False), encoding="utf-8")
    return path


def _order_ids(db_path: str, kp_id: int = 46) -> tuple | None:
    with sqlite3.connect(db_path) as conn:
        return conn.execute(
            """
            SELECT order_number_1c, uid_order_1c, uid_kp_1c
            FROM KP_offers WHERE kp_id = ?
            """,
            (kp_id,),
        ).fetchone()


def _notifications(db_path: str) -> list[tuple]:
    with sqlite3.connect(db_path) as conn:
        return conn.execute(
            "SELECT user_id, kind, payload_json FROM notifications ORDER BY id"
        ).fetchall()


def test_import_writes_invoice_number_and_notifies_owner(tmp_path: Path) -> None:
    from app.services.invoice_reply_import import import_otvet_dir

    db_path = _seed(tmp_path)
    inbox = tmp_path / "from_1c"
    _write_reply(inbox)

    import_otvet_dir(inbox, db_path)

    assert _order_ids(db_path) == ("703", _UID_ORDER, _UID_KP)
    assert list(inbox.glob("otvet_*.json")) == []
    notes = _notifications(db_path)
    assert len(notes) == 1
    user_id, kind, payload_json = notes[0]
    assert user_id == 7
    assert kind == "invoice_number_received"
    assert json.loads(payload_json) == {"kp_id": 46, "order_number": "703"}


def test_repeat_same_reply_deletes_file_without_second_notice(tmp_path: Path) -> None:
    from app.services.invoice_reply_import import import_otvet_dir

    db_path = _seed(tmp_path)
    inbox = tmp_path / "from_1c"
    _write_reply(inbox)
    import_otvet_dir(inbox, db_path)
    _write_reply(inbox, name="otvet_again.json")

    import_otvet_dir(inbox, db_path)

    assert _order_ids(db_path) == ("703", _UID_ORDER, _UID_KP)
    assert list(inbox.glob("otvet_*.json")) == []
    assert len(_notifications(db_path)) == 1


def test_conflicting_uid_keeps_row_and_file(tmp_path: Path) -> None:
    from app.services.invoice_reply_import import import_otvet_dir

    db_path = _seed(tmp_path)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            UPDATE KP_offers
            SET order_number_1c = ?, uid_order_1c = ?, uid_kp_1c = ?
            WHERE kp_id = 46
            """,
            ("111", "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa", _UID_KP),
        )
    inbox = tmp_path / "from_1c"
    path = _write_reply(inbox)

    import_otvet_dir(inbox, db_path)

    assert _order_ids(db_path) == ("111", "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa", _UID_KP)
    assert path.is_file()
    assert _notifications(db_path) == []


def test_unknown_kp_leaves_file(tmp_path: Path) -> None:
    from app.services.invoice_reply_import import import_otvet_dir

    db_path = _seed(tmp_path, kp_id=46)
    inbox = tmp_path / "from_1c"
    path = _write_reply(inbox, body={**_REPLY, "Номер_КП": "99"})

    import_otvet_dir(inbox, db_path)

    assert path.is_file()
    assert _order_ids(db_path, 46) == (None, None, None)
    assert _notifications(db_path) == []
