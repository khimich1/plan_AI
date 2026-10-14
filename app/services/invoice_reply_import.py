"""Читает ``otvet_*.json`` и записывает номер счёта на КП."""

from __future__ import annotations

import json
import logging
from datetime import datetime
from pathlib import Path

from app.repositories.promise_repository import PromiseRepository
from app.schemas.notifications import KIND_INVOICE_NUMBER_RECEIVED
from core.invoice_reply import InvoiceReplyError, parse_invoice_reply
from core.kp_db_schema import KpOrderIdConflictError, record_kp_order_ids

logger = logging.getLogger(__name__)

_APPLIED = "applied"
_UNCHANGED = "unchanged"
_MISSING = "missing"
_CONFLICT = "conflict"
_INVALID = "invalid"


def import_otvet_dir(directory: Path, db_path: str) -> None:
    if not directory.is_dir():
        return
    for path in sorted(directory.glob("otvet_*.json")):
        if path.is_file():
            import_otvet_file(path, db_path)


def import_otvet_file(path: Path, db_path: str) -> str:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        reply = parse_invoice_reply(payload)
    except (OSError, UnicodeError, json.JSONDecodeError, InvoiceReplyError) as exc:
        logger.warning("invoice reply skipped %s: %s", path.name, exc)
        return _INVALID

    outcome = _apply(db_path, reply)
    if outcome in {_APPLIED, _UNCHANGED}:
        path.unlink()
    elif outcome == _MISSING:
        logger.warning("invoice reply %s: КП №%s нет в базе", path.name, reply.kp_id)
    elif outcome == _CONFLICT:
        logger.warning(
            "invoice reply %s: идентификаторы КП №%s уже другие",
            path.name,
            reply.kp_id,
        )
    return outcome


def _apply(db_path: str, reply) -> str:
    repo = PromiseRepository(db_path=db_path)
    conn = repo._connect()
    try:
        row = conn.execute(
            """
            SELECT order_number_1c, uid_order_1c, uid_kp_1c
            FROM KP_offers WHERE kp_id = ?
            """,
            (reply.kp_id,),
        ).fetchone()
        if row is None:
            return _MISSING
        before = tuple(row)
        try:
            record_kp_order_ids(
                conn,
                reply.kp_id,
                order_number=reply.order_number,
                uid_order=reply.uid_order,
                uid_kp=reply.uid_kp,
            )
        except KpOrderIdConflictError:
            conn.rollback()
            return _CONFLICT
        after = conn.execute(
            """
            SELECT order_number_1c, uid_order_1c, uid_kp_1c
            FROM KP_offers WHERE kp_id = ?
            """,
            (reply.kp_id,),
        ).fetchone()
        applied = tuple(after) != before
        if applied:
            owner = conn.execute(
                "SELECT owner_user_id FROM kp_meta WHERE kp_id = ?",
                (reply.kp_id,),
            ).fetchone()
            owner_id = None if owner is None else owner[0]
            if owner_id is not None:
                repo._insert_notification(
                    conn,
                    user_id=int(owner_id),
                    kind=KIND_INVOICE_NUMBER_RECEIVED,
                    payload={"kp_id": reply.kp_id, "order_number": reply.order_number},
                    created_at=datetime.now(),
                )
        conn.commit()
        return _APPLIED if applied else _UNCHANGED
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
