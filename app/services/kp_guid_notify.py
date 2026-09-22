"""Уведомления экономистам: в КП нет GUID для счёта (архив)."""

from __future__ import annotations

import json
import logging
import sqlite3
from collections.abc import Sequence
from datetime import datetime
from typing import Any

from app.core.constants import DEFAULT_ECONOMIST_ROLE
from app.repositories.auth_repository import AuthRepository
from app.repositories.promise_repository import PromiseRepository
from app.schemas.notifications import KIND_KP_GUID_MISSING
from core.guid_demand import record_guid_demand
from core.guid_gate import check_invoice_guids, order_lines_from_order_data

logger = logging.getLogger(__name__)

NOTIFICATION_KIND = KIND_KP_GUID_MISSING
_LINK = "/prices"


def _prepare_items(order_data: Sequence[object]) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    for raw in order_data:
        if isinstance(raw, dict):
            item = dict(raw)
        elif hasattr(raw, "model_dump"):
            item = raw.model_dump()
        else:
            continue
        if not str(item.get("product_type") or "").strip():
            item["product_type"] = "plates"
        items.append(item)
    return items


def _already_notified_marks(
    repo: PromiseRepository, *, user_id: int, kp_id: int
) -> set[str]:
    seen: set[str] = set()
    for row in repo.list_notifications(user_id=user_id, kind=NOTIFICATION_KIND):
        try:
            payload = json.loads(row.get("payload_json") or "{}")
        except json.JSONDecodeError:
            continue
        if not isinstance(payload, dict):
            continue
        if int(payload.get("kp_id") or 0) != int(kp_id):
            continue
        for mark in payload.get("marks") or []:
            text = str(mark).strip()
            if text:
                seen.add(text)
    return seen


def notify_kp_guid_missing(
    *,
    kp_id: int,
    order_data: Sequence[object],
    pb_db_path: str,
    plita_db_path: str,
    seq: int | None = None,
) -> int:
    """Резолвер по составу КП; фан-аут economist; дедуп (kp_id, марка).

    Возвращает число созданных уведомлений.
    """
    lines = order_lines_from_order_data(_prepare_items(order_data))
    if not lines:
        return 0
    conn = sqlite3.connect(pb_db_path)
    try:
        report = check_invoice_guids(lines, conn)
        if report.missing:
            record_guid_demand(conn, int(kp_id), report.missing)
            conn.commit()
    finally:
        conn.close()

    missing_marks: list[str] = []
    seen_marks: set[str] = set()
    for item in report.missing:
        mark = str(item.line.mark).strip()
        if not mark or mark in seen_marks:
            continue
        seen_marks.add(mark)
        missing_marks.append(mark)
    if not missing_marks:
        return 0

    economists = AuthRepository(db_path=plita_db_path).list_active_users_by_role(
        DEFAULT_ECONOMIST_ROLE
    )
    if not economists:
        logger.warning(
            "Нет активных экономистов — уведомление kp_guid_missing не отправлено (КП №%s)",
            kp_id,
        )
        return 0

    seq_value = int(seq if seq is not None else kp_id)
    repo = PromiseRepository(db_path=plita_db_path)
    created = 0
    now = datetime.now()
    for user in economists:
        user_id = int(user["id"])
        already = _already_notified_marks(repo, user_id=user_id, kp_id=int(kp_id))
        for mark in missing_marks:
            if mark in already:
                continue
            repo.create_notification(
                user_id=user_id,
                kind=NOTIFICATION_KIND,
                payload={
                    "kp_id": int(kp_id),
                    "seq": seq_value,
                    "marks": [mark],
                    "link": _LINK,
                },
                created_at=now,
            )
            created += 1
    return created
