"""In-web notifications: list + mark-read (all authenticated users, own rows)."""

from __future__ import annotations

import json
from datetime import datetime

from fastapi import APIRouter, Depends, HTTPException, Query, status

from app.core.settings import get_settings
from app.dependencies.auth import get_current_user
from app.repositories.promise_repository import PromiseRepository
from app.schemas.notifications import (
    NotificationItem,
    NotificationListResponse,
    NotificationReadResponse,
)

router = APIRouter(prefix="/notifications", tags=["notifications"])


def get_promise_repository() -> PromiseRepository:
    return PromiseRepository(db_path=str(get_settings().plita_db_path))


def _parse_payload(raw: object) -> dict:
    if not raw:
        return {}
    if isinstance(raw, dict):
        return raw
    try:
        data = json.loads(str(raw))
    except json.JSONDecodeError:
        return {}
    return data if isinstance(data, dict) else {}


def _to_item(row: dict) -> NotificationItem:
    return NotificationItem(
        id=int(row["id"]),
        kind=str(row["kind"]),
        payload=_parse_payload(row.get("payload_json")),
        read_at=None if row.get("read_at") is None else str(row["read_at"]),
        created_at=str(row["created_at"]),
    )


@router.get("", response_model=NotificationListResponse)
def list_my_notifications(
    unread: bool = Query(default=False),
    user: dict = Depends(get_current_user),
    repo: PromiseRepository = Depends(get_promise_repository),
) -> NotificationListResponse:
    user_id = int(user["id"])
    rows = repo.list_notifications(user_id=user_id, unread_only=unread)
    unread_count = repo.count_unread_notifications(user_id=user_id)
    items = [_to_item(row) for row in reversed(rows)]
    return NotificationListResponse(items=items, unread_count=unread_count)


@router.post("/{notification_id}/read", response_model=NotificationReadResponse)
def mark_notification_read(
    notification_id: int,
    user: dict = Depends(get_current_user),
    repo: PromiseRepository = Depends(get_promise_repository),
) -> NotificationReadResponse:
    row = repo.mark_notification_read(
        notification_id,
        user_id=int(user["id"]),
        read_at=datetime.now(),
    )
    if row is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Notification not found",
        )
    return NotificationReadResponse(id=int(row["id"]), read_at=str(row["read_at"]))
