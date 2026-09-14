"""In-web notifications (promise exclusions and later lifecycle events)."""

from __future__ import annotations

from typing import Any

from pydantic import BaseModel, ConfigDict, Field


class NotificationItem(BaseModel):
    model_config = ConfigDict(extra="ignore")

    id: int
    kind: str
    payload: dict[str, Any] = Field(default_factory=dict)
    read_at: str | None = None
    created_at: str


class NotificationListResponse(BaseModel):
    items: list[NotificationItem]
    unread_count: int


class NotificationReadResponse(BaseModel):
    id: int
    read_at: str
