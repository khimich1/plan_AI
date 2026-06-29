from __future__ import annotations

from pydantic import BaseModel, ConfigDict


class ManagerItem(BaseModel):
    model_config = ConfigDict(extra="allow")

    id: int
    fio: str | None = None
    contact_number: str | None = None
    email: str | None = None


class ManagerListResponse(BaseModel):
    items: list[ManagerItem]
    count: int
