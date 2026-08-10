"""Pydantic-контракты API «График поставки»."""

from __future__ import annotations

from datetime import date
from typing import Literal

from pydantic import BaseModel, Field, ValidationInfo, field_validator, model_validator


ScheduleStatus = Literal["draft", "active", "completed"]
TrafficLightStatus = Literal["green", "yellow", "red"]


def _parse_iso_date(value: str, *, field_name: str) -> str:
    """Нормализует дату к ISO YYYY-MM-DD; иначе ValueError с понятным текстом."""
    text = (value or "").strip()
    try:
        return date.fromisoformat(text).isoformat()
    except ValueError as exc:
        raise ValueError(
            f"{field_name} должна быть в формате ГГГГ-ММ-ДД, получено «{value}»"
        ) from exc


class BatchItemIn(BaseModel):
    plate_id: int = Field(ge=1)
    qty: int = Field(ge=1)


class BatchItemOut(BaseModel):
    plate_id: int
    qty: int
    plate_name: str | None = None
    # R4: qty партии > текущего qty позиции КП (позиция уменьшилась).
    changed: bool = False


class BatchIn(BaseModel):
    name: str = Field(min_length=1)
    deliver_from: str
    deliver_to: str
    produce_by: str
    items: list[BatchItemIn] = Field(default_factory=list)
    sort_order: int = 0

    @field_validator("name")
    @classmethod
    def _strip_name(cls, value: str) -> str:
        text = value.strip()
        if not text:
            raise ValueError("Название партии не может быть пустым")
        return text

    @field_validator("deliver_from", "deliver_to", "produce_by")
    @classmethod
    def _iso_dates(cls, value: str, info: ValidationInfo) -> str:
        return _parse_iso_date(value, field_name=str(info.field_name))

    @model_validator(mode="after")
    def _deliver_range(self) -> BatchIn:
        if self.deliver_from > self.deliver_to:
            raise ValueError(
                f"deliver_from ({self.deliver_from}) не может быть позже "
                f"deliver_to ({self.deliver_to})"
            )
        return self


class BatchOut(BaseModel):
    id: int
    name: str
    deliver_from: str
    deliver_to: str
    produce_by: str
    items: list[BatchItemOut] = Field(default_factory=list)
    sort_order: int = 0
    # Светофор заполняется в T5; в T4 остаётся null.
    status: TrafficLightStatus | None = None
    ready_date: str | None = None
    hint: str | None = None
    # R4: хотя бы одна позиция партии «changed».
    changed: bool = False


class DeliverySchedulePut(BaseModel):
    """Полная замена партий (идемпотентный PUT)."""

    invoice_number: str | None = None
    contract_number: str | None = None
    batches: list[BatchIn] = Field(default_factory=list)


class DeliveryScheduleView(BaseModel):
    id: int
    kp_id: int
    invoice_number: str | None = None
    contract_number: str | None = None
    status: ScheduleStatus = "draft"
    batches: list[BatchOut] = Field(default_factory=list)
    updated_at: str


class BatchDraftItemOut(BaseModel):
    plate_id: int
    plate_name: str
    qty: int


class BatchDraftOut(BaseModel):
    """Черновик партии из XLSX (без id — ещё не сохранён)."""

    name: str
    deliver_from: str
    deliver_to: str
    produce_by: str
    items: list[BatchDraftItemOut] = Field(default_factory=list)


class UnmatchedRowOut(BaseModel):
    row_number: int
    reason: str
    raw: dict | None = None


class ImportDraftResponse(BaseModel):
    """Результат POST /import: черновик партий без записи в БД."""

    batches: list[BatchDraftOut] = Field(default_factory=list)
    unmatched_rows: list[UnmatchedRowOut] = Field(default_factory=list)
