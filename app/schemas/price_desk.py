"""Pydantic-контракты API стола прайсов."""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field

PriceDeskKind = Literal["plates", "fbs", "march", "step", "bridge_pile", "pile"]

STATUS_KINDS: tuple[PriceDeskKind, ...] = (
    "plates",
    "fbs",
    "march",
    "step",
    "bridge_pile",
    "pile",
)


class PriceDeskDiffExample(BaseModel):
    key: str
    old_price: float | None = None
    new_price: float | None = None


class PriceDeskDiffExamples(BaseModel):
    changed: list[PriceDeskDiffExample] = Field(default_factory=list)
    new: list[PriceDeskDiffExample] = Field(default_factory=list)
    missing: list[PriceDeskDiffExample] = Field(default_factory=list)
    unchanged: list[PriceDeskDiffExample] = Field(default_factory=list)


class PriceDeskPreview(BaseModel):
    product_kind: PriceDeskKind
    price_list_date: str | None = None
    file_sha256: str
    parsed_rows: int
    changed: int
    new: int
    missing: int
    unchanged: int
    examples: PriceDeskDiffExamples


class PriceGroupStatus(BaseModel):
    product_kind: PriceDeskKind
    price_list_date: str | None = None
    imported_at: str | None = None
    row_count: int = 0


class PriceDeskStatus(BaseModel):
    groups: list[PriceGroupStatus]
