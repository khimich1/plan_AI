"""Pydantic schemas for SGP (склад готовой продукции) API."""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field


class SgpProgress(BaseModel):
    n: int = Field(ge=0, description="Qty on SGP linked to this KP")
    m: int = Field(ge=0, description="Ordered qty snapshot")


class SgpPlateItem(BaseModel):
    id: int
    kp_id: int | None = None
    plate_name: str
    length_m: float | None = None
    width_m: float | None = None
    load_class: int | None = None
    qty: int
    completed_date: str | None = None
    production_day: int | None = None
    plan_id: str | None = None
    nomenclature_id: str | None = None
    customer_name: str | None = None
    execution_terms: str | None = None
    sgp_progress: SgpProgress | None = None


class SgpPlatesResponse(BaseModel):
    items: list[SgpPlateItem]
    count: int
    filter: Literal["all", "linked", "unlinked"] = "all"


class SgpUnlinkRequest(BaseModel):
    qty: int = Field(ge=1)


class SgpRelinkRequest(BaseModel):
    target_kp_id: int = Field(ge=1)
    qty: int = Field(ge=1)


class SgpMutationResponse(BaseModel):
    ok: bool = True
    sgp_id: int
    qty: int
    kp_id: int | None = None
    target_kp_id: int | None = None
    message: str = ""


class SgpFreePlateItem(BaseModel):
    id: int
    plate_name: str
    length_m: float | None = None
    width_m: float | None = None
    load_class: int | None = None
    qty: int
    completed_date: str | None = None


class SgpFreePlatesResponse(BaseModel):
    items: list[SgpFreePlateItem]
    count: int
