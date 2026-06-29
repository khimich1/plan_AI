from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field


class OfferSummary(BaseModel):
    model_config = ConfigDict(extra="allow")

    kp_id: int
    creation_date: str | None = None
    customer_name: str | None = None
    manager_name: str | None = None
    discount_percent: float = 0.0
    subtotal: float = 0.0
    vat_amount: float = 0.0
    total_amount: float = 0.0
    delivery_conditions: str | None = None
    payment_conditions: str | None = None
    execution_terms: str | None = None
    status: str = "в работе"
    completion_percentage: float = 0.0


class OfferPlateItem(BaseModel):
    model_config = ConfigDict(extra="allow")


class OfferDetails(OfferSummary):
    plates: list[OfferPlateItem] = Field(default_factory=list)


class OfferListResponse(BaseModel):
    items: list[OfferSummary]
    count: int


class CreateOfferResponse(BaseModel):
    kp_id: int
    status: str
    execution_terms: str | None = None
    used_default_execution_terms: bool = False
    offer: OfferSummary | None = None


class MoveToProductionResponse(BaseModel):
    kp_id: int
    execution_terms: str
    used_default_execution_terms: bool
    offer: OfferSummary


class DeleteOfferResponse(BaseModel):
    ok: bool = True
    kp_id: int


class OfferOrderItem(BaseModel):
    name: str = Field(min_length=1)
    length_m: float = Field(ge=0)
    width_m: float = Field(ge=0)
    qty: int = Field(ge=1)
    load_class: int = Field(default=800)
    unit_price: float = Field(ge=0)
    weight: float = Field(default=0, ge=0)
    length_dm_raw: str = ""
    nomenclature_id: int | str | None = None


class CreateOfferRequest(BaseModel):
    creation_date: str = Field(min_length=1)
    customer_name: str = Field(min_length=1)
    manager_name: str = Field(min_length=1)
    manager_phone: str = ""
    manager_email: str = ""
    discount_percent: float = Field(default=0, ge=0, le=100)
    delivery_conditions: str = ""
    payment_conditions: str = ""
    execution_terms_input: str = ""
    save_mode: Literal["work", "archive"] = "work"
    order_data: list[OfferOrderItem] = Field(min_length=1)


class UpdateOfferDiscountRequest(BaseModel):
    discount_percent: float = Field(ge=0, le=100)


class MoveOfferToProductionRequest(BaseModel):
    execution_terms_input: str = Field(min_length=1)
