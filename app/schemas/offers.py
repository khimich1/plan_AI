from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field


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
