from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, ConfigDict, Field


ArchiveSection = Literal["archived", "in_production", "completed"]
ArchiveFileKind = Literal["pdf", "xlsx", "schema"]


class ArchiveOfferListItem(BaseModel):
    """Строка списка архива (аналог inline-кнопки в боте)."""

    model_config = ConfigDict(populate_by_name=True)

    kp_id: int
    creation_date: str | None = None
    customer_name: str | None = None
    manager_name: str | None = None
    discount_percent: float = 0.0
    subtotal: float = 0.0
    vat_amount: float = 0.0
    total_amount: float = 0.0
    execution_terms: str | None = None
    status: str | None = None
    completion_percentage: float | None = Field(
        default=None,
        description="Процент выполнения плит (для разделов in_production / completed).",
    )


class ArchivePlateItem(BaseModel):
    position_number: int | None = None
    plate_name: str = ""
    length_m: float | None = None
    width_m: float | None = None
    load_class: int | None = None
    qty: int = 0
    unit_price: float | None = None
    discounted_price: float | None = None
    unit_weight: float | None = None
    total_weight: float | None = None
    status: str | None = None


class ArchiveOfferFinance(BaseModel):
    subtotal: float = 0.0
    vat_amount: float = 0.0
    total_amount: float = 0.0
    discount_percent: float = 0.0


class ArchiveOfferDetails(BaseModel):
    """Полная карточка КП для страницы архива."""

    kp_id: int
    creation_date: str | None = None
    customer_name: str | None = None
    manager_name: str | None = None
    status: str | None = None
    execution_terms: str | None = None
    delivery_conditions: str | None = None
    payment_conditions: str | None = None
    finance: ArchiveOfferFinance
    logistics_cost: float = Field(default=0.0, description="Стоимость одного рейса (как logistics_cost при создании КП).")
    total_cargo_weight_kg: float = Field(default=0.0, description="Суммарная масса по строкам через resolve_kp_line_weight_kg (как PDF/XLSX).")
    delivery_service_total_rub: float = Field(
        default=0.0,
        description="Строка «Услуга по доставке грузов»: рейсы × стоимость рейса.",
    )
    plates: list[ArchivePlateItem] = Field(default_factory=list)
    completion_percentage: float | None = None


class UpdateDiscountRequest(BaseModel):
    discount: float = Field(ge=0, le=100)


class UpdateLogisticsCostRequest(BaseModel):
    logistics_cost: float = Field(ge=0, description="Новая стоимость одного рейса.")


class MoveToProductionRequest(BaseModel):
    execution_terms: str = Field(min_length=1, max_length=128)


class ArchiveSearchResponse(BaseModel):
    mode: Literal["number", "customer"]
    items: list[ArchiveOfferListItem] = Field(default_factory=list)
    total: int = 0
    truncated: bool = False
