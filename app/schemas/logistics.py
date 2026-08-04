"""Pydantic schemas for Logistics (раздел «Логистика», SHIP-xxx)."""

from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field

DeliveryTypeStr = Literal["delivery", "pickup"]
ShipmentStatusStr = Literal["in_work", "done"]
ShipmentItemTypeStr = Literal["plate", "free"]


# ---------- Справочники ----------


class CarrierItem(BaseModel):
    id: int
    name: str
    source_sheet: str | None = None
    note: str | None = None
    active: bool = True
    merged_into_id: int | None = None
    shipments_count: int = 0


class CarrierListResponse(BaseModel):
    items: list[CarrierItem]
    count: int


class CarrierMergeRequest(BaseModel):
    into_id: int = Field(ge=1)


class CarrierMergeResponse(BaseModel):
    ok: bool = True
    carrier_id: int
    into_id: int
    moved_shipments: int
    message: str = ""


class PileCatalogItem(BaseModel):
    id: int
    mark: str
    length_m: float | None = None
    section_mm: int | None = None
    volume_m3: float | None = None
    weight_kg: float
    pcs_per_20t: int | None = None


class PileCatalogResponse(BaseModel):
    items: list[PileCatalogItem]
    count: int


# ---------- Поиск КП (ACL B: в работе / На СГП) ----------


class LogisticsKpSearchItem(BaseModel):
    """Минимальный read-model для поиска КП логистикой (без финансов)."""

    kp_id: int
    customer_name: str | None = None
    status: str | None = None
    product_type: Literal["plates", "piles"] = "plates"


class LogisticsKpSearchResponse(BaseModel):
    mode: Literal["number", "customer"]
    items: list[LogisticsKpSearchItem] = Field(default_factory=list)
    total: int = 0
    truncated: bool = False


# ---------- Рейсы ----------


class ShipmentCreateRequest(BaseModel):
    shipment_date: str = Field(min_length=1, max_length=32)
    delivery_type: DeliveryTypeStr
    kp_ids: list[int] = Field(min_length=1)


class ShipmentOrderPatch(BaseModel):
    kp_id: int = Field(ge=1)
    ya_order_no: str | None = None


class ShipmentPatchRequest(BaseModel):
    shipment_date: str | None = None
    delivery_type: DeliveryTypeStr | None = None
    attention: bool | None = None
    attention_comment: str | None = None
    carrier_id: int | None = None
    driver_name: str | None = None
    vehicle_text: str | None = None
    vehicle_class: str | None = None
    proxy_no: str | None = None
    upd_no: str | None = None
    freight_request_no: str | None = None
    planned_cost: float | None = None
    time_slot: str | None = None
    orders: list[ShipmentOrderPatch] | None = None


class ShipmentOrderItem(BaseModel):
    id: int
    kp_id: int | None = None
    ya_order_no: str | None = None
    customer_name: str | None = None


class ShipmentItemInput(BaseModel):
    item_type: ShipmentItemTypeStr
    completed_plate_id: int | None = None
    kp_id: int | None = None
    mark: str | None = None
    qty: int = Field(ge=1)
    weight_kg: float | None = Field(default=None, description="Ручная правка веса строки")
    sort_order: int | None = None
    note: str | None = None


class ShipmentItemsPutRequest(BaseModel):
    items: list[ShipmentItemInput] = Field(default_factory=list)


class ShipmentItem(BaseModel):
    id: int
    item_type: str
    completed_plate_id: int | None = None
    kp_id: int | None = None
    mark: str | None = None
    plate_name: str | None = None
    length_m: float | None = None
    width_m: float | None = None
    load_class: int | None = None
    qty: int
    unit_weight_kg: float | None = None
    weight_kg: float | None = None
    sort_order: int = 0
    note: str | None = None


class ShipmentAvailableSgpRow(BaseModel):
    completed_plate_id: int
    kp_id: int | None = None
    plate_name: str
    length_m: float | None = None
    width_m: float | None = None
    load_class: int | None = None
    qty: int
    available_qty: int
    unit_weight_kg: float | None = None
    completed_date: str | None = None


class ShipmentAvailableByKp(BaseModel):
    kp_id: int
    plates: list[ShipmentAvailableSgpRow] = Field(default_factory=list)


class ShipmentCard(BaseModel):
    id: int
    shipment_date: str
    delivery_type: str
    status: str
    attention: bool = False
    attention_comment: str | None = None
    carrier_id: int | None = None
    carrier_name: str | None = None
    driver_name: str | None = None
    vehicle_text: str | None = None
    vehicle_class: str | None = None
    proxy_no: str | None = None
    upd_no: str | None = None
    freight_request_no: str | None = None
    planned_cost: float | None = None
    time_slot: str | None = None
    completed_at: str | None = None
    actor: str | None = None
    created_at: str | None = None
    orders: list[ShipmentOrderItem] = Field(default_factory=list)
    items: list[ShipmentItem] = Field(default_factory=list)
    total_weight_kg: float = 0.0
    available_by_kp: list[ShipmentAvailableByKp] = Field(default_factory=list)


class ShipmentListItem(BaseModel):
    id: int
    shipment_date: str
    delivery_type: str
    status: str
    attention: bool = False
    attention_comment: str | None = None
    carrier_id: int | None = None
    carrier_name: str | None = None
    driver_name: str | None = None
    vehicle_text: str | None = None
    vehicle_class: str | None = None
    proxy_no: str | None = None
    upd_no: str | None = None
    freight_request_no: str | None = None
    planned_cost: float | None = None
    time_slot: str | None = None
    created_at: str | None = None
    orders: list[ShipmentOrderItem] = Field(default_factory=list)
    total_weight_kg: float = 0.0


class ShipmentListResponse(BaseModel):
    items: list[ShipmentListItem]
    count: int


class ShipmentProposeRequest(BaseModel):
    vehicle_class: str | None = None


class ShipmentProposeItem(BaseModel):
    item_type: str = "plate"
    completed_plate_id: int
    kp_id: int | None = None
    plate_name: str
    length_m: float | None = None
    width_m: float | None = None
    load_class: int | None = None
    qty: int
    available_qty: int
    unit_weight_kg: float | None = None
    weight_kg: float | None = None
    completed_date: str | None = None
    reason_code: str | None = None
    reason_text: str | None = None


class ShipmentProposeWarning(BaseModel):
    code: str
    message: str
    kp_ids: list[int] = Field(default_factory=list)


class ShipmentOrderRemainderItem(BaseModel):
    completed_plate_id: int
    kp_id: int
    plate_name: str
    qty_remaining: int


class ShipmentLayoutUnit(BaseModel):
    completed_plate_id: int
    kp_id: int
    plate_name: str
    width_m: float | None = None


class ShipmentLayoutTier(BaseModel):
    index: int
    units: list[ShipmentLayoutUnit] = Field(default_factory=list)


class ShipmentLayoutStack(BaseModel):
    index: int
    marking_length_m: float
    tiers: list[ShipmentLayoutTier] = Field(default_factory=list)


class ShipmentLoadingStep(BaseModel):
    step: int
    stack_index: int
    tier_index: int
    description: str


class ShipmentLayoutMetadata(BaseModel):
    body_length_m: float
    body_used_m: float
    stacks: list[ShipmentLayoutStack] = Field(default_factory=list)
    loading_steps: list[ShipmentLoadingStep] = Field(default_factory=list)


class ShipmentProposeResponse(BaseModel):
    items: list[ShipmentProposeItem] = Field(default_factory=list)
    not_fit: list[ShipmentProposeItem] = Field(default_factory=list)
    order_remainder: list[ShipmentOrderRemainderItem] = Field(default_factory=list)
    warnings: list[ShipmentProposeWarning] = Field(default_factory=list)
    total_weight_kg: float = 0.0
    overload: bool = False
    vehicle_class: str | None = None
    vehicle_class_limits_kg: dict[str, int] = Field(default_factory=dict)
    layout: ShipmentLayoutMetadata | None = None


class ShipmentMutationResponse(BaseModel):
    ok: bool = True
    shipment_id: int
    status: str
    message: str = ""
