"""Pydantic-контракты «почему нет GUID»: причина, та же карточка, срез и поиск."""

from __future__ import annotations

from pydantic import BaseModel, Field


class GuidCardOut(BaseModel):
    """Карточка 1С для кнопки: имя как в выгрузке и GUID, который уйдёт в замену."""

    name: str
    guid: str


class GuidCandidateOut(BaseModel):
    mark: str
    guid_count: int
    label: str | None = None
    selectable: bool = False
    cards: list[GuidCardOut] = Field(default_factory=list)


class GuidExplainResponse(BaseModel):
    product_kind: str
    mark: str
    reason: str = ""
    hint: str = ""
    same_cards: list[GuidCandidateOut] = Field(default_factory=list)
    slice: list[GuidCandidateOut] = Field(default_factory=list)


class GuidSearchResponse(BaseModel):
    items: list[GuidCandidateOut] = Field(default_factory=list)


class GuidReplaceRequest(BaseModel):
    product_kind: str = Field(min_length=1)
    mark: str = Field(min_length=1)
    target_mark: str = Field(min_length=1)
    nomenclature_id: str | None = None
    replace_oneoff: bool = False


class DraftGuidReplaceRequest(BaseModel):
    """Замена менеджера: номер карточки сервер берёт сам, договорную цену не трогает."""

    product_kind: str = Field(min_length=1)
    mark: str = Field(min_length=1)
    target_mark: str = Field(min_length=1)


class GuidReplaceSkipOut(BaseModel):
    kp_id: int
    reason: str


class GuidReplaceMoneyOut(BaseModel):
    unit_price: float
    unit_discounted: float
    line_sum: float
    total: float
    subtotal: float
    vat_amount: float


class GuidReplacePreviewItemOut(BaseModel):
    kp_id: int
    qty: int
    was: GuidReplaceMoneyOut
    will: GuidReplaceMoneyOut
    delivery_was: float | None = None
    delivery_will: float | None = None
    oneoff_notice: str | None = None


class GuidReplacePreviewResponse(BaseModel):
    items: list[GuidReplacePreviewItemOut] = Field(default_factory=list)
    skipped: list[GuidReplaceSkipOut] = Field(default_factory=list)


class GuidReplaceWrittenOut(BaseModel):
    kp_id: int
    total: float


class GuidReplaceResponse(BaseModel):
    written: list[GuidReplaceWrittenOut] = Field(default_factory=list)
    skipped: list[GuidReplaceSkipOut] = Field(default_factory=list)
    queue_notice: str | None = None
