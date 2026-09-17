"""Pydantic-контракты API номенклатуры 1С (GUID)."""

from __future__ import annotations

from pydantic import BaseModel, Field

IMPORT_1C_LIST_LIMIT = 50


class GuidWriteOut(BaseModel):
    product_kind: str
    mark: str
    field: str
    guid: str


class AmbiguousMatchOut(BaseModel):
    product_kind: str
    mark: str | None = None
    name: str
    guids: list[str] = Field(default_factory=list)
    note: str = ""


class DisappearedMarkOut(BaseModel):
    product_kind: str
    mark: str
    guid_1c: str | None = None
    match_status: str


class Unmatched1COut(BaseModel):
    name: str
    guid: str
    row_index: int
    source_file: str = ""


class Create1cTaskOut(BaseModel):
    product_kind: str
    mark: str
    hint: str
    field: str = "guid_1c"


class PriceTaskOut(BaseModel):
    guid: str
    product_kind: str
    mark: str
    name: str


class DuplicateCandidateOut(BaseModel):
    guid: str
    name: str
    price: float | None = None


class DuplicateTaskOut(BaseModel):
    scope: str
    key: str
    product_kind: str
    candidates: list[DuplicateCandidateOut] = Field(default_factory=list)


class GuidTasksResponse(BaseModel):
    to_create_1c: list[Create1cTaskOut] = Field(default_factory=list)
    to_price: list[PriceTaskOut] = Field(default_factory=list)
    duplicates: list[DuplicateTaskOut] = Field(default_factory=list)


class PriceQueueResolveRequest(BaseModel):
    guid: str
    price: float


class PriceQueueResolveResponse(BaseModel):
    guid: str
    mark: str
    product_kind: str
    price: float
    state: str


class DuplicateResolveRequest(BaseModel):
    scope: str
    key: str
    chosen_guid: str
    note: str = ""
    product_kind: str | None = None


class DuplicateResolveResponse(BaseModel):
    scope: str
    key: str
    chosen_guid: str
    match_status: str | None = None


class Import1cResponse(BaseModel):
    """Отчёт сверки POST /api/v1/nomenclature/import-1c."""

    product_kind: str
    summary: str
    new_guids_count: int
    waiting_price: int
    ambiguous_count: int
    disappeared_count: int
    unmatched_1c_count: int
    unchanged: int
    updated_guids_count: int
    mode: str = "full"
    weights_updated: int = 0
    list_limit: int = IMPORT_1C_LIST_LIMIT
    new_guids: list[GuidWriteOut] = Field(default_factory=list)
    updated_guids: list[GuidWriteOut] = Field(default_factory=list)
    ambiguous: list[AmbiguousMatchOut] = Field(default_factory=list)
    disappeared: list[DisappearedMarkOut] = Field(default_factory=list)
    unmatched_1c: list[Unmatched1COut] = Field(default_factory=list)
