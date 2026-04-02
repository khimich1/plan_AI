from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field


class CommercialParseRequest(BaseModel):
    text: str = Field(min_length=1)


class CommercialPreviewRequest(BaseModel):
    text: str = Field(min_length=1)


class CommercialGeneratedFile(BaseModel):
    kind: Literal["pdf", "xlsx", "breakdown", "schema"]
    filename: str
    display_name: str
    download_url: str


class CommercialDraftMetadata(BaseModel):
    source_type: Literal["text", "image"] | None = None
    original_text: str = ""
    ocr_text: str = ""
    input_text: str = ""
    manager_id: int | None = None
    manager_name: str = ""
    manager_phone: str = ""
    manager_email: str = ""
    client_name: str = ""
    discount_percent: float = 0.0
    delivery_conditions: str = ""
    payment_conditions: str = ""
    warnings: list[str] = Field(default_factory=list)
    unparsed_lines: list[str] = Field(default_factory=list)
    normalized_text: str = ""
    normalized_lines: list[str] = Field(default_factory=list)
    wide_plate_lines: list[Any] = Field(default_factory=list)
    diagnostics: list[dict[str, Any]] = Field(default_factory=list)
    price_rows_count: int = 0
    breakdown_tables_count: int = 0
    total_sum: float = 0.0


class CommercialDraftDetailsResponse(BaseModel):
    draft_id: str
    order: dict[str, Any]
    optimization: dict[str, Any]
    order_data: list[dict[str, Any]]
    metadata: CommercialDraftMetadata
    files: list[CommercialGeneratedFile] = Field(default_factory=list)
    saved_offer: dict[str, Any] | None = None
    totals: dict[str, Any]


class CommercialCreateFromFormResponse(CommercialDraftDetailsResponse):
    pass


class CommercialGenerateFilesRequest(BaseModel):
    file_types: list[Literal["pdf", "xlsx", "breakdown", "schema"]] = Field(
        default_factory=lambda: ["pdf", "xlsx", "breakdown", "schema"]
    )


class CommercialGenerateFilesResponse(BaseModel):
    draft_id: str
    files: list[CommercialGeneratedFile]


class CommercialSaveOfferResponse(BaseModel):
    draft_id: str
    kp_id: int
    status: str
    totals: dict[str, Any]

