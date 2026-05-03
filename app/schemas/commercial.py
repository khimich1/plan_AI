from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

CommercialFileKind = Literal["pdf", "xlsx", "breakdown", "schema"]
CommercialSourceType = Literal["text", "image"]
CommercialConditionsMode = Literal["standard", "custom"]
CommercialPlateUpdateMode = Literal["append", "replace"]
CommercialWidePlateAction = Literal["confirm", "exclude", "replace"]
CommercialSaveMode = Literal["database", "archive", "skip"]


class CommercialParseRequest(BaseModel):
    text: str = Field(min_length=1)


class CommercialPreviewRequest(BaseModel):
    text: str = Field(min_length=1)


class CommercialGeneratedFile(BaseModel):
    kind: CommercialFileKind
    filename: str
    display_name: str
    download_url: str


class CommercialWidePlateLine(BaseModel):
    id: str = Field(min_length=1)
    line: str
    qty: int = 1


class CommercialPlateBatch(BaseModel):
    source_type: CommercialSourceType
    original_text: str = ""
    normalized_text: str = ""
    ocr_text: str = ""
    filename: str = ""


class CommercialOfferIdentity(BaseModel):
    offer_number: str
    offer_date: str
    file_stem: str


class CommercialSavedOffer(BaseModel):
    kp_id: int | None = None
    status: str
    mode: CommercialSaveMode
    execution_terms: str = ""
    saved_at: str = ""


class CommercialDraftMetadata(BaseModel):
    source_type: CommercialSourceType | None = None
    original_text: str = ""
    ocr_text: str = ""
    input_text: str = ""
    accumulated_text: str = ""
    manager_id: int | None = None
    manager_name: str = ""
    manager_phone: str = ""
    manager_email: str = ""
    client_name: str = ""
    discount_percent: float = 0.0
    conditions_mode: CommercialConditionsMode = "standard"
    delivery_conditions: str = ""
    payment_conditions: str = ""
    warnings: list[str] = Field(default_factory=list)
    unparsed_lines: list[str] = Field(default_factory=list)
    normalized_text: str = ""
    normalized_lines: list[str] = Field(default_factory=list)
    wide_plate_lines: list[CommercialWidePlateLine] = Field(default_factory=list)
    diagnostics: list[dict[str, Any]] = Field(default_factory=list)
    price_rows_count: int = 0
    breakdown_tables_count: int = 0
    total_sum: float = 0.0
    plate_batches: list[CommercialPlateBatch] = Field(default_factory=list)
    wide_plates_resolved: bool = False
    last_source_filename: str = ""
    current_step: str = "plates"
    current_save_mode: CommercialSaveMode | None = None
    execution_terms: str = ""
    logistics_cost: float = 0.0


class CommercialDraftDetailsResponse(BaseModel):
    draft_id: str
    order: dict[str, Any]
    optimization: dict[str, Any]
    order_data: list[dict[str, Any]]
    metadata: CommercialDraftMetadata
    files: list[CommercialGeneratedFile] = Field(default_factory=list)
    saved_offer: CommercialSavedOffer | None = None
    totals: dict[str, Any]
    offer_identity: CommercialOfferIdentity


class CommercialCreateFromFormResponse(CommercialDraftDetailsResponse):
    pass


class CommercialDraftMetaUpdateRequest(BaseModel):
    manager_id: int | None = None
    client_name: str | None = None
    discount_percent: float | None = None
    conditions_mode: CommercialConditionsMode | None = None
    delivery_conditions: str | None = None
    payment_conditions: str | None = None
    logistics_cost: float | None = None


class CommercialWidePlateDecision(BaseModel):
    line_id: str | None = None
    source_line: str | None = None
    action: CommercialWidePlateAction
    replacement_text: str = ""


class CommercialWidePlatesResolveRequest(BaseModel):
    decisions: list[CommercialWidePlateDecision] = Field(min_length=1)


class CommercialGenerateFilesRequest(BaseModel):
    file_types: list[CommercialFileKind] = Field(default_factory=lambda: ["pdf", "xlsx", "breakdown", "schema"])


class CommercialGenerateFilesResponse(BaseModel):
    draft_id: str
    files: list[CommercialGeneratedFile]


class CommercialSaveDraftRequest(BaseModel):
    mode: CommercialSaveMode = "skip"
    execution_terms_input: str = ""


class CommercialSaveResultCard(BaseModel):
    kp_id: int | None = None
    offer_number: str
    offer_date: str
    client_name: str
    manager_name: str
    total_amount: float
    status: str
    execution_terms: str = ""


class CommercialSaveOfferResponse(BaseModel):
    draft_id: str
    saved_offer: CommercialSavedOffer | None = None
    totals: dict[str, Any]
    offer_identity: CommercialOfferIdentity
    result_card: CommercialSaveResultCard

