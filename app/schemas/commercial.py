from __future__ import annotations

from enum import Enum
from typing import Annotated, Any, Literal

from pydantic import BaseModel, BeforeValidator, ConfigDict, Field

CommercialFileKind = Literal["pdf", "xlsx", "breakdown", "schema"]
CommercialSourceType = Literal["text", "image", "ai"]
CommercialConditionsMode = Literal["standard", "custom"]
CommercialPlateUpdateMode = Literal["append", "replace"]
CommercialPileUpdateMode = Literal["append", "replace"]
CommercialStepUpdateMode = Literal["append", "replace"]
CommercialMarchUpdateMode = Literal["append", "replace"]
CommercialWidePlateAction = Literal["confirm", "exclude", "replace"]
CommercialUnpricedPlateAction = Literal["replace_load", "exclude"]
CommercialInvalidWidthAction = Literal["replace_width", "exclude"]
CommercialSaveMode = Literal["database", "archive", "skip"]
ProductType = Literal["plates", "piles", "steps", "marches", "bridge_piles", "fbs"]


class WizardStepId(str, Enum):
    """Канонические шаги мастера КП (совпадают с JSON-значениями на фронтенде)."""

    plates = "plates"
    piles = "piles"
    steps = "steps"
    marches = "marches"
    bridge_piles = "bridge_piles"
    fbs = "fbs"
    client = "client"
    result = "result"


class WizardNextRequiredAction(str, Enum):
    """Следующее обязательное действие относительно состояния черновика."""

    none = "none"
    ingest_plates = "ingest_plates"
    ingest_piles = "ingest_piles"
    ingest_steps = "ingest_steps"
    ingest_marches = "ingest_marches"
    ingest_bridge_piles = "ingest_bridge_piles"
    ingest_fbs = "ingest_fbs"
    resolve_wide_plates = "resolve_wide_plates"
    resolve_invalid_widths = "resolve_invalid_widths"
    resolve_unpriced_plates = "resolve_unpriced_plates"
    select_manager = "select_manager"
    complete_client_terms = "complete_client_terms"
    post_calculate = "post_calculate"


def _coerce_wizard_step_id(value: Any) -> WizardStepId:
    if isinstance(value, WizardStepId):
        return value
    raw = str(value or "").strip().lower()
    if not raw:
        return WizardStepId.plates
    legacy_aliases = {
        "wide-plates": WizardStepId.plates,
        "wide_plates": WizardStepId.plates,
        "manager": WizardStepId.client,
        "calculate": WizardStepId.client,
    }
    if raw in legacy_aliases:
        return legacy_aliases[raw]
    try:
        return WizardStepId(raw)
    except ValueError:
        return WizardStepId.plates


WizardStepCoerced = Annotated[WizardStepId, BeforeValidator(_coerce_wizard_step_id)]


class CommercialWizardState(BaseModel):
    """Формализованное состояние мастера (контракт оркестрации): сервер — источник истины."""

    current_step: WizardStepId
    can_proceed_to: list[WizardStepId] = Field(default_factory=list)
    next_required_action: WizardNextRequiredAction
    validation_errors: list[str] = Field(default_factory=list)


# Алиас для OpenAPI/аудита — то же поле, что и в теле ответа черновика (`wizard_state`).
CommercialWizardStateResponse = CommercialWizardState


class CommercialParseLine(BaseModel):
    index: int = Field(ge=0)
    text: str
    empty: bool = False
    ok: bool
    reason_text: str | None = None


COMMERCIAL_PARSE_TEXT_MAX_LENGTH = 50_000


class CommercialParseRequest(BaseModel):
    text: str = Field(min_length=1, max_length=COMMERCIAL_PARSE_TEXT_MAX_LENGTH)
    product_type: ProductType = "plates"
    lint_only: bool = False


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


class CommercialUnpricedPlateReplacement(BaseModel):
    load_code: int
    price: float


class CommercialUnpricedPlateLine(BaseModel):
    id: str = Field(min_length=1)
    name: str = ""
    line: str = ""
    qty: int = 1
    length_m: float = 0.0
    width_m: float = 0.0
    load_class: int = 0
    replacements: list[CommercialUnpricedPlateReplacement] = Field(default_factory=list)


class CommercialInvalidWidthReplacement(BaseModel):
    width_mm: int
    width_label: str
    price: float | None = None


class CommercialInvalidWidthLine(BaseModel):
    id: str = Field(min_length=1)
    name: str = ""
    line: str = ""
    qty: int = 1
    length_m: float = 0.0
    width_m: float = 0.0
    width_mm: int = 0
    load_class: int = 0
    replacements: list[CommercialInvalidWidthReplacement] = Field(default_factory=list)


class CommercialDoborPair(BaseModel):
    id: str = Field(min_length=1)
    source_line: str
    primary_line: str
    complement_line: str


class CommercialPlateBatch(BaseModel):
    source_type: CommercialSourceType
    original_text: str = ""
    normalized_text: str = ""
    ocr_text: str = ""
    filename: str = ""


class CommercialPileBatch(BaseModel):
    source_type: CommercialSourceType
    original_text: str = ""
    normalized_text: str = ""
    ocr_text: str = ""
    filename: str = ""


class CommercialStepBatch(BaseModel):
    source_type: CommercialSourceType
    original_text: str = ""
    normalized_text: str = ""
    ocr_text: str = ""
    filename: str = ""


class CommercialMarchBatch(BaseModel):
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


class CommercialAppendBatch(BaseModel):
    """Один цикл append: product_type + line_ids, добавленные в этом батче."""

    batch_id: str = Field(min_length=1)
    product_type: ProductType
    line_ids: list[str] = Field(default_factory=list)


class CommercialOrderLine(BaseModel):
    """Строка заказа КП; product-specific поля допускаются через extra."""

    model_config = ConfigDict(extra="allow")

    line_id: Annotated[str, Field(min_length=1)] | None = None
    product_type: ProductType | None = None
    append_batch_id: Annotated[str, Field(min_length=1)] | None = None


class CommercialDraftMetadata(BaseModel):
    """Метаданные черновика; owner_user_id хранится на сервере и не отдаётся клиенту."""

    owner_user_id: int | None = Field(default=None, exclude=True)
    product_type: ProductType = "plates"
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
    unpriced_plate_lines: list[CommercialUnpricedPlateLine] = Field(default_factory=list)
    invalid_width_lines: list[CommercialInvalidWidthLine] = Field(default_factory=list)
    dobor_pairs: list[CommercialDoborPair] = Field(default_factory=list)
    diagnostics: list[dict[str, Any]] = Field(default_factory=list)
    price_rows_count: int = 0
    breakdown_tables_count: int = 0
    total_sum: float = 0.0
    plate_batches: list[CommercialPlateBatch] = Field(default_factory=list)
    pile_batches: list[CommercialPileBatch] = Field(default_factory=list)
    step_batches: list[CommercialStepBatch] = Field(default_factory=list)
    march_batches: list[CommercialMarchBatch] = Field(default_factory=list)
    default_concrete_grade: str = "B25"
    wide_plates_resolved: bool = False
    unpriced_plates_resolved: bool = True
    invalid_widths_resolved: bool = True
    last_source_filename: str = ""
    ai_applied: bool = False
    last_ai_instruction: str = ""
    current_step: WizardStepCoerced = WizardStepId.plates
    current_save_mode: CommercialSaveMode | None = None
    execution_terms: str = ""
    logistics_cost: float = 0.0
    pile_logistics_cost: float = 0.0
    pile_trip_overrides: dict[str, int] = Field(default_factory=dict)
    ocr_recognition_mode: str = ""
    ocr_cost_usd: float = 0.0
    ocr_cost_rub: float = 0.0
    ocr_api_calls: int = 0
    ocr_method: str = ""
    ocr_verify_applied: bool = False
    ocr_verify_failed: bool = False
    ocr_verify_skipped_reason: str | None = None
    ocr_verify_applied_reason: str | None = None
    ocr_verify_select_reason: str | None = None
    ocr_preprocess: str | None = None
    ocr_corrections: list[dict[str, Any]] = Field(default_factory=list)
    ocr_row_count_on_image: int | None = None
    append_batches: list[CommercialAppendBatch] = Field(default_factory=list)
    resume_kp_id: int | None = Field(default=None, ge=1)


class CommercialBreakdownTable(BaseModel):
    name: str
    rows: list[list[str]] = Field(default_factory=list)


class CommercialDraftBreakdownResponse(BaseModel):
    draft_id: str
    items: list[CommercialBreakdownTable] = Field(default_factory=list)


class CommercialOcrPageResponse(BaseModel):
    normalized_text: str
    ocr_verify_failed: bool = False
    ocr_corrections: list[dict[str, Any]] = Field(default_factory=list)


class CommercialDraftDetailsResponse(BaseModel):
    draft_id: str
    order: dict[str, Any]
    optimization: dict[str, Any]
    order_data: list[CommercialOrderLine]
    metadata: CommercialDraftMetadata
    wizard_state: CommercialWizardState
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
    pile_logistics_cost: float | None = None
    pile_trip_overrides: dict[str, int] | None = None


class CommercialDraftLinePatchRequest(BaseModel):
    """Partial update of one draft order line: qty and/or source_text (mark-as-in-list)."""

    qty: int | None = None
    source_text: str | None = None


class CommercialRestoreLinesRequest(BaseModel):
    """Undo a row delete/replace by splicing snapshot lines at index."""

    index: int = Field(ge=0)
    lines: list[dict[str, Any]] = Field(min_length=1)
    replace_line_ids: list[str] = Field(default_factory=list)


class CommercialAppendStartRequest(BaseModel):
    """Start a new append cycle: switch product_type, clear cycle input, keep header."""

    product_type: ProductType


class CommercialWidePlateDecision(BaseModel):
    line_id: str | None = None
    source_line: str | None = None
    action: CommercialWidePlateAction
    replacement_text: str = ""


class CommercialWidePlatesResolveRequest(BaseModel):
    decisions: list[CommercialWidePlateDecision] = Field(min_length=1)


class CommercialUnpricedPlateDecision(BaseModel):
    line_id: str | None = None
    source_line: str | None = None
    action: CommercialUnpricedPlateAction
    load_code: int | None = None


class CommercialUnpricedPlatesResolveRequest(BaseModel):
    decisions: list[CommercialUnpricedPlateDecision] = Field(min_length=1)


class CommercialInvalidWidthDecision(BaseModel):
    line_id: str | None = None
    source_line: str | None = None
    action: CommercialInvalidWidthAction
    width_mm: int | None = None


class CommercialInvalidWidthsResolveRequest(BaseModel):
    decisions: list[CommercialInvalidWidthDecision] = Field(min_length=1)


class CommercialPileGradesUpdateRequest(BaseModel):
    concrete_grade: str = Field(min_length=2)


class CommercialMarchGradesUpdateRequest(BaseModel):
    concrete_grade: str = Field(min_length=2)


class CommercialBridgePileGradesUpdateRequest(BaseModel):
    concrete_grade: str = Field(min_length=2)



class CommercialFbsGradesUpdateRequest(BaseModel):
    concrete_grade: str = Field(min_length=2)


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

