from __future__ import annotations

from collections.abc import Awaitable, Callable
from typing import Any, NoReturn

from fastapi import APIRouter, Depends, File, Form, HTTPException, Query, UploadFile, status
from fastapi.responses import FileResponse

from app.concurrency.cpu_bound import run_cpu_bound
from app.core.http_errors import (
    raise_parse_client_error, raise_unexpected_server_error,
    raise_unpriced_plates_error, raise_validation_client_error,
)
from app.dependencies.auth import REQUIRE_ADMIN_OR_MANAGER
from app.dependencies.commercial_draft import check_draft_ownership, verify_draft_ownership
from app.dependencies.plate_context import get_plate_order_context
from app.dependencies.services import get_commercial_service, get_commercial_workflow_service
from app.schemas.commercial import (
    CommercialAppendStartRequest, CommercialBridgePileGradesUpdateRequest,
    CommercialCreateFromFormResponse, CommercialDraftBreakdownResponse,
    CommercialDraftDetailsResponse, CommercialDraftMetaUpdateRequest,
    CommercialFbsGradesUpdateRequest, CommercialGenerateFilesRequest,
    CommercialGenerateFilesResponse, CommercialMarchGradesUpdateRequest,
    CommercialParseRequest, CommercialPileGradesUpdateRequest,
    CommercialPreviewRequest, CommercialSaveDraftRequest, CommercialSaveOfferResponse,
    CommercialUnpricedPlatesResolveRequest, CommercialWidePlatesResolveRequest,
)
from app.services.commercial_draft_service import CommercialDraftService
from app.services.commercial_service import CommercialService
from app.services.commercial_upload_validation import (
    ensure_external_ocr_enabled, prepare_commercial_ocr_upload,
)
from app.services.commercial_workflow_service import CommercialWorkflowService
from core.exceptions import PlateParseError, UnpricedPlatesError
from core.plate_order_context import PlateOrderContext

router = APIRouter(prefix="/commercial", tags=["commercial"])


class _ProductUpdateForm:
    def __init__(
        self, mode: str = Form(default="append"), text: str = Form(default=""),
        image: UploadFile | None = File(default=None),
    ) -> None:
        self.mode, self.text, self.image = mode, text, image


class _ProductAiForm:
    def __init__(
        self, instruction: str = Form(...), image: UploadFile | None = File(default=None),
    ) -> None:
        self.instruction, self.image = instruction, image


def _details(result: dict[str, Any]) -> CommercialDraftDetailsResponse:
    return CommercialDraftDetailsResponse.model_validate(result)


def _raise_draft_http(
    exc: BaseException,
    *,
    where: str,
    not_found: bool = True,
    plate_parse: bool = False,
    unpriced: bool = False,
    validation: bool = True,
    not_found_detail: str = "Черновик не найден.",
) -> NoReturn:
    if not_found and isinstance(exc, FileNotFoundError):
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=not_found_detail) from exc
    if plate_parse and isinstance(exc, PlateParseError):
        raise_parse_client_error(exc, where=where)
    if unpriced and isinstance(exc, UnpricedPlatesError):
        raise_unpriced_plates_error(exc, where=where)
    if validation and isinstance(exc, ValueError):
        raise_validation_client_error(exc, where=where, detail=str(exc))
    raise_unexpected_server_error(exc, where=where)


def _sync_draft(where: str, call: Callable[[], Any], **flags: Any) -> Any:
    try:
        return call()
    except Exception as exc:
        _raise_draft_http(exc, where=where, **flags)


async def _run_product_update(
    draft_id: str, user: dict, body: _ProductUpdateForm, call: Callable[..., Awaitable[Any]],
    where: str, *, plate_order_ctx: PlateOrderContext | None = None, plate_parse: bool = False,
) -> CommercialDraftDetailsResponse:
    image_bytes, image_name = await prepare_commercial_ocr_upload(
        image=body.image, user_id=int(user["id"])
    )
    kwargs: dict[str, Any] = {
        "mode": body.mode, "text": body.text, "image_bytes": image_bytes, "image_filename": image_name,
    }
    if plate_order_ctx is not None:
        kwargs["plate_order_ctx"] = plate_order_ctx
    try:
        result = await call(draft_id, **kwargs)
    except Exception as exc:
        _raise_draft_http(exc, where=where, plate_parse=plate_parse)
    return _details(result)


async def _run_product_ai(
    draft_id: str, user: dict, body: _ProductAiForm, call: Callable[..., Awaitable[Any]],
    where: str, *, plate_order_ctx: PlateOrderContext | None = None, plate_parse: bool = False,
) -> CommercialDraftDetailsResponse:
    ensure_external_ocr_enabled()
    image_bytes, image_name = await prepare_commercial_ocr_upload(
        image=body.image, user_id=int(user["id"])
    )
    kwargs: dict[str, Any] = {
        "instruction": body.instruction, "image_bytes": image_bytes, "image_filename": image_name,
    }
    if plate_order_ctx is not None:
        kwargs["plate_order_ctx"] = plate_order_ctx
    try:
        result = await call(draft_id, **kwargs)
    except Exception as exc:
        _raise_draft_http(exc, where=where, plate_parse=plate_parse)
    return _details(result)


def _run_product_grades(
    draft_id: str, concrete_grade: str, call: Callable[..., Any], where: str,
) -> CommercialDraftDetailsResponse:
    try:
        result = call(draft_id, concrete_grade=concrete_grade)
    except Exception as exc:
        _raise_draft_http(exc, where=where)
    return _details(result)


@router.post("/parse")
def parse_commercial_text(
    payload: CommercialParseRequest,
    _user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
    service: CommercialService = Depends(get_commercial_service),
) -> dict:
    try:
        result = service.parse(payload.text)
    except Exception as exc:
        _raise_draft_http(
            exc, where="parse_commercial_text", plate_parse=True, validation=False, not_found=False
        )
    return {
        "order": result.order.to_dict(), "normalized_text": result.normalized_text,
        "normalized_lines": result.normalized_lines, "unparsed_lines": result.unparsed_lines,
        "warnings": result.warnings, "wide_plate_lines": result.wide_plate_lines,
        "dobor_pairs": CommercialDraftService.serialize_dobor_pairs(result.dobor_pairs),
        "diagnostics": result.diagnostics,
    }


@router.post("/drafts", response_model=CommercialDraftDetailsResponse)
async def create_commercial_draft(
    text: str = Form(default=""),
    image: UploadFile | None = File(default=None),
    product_type: str = Form(default="plates"),
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
    plate_order_ctx: PlateOrderContext = Depends(get_plate_order_context),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    image_bytes, image_name = await prepare_commercial_ocr_upload(image=image, user_id=int(user["id"]))
    try:
        result = await workflow.create_draft(
            text=text, image_bytes=image_bytes, image_filename=image_name,
            owner_user_id=int(user["id"]), plate_order_ctx=plate_order_ctx, product_type=product_type,
        )
    except Exception as exc:
        _raise_draft_http(exc, where="create_commercial_draft", plate_parse=True, not_found=False)
    return _details(result)


@router.patch("/drafts/{draft_id}/piles", response_model=CommercialDraftDetailsResponse)
async def update_commercial_draft_piles(
    body: _ProductUpdateForm = Depends(), draft_id: str = Depends(verify_draft_ownership),
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return await _run_product_update(draft_id, user, body, workflow.update_draft_piles, "update_commercial_draft_piles")

@router.post("/drafts/{draft_id}/piles/ai", response_model=CommercialDraftDetailsResponse)
async def apply_ai_piles_to_draft(
    body: _ProductAiForm = Depends(), draft_id: str = Depends(verify_draft_ownership),
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return await _run_product_ai(draft_id, user, body, workflow.apply_ai_piles_instruction, "apply_ai_piles_to_draft")

@router.patch("/drafts/{draft_id}/piles/grades", response_model=CommercialDraftDetailsResponse)
def update_draft_pile_grades(
    payload: CommercialPileGradesUpdateRequest, draft_id: str = Depends(verify_draft_ownership),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return _run_product_grades(draft_id, payload.concrete_grade, workflow.update_draft_pile_grades, "update_draft_pile_grades")

@router.patch("/drafts/{draft_id}/marches", response_model=CommercialDraftDetailsResponse)
async def update_commercial_draft_marches(
    body: _ProductUpdateForm = Depends(), draft_id: str = Depends(verify_draft_ownership),
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return await _run_product_update(draft_id, user, body, workflow.update_draft_marches, "update_commercial_draft_marches")

@router.post("/drafts/{draft_id}/marches/ai", response_model=CommercialDraftDetailsResponse)
async def apply_ai_marches_to_draft(
    body: _ProductAiForm = Depends(), draft_id: str = Depends(verify_draft_ownership),
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return await _run_product_ai(draft_id, user, body, workflow.apply_ai_marches_instruction, "apply_ai_marches_to_draft")

@router.patch("/drafts/{draft_id}/marches/grades", response_model=CommercialDraftDetailsResponse)
def update_draft_march_grades(
    payload: CommercialMarchGradesUpdateRequest, draft_id: str = Depends(verify_draft_ownership),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return _run_product_grades(draft_id, payload.concrete_grade, workflow.update_draft_march_grades, "update_draft_march_grades")

@router.patch("/drafts/{draft_id}/bridge-piles", response_model=CommercialDraftDetailsResponse)
async def update_commercial_draft_bridge_piles(
    body: _ProductUpdateForm = Depends(), draft_id: str = Depends(verify_draft_ownership),
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return await _run_product_update(draft_id, user, body, workflow.update_draft_bridge_piles, "update_commercial_draft_bridge_piles")

@router.post("/drafts/{draft_id}/bridge-piles/ai", response_model=CommercialDraftDetailsResponse)
async def apply_ai_bridge_piles_to_draft(
    body: _ProductAiForm = Depends(), draft_id: str = Depends(verify_draft_ownership),
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return await _run_product_ai(draft_id, user, body, workflow.apply_ai_bridge_piles_instruction, "apply_ai_bridge_piles_to_draft")

@router.patch("/drafts/{draft_id}/bridge-piles/grades", response_model=CommercialDraftDetailsResponse)
def update_draft_bridge_pile_grades(
    payload: CommercialBridgePileGradesUpdateRequest, draft_id: str = Depends(verify_draft_ownership),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return _run_product_grades(draft_id, payload.concrete_grade, workflow.update_draft_bridge_pile_grades, "update_draft_bridge_pile_grades")

@router.patch("/drafts/{draft_id}/fbs", response_model=CommercialDraftDetailsResponse)
async def update_commercial_draft_fbs(
    body: _ProductUpdateForm = Depends(), draft_id: str = Depends(verify_draft_ownership),
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return await _run_product_update(draft_id, user, body, workflow.update_draft_fbs, "update_commercial_draft_fbs")

@router.post("/drafts/{draft_id}/fbs/ai", response_model=CommercialDraftDetailsResponse)
async def apply_ai_fbs_to_draft(
    body: _ProductAiForm = Depends(), draft_id: str = Depends(verify_draft_ownership),
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return await _run_product_ai(draft_id, user, body, workflow.apply_ai_fbs_instruction, "apply_ai_fbs_to_draft")

@router.patch("/drafts/{draft_id}/fbs/grades", response_model=CommercialDraftDetailsResponse)
def update_draft_fbs_grades(
    payload: CommercialFbsGradesUpdateRequest, draft_id: str = Depends(verify_draft_ownership),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return _run_product_grades(draft_id, payload.concrete_grade, workflow.update_draft_fbs_grades, "update_draft_fbs_grades")

@router.patch("/drafts/{draft_id}/steps", response_model=CommercialDraftDetailsResponse)
async def update_commercial_draft_steps(
    body: _ProductUpdateForm = Depends(), draft_id: str = Depends(verify_draft_ownership),
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return await _run_product_update(draft_id, user, body, workflow.update_draft_steps, "update_commercial_draft_steps")

@router.post("/drafts/{draft_id}/steps/ai", response_model=CommercialDraftDetailsResponse)
async def apply_ai_steps_to_draft(
    body: _ProductAiForm = Depends(), draft_id: str = Depends(verify_draft_ownership),
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return await _run_product_ai(draft_id, user, body, workflow.apply_ai_steps_instruction, "apply_ai_steps_to_draft")

@router.patch("/drafts/{draft_id}/plates", response_model=CommercialDraftDetailsResponse)
async def update_commercial_draft_plates(
    body: _ProductUpdateForm = Depends(), draft_id: str = Depends(verify_draft_ownership),
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
    plate_order_ctx: PlateOrderContext = Depends(get_plate_order_context),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return await _run_product_update(
        draft_id, user, body, workflow.update_draft_plates, "update_commercial_draft_plates",
        plate_order_ctx=plate_order_ctx, plate_parse=True,
    )

@router.post("/drafts/{draft_id}/plates/ai", response_model=CommercialDraftDetailsResponse)
async def apply_ai_plates_to_draft(
    body: _ProductAiForm = Depends(), draft_id: str = Depends(verify_draft_ownership),
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
    plate_order_ctx: PlateOrderContext = Depends(get_plate_order_context),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return await _run_product_ai(
        draft_id, user, body, workflow.apply_ai_plates_instruction, "apply_ai_plates_to_draft",
        plate_order_ctx=plate_order_ctx, plate_parse=True,
    )

@router.post("/drafts/{draft_id}/wide-plates/resolve", response_model=CommercialDraftDetailsResponse)
def resolve_draft_wide_plates(
    payload: CommercialWidePlatesResolveRequest, draft_id: str = Depends(verify_draft_ownership),
    plate_order_ctx: PlateOrderContext = Depends(get_plate_order_context),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return _details(_sync_draft(
        "resolve_draft_wide_plates",
        lambda: workflow.resolve_wide_plates(
            draft_id, decisions=[item.model_dump() for item in payload.decisions],
            plate_order_ctx=plate_order_ctx,
        ),
        plate_parse=True,
    ))

@router.post("/drafts/{draft_id}/unpriced-plates/resolve", response_model=CommercialDraftDetailsResponse)
def resolve_draft_unpriced_plates(
    payload: CommercialUnpricedPlatesResolveRequest, draft_id: str = Depends(verify_draft_ownership),
    plate_order_ctx: PlateOrderContext = Depends(get_plate_order_context),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return _details(_sync_draft(
        "resolve_draft_unpriced_plates",
        lambda: workflow.resolve_unpriced_plates(
            draft_id, decisions=[item.model_dump() for item in payload.decisions],
            plate_order_ctx=plate_order_ctx,
        ),
        plate_parse=True,
    ))

@router.patch("/drafts/{draft_id}/meta", response_model=CommercialDraftDetailsResponse)
def update_draft_meta(
    payload: CommercialDraftMetaUpdateRequest, draft_id: str = Depends(verify_draft_ownership),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return _details(_sync_draft("update_draft_meta", lambda: workflow.update_draft_meta(
        draft_id, manager_id=payload.manager_id, client_name=payload.client_name,
        discount_percent=payload.discount_percent, conditions_mode=payload.conditions_mode,
        delivery_conditions=payload.delivery_conditions, payment_conditions=payload.payment_conditions,
        logistics_cost=payload.logistics_cost,
    )))

@router.post("/drafts/{draft_id}/calculate", response_model=CommercialDraftDetailsResponse)
async def calculate_draft(
    draft_id: str = Depends(verify_draft_ownership),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    try:
        result = await run_cpu_bound(lambda: workflow.calculate_draft(draft_id))
    except Exception as exc:
        _raise_draft_http(exc, where="calculate_draft", unpriced=True)
    return _details(result)

@router.post("/drafts/{draft_id}/append/start", response_model=CommercialDraftDetailsResponse)
def start_append_cycle(
    payload: CommercialAppendStartRequest, draft_id: str = Depends(verify_draft_ownership),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return _details(_sync_draft(
        "start_append_cycle",
        lambda: workflow.start_append_cycle(draft_id, product_type=payload.product_type),
    ))

@router.post("/drafts/{draft_id}/append/undo-last", response_model=CommercialDraftDetailsResponse)
def undo_last_append_batch(
    draft_id: str = Depends(verify_draft_ownership),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return _details(_sync_draft("undo_last_append_batch", lambda: workflow.undo_last_append_batch(draft_id)))

@router.delete("/drafts/{draft_id}/lines/{line_id}", response_model=CommercialDraftDetailsResponse)
def delete_draft_line(
    line_id: str, draft_id: str = Depends(verify_draft_ownership),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return _details(_sync_draft(
        "delete_draft_line", lambda: workflow.delete_order_line(draft_id, line_id),
        not_found_detail="Строка не найдена.",
    ))

@router.post("/generate-preview")
async def generate_preview(
    payload: CommercialPreviewRequest, user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
    plate_order_ctx: PlateOrderContext = Depends(get_plate_order_context),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> dict:
    try:
        return await run_cpu_bound(
            lambda: workflow.generate_and_persist_preview(
                text=payload.text, owner_user_id=int(user["id"]), plate_order_ctx=plate_order_ctx,
            ),
            plate_order_ctx=plate_order_ctx,
        )
    except Exception as exc:
        _raise_draft_http(exc, where="generate_preview", plate_parse=True, validation=False, not_found=False)

@router.post("/from-form", response_model=CommercialCreateFromFormResponse)
async def create_draft_from_form(
    text: str = Form(default=""), manager_id: int = Form(...), client_name: str = Form(...),
    discount_percent: float = Form(default=0.0), delivery_conditions: str = Form(default=""),
    payment_conditions: str = Form(default=""), image: UploadFile | None = File(default=None),
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
    plate_order_ctx: PlateOrderContext = Depends(get_plate_order_context),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialCreateFromFormResponse:
    image_bytes, image_name = await prepare_commercial_ocr_upload(image=image, user_id=int(user["id"]))
    try:
        result = await workflow.create_draft_from_form(
            text=text, image_bytes=image_bytes, image_filename=image_name, manager_id=manager_id,
            client_name=client_name, discount_percent=discount_percent,
            delivery_conditions=delivery_conditions, payment_conditions=payment_conditions,
            owner_user_id=int(user["id"]), plate_order_ctx=plate_order_ctx,
        )
    except Exception as exc:
        _raise_draft_http(exc, where="create_draft_from_form", plate_parse=True, not_found=False)
    return CommercialCreateFromFormResponse.model_validate(result)

@router.post("/drafts/{draft_id}/generate-files", response_model=CommercialGenerateFilesResponse)
def generate_draft_files(
    draft_id: str = Depends(verify_draft_ownership), payload: CommercialGenerateFilesRequest | None = None,
    plate_order_ctx: PlateOrderContext = Depends(get_plate_order_context),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialGenerateFilesResponse:
    files = _sync_draft(
        "generate_draft_files",
        lambda: workflow.generate_files(draft_id, payload.file_types if payload else None, plate_order_ctx=plate_order_ctx),
        unpriced=True,
    )
    return CommercialGenerateFilesResponse(draft_id=draft_id, files=files)

@router.post("/drafts/{draft_id}/save", response_model=CommercialSaveOfferResponse)
def save_draft_offer(
    payload: CommercialSaveDraftRequest, draft_id: str = Depends(verify_draft_ownership),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialSaveOfferResponse:
    result = _sync_draft(
        "save_draft_offer",
        lambda: workflow.save_draft(draft_id, mode=payload.mode, execution_terms_input=payload.execution_terms_input),
        unpriced=True,
    )
    return CommercialSaveOfferResponse(draft_id=draft_id, **result)

@router.get("/files/{filename}")
def download_generated_file(
    filename: str, draft_id: str = Query(..., min_length=1),
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> FileResponse:
    check_draft_ownership(draft_id, user)
    try:
        target_file = workflow.resolve_downloadable_file(draft_id, filename)
    except ValueError as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc
    except FileNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=str(exc)) from exc
    return FileResponse(path=target_file, filename=target_file.name)

@router.get("/drafts/{draft_id}", response_model=CommercialDraftDetailsResponse)
def get_preview_draft(
    draft_id: str = Depends(verify_draft_ownership),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftDetailsResponse:
    return _details(_sync_draft("get_preview_draft", lambda: workflow.get_draft_details(draft_id), validation=False))

@router.get("/drafts/{draft_id}/breakdown", response_model=CommercialDraftBreakdownResponse)
def get_draft_breakdown(
    draft_id: str = Depends(verify_draft_ownership),
    workflow: CommercialWorkflowService = Depends(get_commercial_workflow_service),
) -> CommercialDraftBreakdownResponse:
    result = _sync_draft("get_draft_breakdown", lambda: workflow.get_draft_breakdown(draft_id), validation=False)
    return CommercialDraftBreakdownResponse.model_validate(result)
