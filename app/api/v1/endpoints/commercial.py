from __future__ import annotations

from pathlib import Path

from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile, status
from fastapi.responses import FileResponse

from app.dependencies.auth import require_roles
from app.schemas.commercial import (
    CommercialCreateFromFormResponse,
    CommercialDraftDetailsResponse,
    CommercialGenerateFilesRequest,
    CommercialGenerateFilesResponse,
    CommercialParseRequest,
    CommercialPreviewRequest,
    CommercialSaveOfferResponse,
)
from app.services.commercial_workflow_service import CommercialWorkflowService
from app.services.commercial_service import CommercialService
from app.services.draft_store import DraftStore
from core.exceptions import PlateParseError

router = APIRouter(prefix="/commercial", tags=["commercial"])


@router.post("/parse")
def parse_commercial_text(
    payload: CommercialParseRequest,
    _user: dict = Depends(require_roles("admin", "manager")),
) -> dict:
    service = CommercialService()
    try:
        result = service.parse(payload.text)
    except PlateParseError as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc
    return {
        "order": result.order.to_dict(),
        "normalized_text": result.normalized_text,
        "normalized_lines": result.normalized_lines,
        "unparsed_lines": result.unparsed_lines,
        "warnings": result.warnings,
        "wide_plate_lines": result.wide_plate_lines,
        "diagnostics": result.diagnostics,
    }


@router.post("/generate-preview")
def generate_preview(
    payload: CommercialPreviewRequest,
    _user: dict = Depends(require_roles("admin", "manager")),
) -> dict:
    service = CommercialService()
    try:
        preview = service.generate_preview(text=payload.text)
    except PlateParseError as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc
    draft_id = DraftStore().save_preview(
        order=preview.parse_result.order,
        optimization_context=preview.optimization_context,
        order_data=preview.order_data,
        metadata={
            "normalized_text": preview.parse_result.normalized_text,
            "warnings": preview.parse_result.warnings,
            "unparsed_lines": preview.parse_result.unparsed_lines,
            "normalized_lines": preview.parse_result.normalized_lines,
            "wide_plate_lines": preview.parse_result.wide_plate_lines,
            "diagnostics": preview.parse_result.diagnostics,
            "breakdown_tables": preview.breakdown_tables,
            "price_rows_count": len(preview.price_rows),
            "breakdown_tables_count": len(preview.breakdown_tables),
            "total_sum": preview.total_sum,
        },
    )
    return {
        "draft_id": draft_id,
        "order": preview.parse_result.order.to_dict(),
        "unparsed_lines": preview.parse_result.unparsed_lines,
        "warnings": preview.parse_result.warnings,
        "optimization": {
            "total_plates": preview.optimization_context.total_plates,
            "total_cost": preview.optimization_context.total_cost,
        },
        "order_data": preview.order_data,
        "price_rows_count": len(preview.price_rows),
        "breakdown_tables_count": len(preview.breakdown_tables),
        "total_sum": preview.total_sum,
    }


@router.post("/from-form", response_model=CommercialCreateFromFormResponse)
async def create_draft_from_form(
    text: str = Form(default=""),
    manager_id: int = Form(...),
    client_name: str = Form(...),
    discount_percent: float = Form(default=0.0),
    delivery_conditions: str = Form(default=""),
    payment_conditions: str = Form(default=""),
    image: UploadFile | None = File(default=None),
    _user: dict = Depends(require_roles("admin", "manager")),
) -> CommercialCreateFromFormResponse:
    if image and image.content_type and not image.content_type.startswith("image/"):
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Поддерживаются только изображения.")

    image_bytes = await image.read() if image else None
    workflow = CommercialWorkflowService()
    try:
        result = await workflow.create_draft_from_form(
            text=text,
            image_bytes=image_bytes,
            image_filename=image.filename if image else None,
            manager_id=manager_id,
            client_name=client_name,
            discount_percent=discount_percent,
            delivery_conditions=delivery_conditions,
            payment_conditions=payment_conditions,
        )
    except PlateParseError as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc
    except ValueError as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc
    return CommercialCreateFromFormResponse.model_validate(result)


@router.post("/drafts/{draft_id}/generate-files", response_model=CommercialGenerateFilesResponse)
def generate_draft_files(
    draft_id: str,
    payload: CommercialGenerateFilesRequest | None = None,
    _user: dict = Depends(require_roles("admin", "manager")),
) -> CommercialGenerateFilesResponse:
    workflow = CommercialWorkflowService()
    try:
        files = workflow.generate_files(draft_id, payload.file_types if payload else None)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Черновик не найден.") from exc
    except ValueError as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc
    return CommercialGenerateFilesResponse(draft_id=draft_id, files=files)


@router.post("/drafts/{draft_id}/save", response_model=CommercialSaveOfferResponse)
def save_draft_offer(
    draft_id: str,
    _user: dict = Depends(require_roles("admin", "manager")),
) -> CommercialSaveOfferResponse:
    workflow = CommercialWorkflowService()
    try:
        result = workflow.save_offer(draft_id)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Черновик не найден.") from exc
    return CommercialSaveOfferResponse(draft_id=draft_id, **result)


@router.get("/files/{filename}")
def download_generated_file(
    filename: str,
    _user: dict = Depends(require_roles("admin", "manager")),
) -> FileResponse:
    safe_name = Path(filename).name
    if safe_name != filename:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Некорректное имя файла.")

    workflow = CommercialWorkflowService()
    target_file = workflow._resolve_generated_file(safe_name).resolve()
    outputs_dir = Path(workflow.settings.outputs_dir).resolve()
    if target_file.parent != outputs_dir or not target_file.exists():
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Файл не найден.")
    return FileResponse(path=target_file, filename=safe_name)


@router.get("/drafts/{draft_id}", response_model=CommercialDraftDetailsResponse)
def get_preview_draft(
    draft_id: str,
    _user: dict = Depends(require_roles("admin", "manager")),
) -> CommercialDraftDetailsResponse:
    workflow = CommercialWorkflowService()
    try:
        result = workflow.get_draft_details(draft_id)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Черновик не найден.") from exc
    return CommercialDraftDetailsResponse.model_validate(result)

