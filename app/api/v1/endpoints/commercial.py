from __future__ import annotations

from pathlib import Path

from fastapi import APIRouter, Depends, File, Form, HTTPException, Query, UploadFile, status
from fastapi.responses import FileResponse

from app.dependencies.auth import REQUIRE_ADMIN_OR_MANAGER
from app.dependencies.commercial_draft import check_draft_ownership, verify_draft_ownership
from app.schemas.commercial import (
    CommercialCreateFromFormResponse,
    CommercialDraftDetailsResponse,
    CommercialDraftMetaUpdateRequest,
    CommercialGenerateFilesRequest,
    CommercialGenerateFilesResponse,
    CommercialParseRequest,
    CommercialPreviewRequest,
    CommercialSaveDraftRequest,
    CommercialSaveOfferResponse,
    CommercialWidePlatesResolveRequest,
)
from app.core.http_errors import (
    raise_parse_client_error,
    raise_unexpected_server_error,
    raise_validation_client_error,
)
from app.services.commercial_workflow_service import CommercialWorkflowService
from app.services.commercial_service import CommercialService
from app.services.commercial_upload_validation import prepare_commercial_ocr_upload
from app.services.draft_store import DraftStore
from core.exceptions import PlateParseError

router = APIRouter(prefix="/commercial", tags=["commercial"])


@router.post("/parse")
def parse_commercial_text(
    payload: CommercialParseRequest,
    _user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
) -> dict:
    service = CommercialService()
    try:
        result = service.parse(payload.text)
    except PlateParseError as exc:
        raise_parse_client_error(exc, where="parse_commercial_text")
    except Exception as exc:
        raise_unexpected_server_error(exc, where="parse_commercial_text")
    return {
        "order": result.order.to_dict(),
        "normalized_text": result.normalized_text,
        "normalized_lines": result.normalized_lines,
        "unparsed_lines": result.unparsed_lines,
        "warnings": result.warnings,
        "wide_plate_lines": result.wide_plate_lines,
        "diagnostics": result.diagnostics,
    }


@router.post("/drafts", response_model=CommercialDraftDetailsResponse)
async def create_commercial_draft(
    text: str = Form(default=""),
    image: UploadFile | None = File(default=None),
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
) -> CommercialDraftDetailsResponse:
    image_bytes, image_name = await prepare_commercial_ocr_upload(
        image=image,
        user_id=int(user["id"]),
    )

    workflow = CommercialWorkflowService()
    try:
        result = await workflow.create_draft(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_name,
            owner_user_id=int(user["id"]),
        )
    except PlateParseError as exc:
        raise_parse_client_error(exc, where="create_commercial_draft")
    except ValueError as exc:
        raise_validation_client_error(exc, where="create_commercial_draft")
    except Exception as exc:
        raise_unexpected_server_error(exc, where="create_commercial_draft")
    return CommercialDraftDetailsResponse.model_validate(result)


@router.patch("/drafts/{draft_id}/plates", response_model=CommercialDraftDetailsResponse)
async def update_commercial_draft_plates(
    draft_id: str = Depends(verify_draft_ownership),
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
    mode: str = Form(default="append"),
    text: str = Form(default=""),
    image: UploadFile | None = File(default=None),
) -> CommercialDraftDetailsResponse:
    image_bytes, image_name = await prepare_commercial_ocr_upload(
        image=image,
        user_id=int(user["id"]),
    )

    workflow = CommercialWorkflowService()
    try:
        result = await workflow.update_draft_plates(
            draft_id,
            mode=mode,
            text=text,
            image_bytes=image_bytes,
            image_filename=image_name,
        )
    except FileNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Черновик не найден.") from exc
    except PlateParseError as exc:
        raise_parse_client_error(exc, where="update_commercial_draft_plates")
    except ValueError as exc:
        raise_validation_client_error(exc, where="update_commercial_draft_plates")
    except Exception as exc:
        raise_unexpected_server_error(exc, where="update_commercial_draft_plates")
    return CommercialDraftDetailsResponse.model_validate(result)


@router.post("/drafts/{draft_id}/wide-plates/resolve", response_model=CommercialDraftDetailsResponse)
def resolve_draft_wide_plates(
    payload: CommercialWidePlatesResolveRequest,
    draft_id: str = Depends(verify_draft_ownership),
) -> CommercialDraftDetailsResponse:
    workflow = CommercialWorkflowService()
    try:
        result = workflow.resolve_wide_plates(
            draft_id,
            decisions=[item.model_dump() for item in payload.decisions],
        )
    except FileNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Черновик не найден.") from exc
    except PlateParseError as exc:
        raise_parse_client_error(exc, where="resolve_draft_wide_plates")
    except ValueError as exc:
        raise_validation_client_error(exc, where="resolve_draft_wide_plates")
    except Exception as exc:
        raise_unexpected_server_error(exc, where="resolve_draft_wide_plates")
    return CommercialDraftDetailsResponse.model_validate(result)


@router.patch("/drafts/{draft_id}/meta", response_model=CommercialDraftDetailsResponse)
def update_draft_meta(
    payload: CommercialDraftMetaUpdateRequest,
    draft_id: str = Depends(verify_draft_ownership),
) -> CommercialDraftDetailsResponse:
    workflow = CommercialWorkflowService()
    try:
        result = workflow.update_draft_meta(
            draft_id,
            manager_id=payload.manager_id,
            client_name=payload.client_name,
            discount_percent=payload.discount_percent,
            conditions_mode=payload.conditions_mode,
            delivery_conditions=payload.delivery_conditions,
            payment_conditions=payload.payment_conditions,
            logistics_cost=payload.logistics_cost,
        )
    except FileNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Черновик не найден.") from exc
    except ValueError as exc:
        raise_validation_client_error(exc, where="update_draft_meta")
    except Exception as exc:
        raise_unexpected_server_error(exc, where="update_draft_meta")
    return CommercialDraftDetailsResponse.model_validate(result)


@router.post("/drafts/{draft_id}/calculate", response_model=CommercialDraftDetailsResponse)
def calculate_draft(
    draft_id: str = Depends(verify_draft_ownership),
) -> CommercialDraftDetailsResponse:
    workflow = CommercialWorkflowService()
    try:
        result = workflow.calculate_draft(draft_id)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Черновик не найден.") from exc
    except ValueError as exc:
        raise_validation_client_error(exc, where="calculate_draft")
    except Exception as exc:
        raise_unexpected_server_error(exc, where="calculate_draft")
    return CommercialDraftDetailsResponse.model_validate(result)


@router.post("/generate-preview")
def generate_preview(
    payload: CommercialPreviewRequest,
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
) -> dict:
    service = CommercialService()
    try:
        preview = service.generate_preview(text=payload.text)
        draft_id = DraftStore().save_preview(
            order=preview.parse_result.order,
            optimization_context=preview.optimization_context,
            order_data=preview.order_data,
            metadata={
                "owner_user_id": int(user["id"]),
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
    except PlateParseError as exc:
        raise_parse_client_error(exc, where="generate_preview")
    except Exception as exc:
        raise_unexpected_server_error(exc, where="generate_preview")
    return {
        "draft_id": draft_id,
        "order": preview.parse_result.order.to_dict(),
        "unparsed_lines": preview.parse_result.unparsed_lines,
        "warnings": preview.parse_result.warnings,
        "optimization": {
            "total_plates": preview.optimization_context.total_plates,
            "total_cost": preview.optimization_context.total_cost,
            "status": preview.optimization_context.optimization_status,
            "success": preview.optimization_context.optimization_success,
            "error_code": preview.optimization_context.optimization_error_code,
            "error_message": preview.optimization_context.optimization_error_message,
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
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
) -> CommercialCreateFromFormResponse:
    image_bytes, image_name = await prepare_commercial_ocr_upload(
        image=image,
        user_id=int(user["id"]),
    )

    workflow = CommercialWorkflowService()
    try:
        result = await workflow.create_draft_from_form(
            text=text,
            image_bytes=image_bytes,
            image_filename=image_name,
            manager_id=manager_id,
            client_name=client_name,
            discount_percent=discount_percent,
            delivery_conditions=delivery_conditions,
            payment_conditions=payment_conditions,
            owner_user_id=int(user["id"]),
        )
    except PlateParseError as exc:
        raise_parse_client_error(exc, where="create_draft_from_form")
    except ValueError as exc:
        raise_validation_client_error(exc, where="create_draft_from_form")
    except Exception as exc:
        raise_unexpected_server_error(exc, where="create_draft_from_form")
    return CommercialCreateFromFormResponse.model_validate(result)


@router.post("/drafts/{draft_id}/generate-files", response_model=CommercialGenerateFilesResponse)
def generate_draft_files(
    draft_id: str = Depends(verify_draft_ownership),
    payload: CommercialGenerateFilesRequest | None = None,
) -> CommercialGenerateFilesResponse:
    workflow = CommercialWorkflowService()
    try:
        files = workflow.generate_files(draft_id, payload.file_types if payload else None)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Черновик не найден.") from exc
    except ValueError as exc:
        raise_validation_client_error(exc, where="generate_draft_files")
    except Exception as exc:
        raise_unexpected_server_error(exc, where="generate_draft_files")
    return CommercialGenerateFilesResponse(draft_id=draft_id, files=files)


@router.post("/drafts/{draft_id}/save", response_model=CommercialSaveOfferResponse)
def save_draft_offer(
    payload: CommercialSaveDraftRequest,
    draft_id: str = Depends(verify_draft_ownership),
) -> CommercialSaveOfferResponse:
    workflow = CommercialWorkflowService()
    try:
        result = workflow.save_draft(
            draft_id,
            mode=payload.mode,
            execution_terms_input=payload.execution_terms_input,
        )
    except FileNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Черновик не найден.") from exc
    except ValueError as exc:
        raise_validation_client_error(exc, where="save_draft_offer")
    except Exception as exc:
        raise_unexpected_server_error(exc, where="save_draft_offer")
    return CommercialSaveOfferResponse(draft_id=draft_id, **result)


@router.get("/files/{filename}")
def download_generated_file(
    filename: str,
    draft_id: str = Query(..., min_length=1),
    user: dict = Depends(REQUIRE_ADMIN_OR_MANAGER),
) -> FileResponse:
    check_draft_ownership(draft_id, user)
    safe_name = Path(filename).name
    if safe_name != filename:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Некорректное имя файла.")

    store = DraftStore()
    if safe_name not in store.generated_files_filenames(draft_id):
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Файл не найден.")

    workflow = CommercialWorkflowService()
    target_file = workflow.get_or_generate_file(safe_name).resolve()
    outputs_dir = Path(workflow.settings.outputs_dir).resolve()
    if target_file.parent != outputs_dir or not target_file.exists():
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Файл не найден.")
    return FileResponse(path=target_file, filename=safe_name)


@router.get("/drafts/{draft_id}", response_model=CommercialDraftDetailsResponse)
def get_preview_draft(
    draft_id: str = Depends(verify_draft_ownership),
) -> CommercialDraftDetailsResponse:
    workflow = CommercialWorkflowService()
    try:
        result = workflow.get_draft_details(draft_id)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Черновик не найден.") from exc
    except Exception as exc:
        raise_unexpected_server_error(exc, where="get_preview_draft")
    return CommercialDraftDetailsResponse.model_validate(result)

