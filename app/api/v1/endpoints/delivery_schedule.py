from __future__ import annotations

from typing import Literal

from fastapi import APIRouter, Depends, File, Query, UploadFile
from fastapi.responses import FileResponse, Response

from app.core.http_errors import (
    MSG_NOT_FOUND,
    MSG_VALIDATION,
    raise_not_found_client_error,
    raise_unexpected_server_error,
    raise_unprocessable_client_error,
)
from app.dependencies.auth import require_roles
from app.dependencies.services import get_delivery_schedule_service
from app.schemas.delivery_schedule import (
    DeliverySchedulePut,
    DeliveryScheduleView,
    ImportDraftResponse,
)
from app.services.commercial_upload_validation import read_upload_file_capped
from app.services.delivery_schedule_service import (
    DeliveryScheduleNotFoundError,
    DeliveryScheduleService,
    DeliveryScheduleValidationError,
)

router = APIRouter(
    prefix="/commercial/archive/{kp_id}/delivery-schedule",
    tags=["delivery-schedule"],
)

_XLSX_MEDIA = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
_PDF_MEDIA = "application/pdf"
_TEMPLATE_FILENAME = "delivery_schedule_template.xlsx"


@router.get("", response_model=DeliveryScheduleView)
def get_delivery_schedule(
    kp_id: int,
    user: dict = Depends(require_roles("admin", "manager")),
    service: DeliveryScheduleService = Depends(get_delivery_schedule_service),
) -> DeliveryScheduleView:
    try:
        return service.get(kp_id, user=user)
    except DeliveryScheduleNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="delivery_schedule.get_delivery_schedule",
            detail=str(exc) or MSG_NOT_FOUND,
        )


@router.put("", response_model=DeliveryScheduleView)
def put_delivery_schedule(
    kp_id: int,
    payload: DeliverySchedulePut,
    user: dict = Depends(require_roles("admin", "manager")),
    service: DeliveryScheduleService = Depends(get_delivery_schedule_service),
) -> DeliveryScheduleView:
    try:
        return service.replace(kp_id, payload, user)
    except DeliveryScheduleNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="delivery_schedule.put_delivery_schedule",
            detail=str(exc) or MSG_NOT_FOUND,
        )
    except DeliveryScheduleValidationError as exc:
        raise_unprocessable_client_error(
            exc,
            where="delivery_schedule.put_delivery_schedule",
            detail=str(exc) or MSG_VALIDATION,
        )


@router.get("/template")
def download_delivery_schedule_template(
    kp_id: int,
    user: dict = Depends(require_roles("admin", "manager")),
    service: DeliveryScheduleService = Depends(get_delivery_schedule_service),
) -> Response:
    try:
        data = service.build_template_bytes(kp_id, user=user)
    except DeliveryScheduleNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="delivery_schedule.download_delivery_schedule_template",
            detail=str(exc) or MSG_NOT_FOUND,
        )
    except RuntimeError as exc:
        raise_unexpected_server_error(
            exc,
            where="delivery_schedule.download_delivery_schedule_template",
        )
    return Response(
        content=data,
        media_type=_XLSX_MEDIA,
        headers={
            "Content-Disposition": f'attachment; filename="{_TEMPLATE_FILENAME}"'
        },
    )


@router.post("/import", response_model=ImportDraftResponse)
async def import_delivery_schedule_template(
    kp_id: int,
    file: UploadFile = File(...),
    user: dict = Depends(require_roles("admin", "manager")),
    service: DeliveryScheduleService = Depends(get_delivery_schedule_service),
) -> ImportDraftResponse:
    file_bytes = await read_upload_file_capped(file)
    try:
        return service.import_draft(kp_id, file_bytes, user=user)
    except DeliveryScheduleNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="delivery_schedule.import_delivery_schedule_template",
            detail=str(exc) or MSG_NOT_FOUND,
        )
    except DeliveryScheduleValidationError as exc:
        raise_unprocessable_client_error(
            exc,
            where="delivery_schedule.import_delivery_schedule_template",
            detail=str(exc) or MSG_VALIDATION,
        )
    except RuntimeError as exc:
        raise_unexpected_server_error(
            exc,
            where="delivery_schedule.import_delivery_schedule_template",
        )


@router.get("/document")
def download_delivery_schedule_document(
    kp_id: int,
    fmt: Literal["xlsx", "pdf"] = Query(...),
    user: dict = Depends(require_roles("admin", "manager")),
    service: DeliveryScheduleService = Depends(get_delivery_schedule_service),
) -> FileResponse:
    try:
        path = service.generate_document(kp_id, fmt, user)
    except DeliveryScheduleNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="delivery_schedule.download_delivery_schedule_document",
            detail=str(exc) or MSG_NOT_FOUND,
        )
    except DeliveryScheduleValidationError as exc:
        raise_unprocessable_client_error(
            exc,
            where="delivery_schedule.download_delivery_schedule_document",
            detail=str(exc) or MSG_VALIDATION,
        )
    except RuntimeError as exc:
        raise_unexpected_server_error(
            exc,
            where="delivery_schedule.download_delivery_schedule_document",
        )

    media_type = _PDF_MEDIA if fmt == "pdf" else _XLSX_MEDIA
    return FileResponse(path=path, filename=path.name, media_type=media_type)
