from __future__ import annotations

import logging

from fastapi import APIRouter, Depends, HTTPException, Query, status
from fastapi.responses import FileResponse, Response

from app.dependencies.auth import require_roles
from app.dependencies.plate_context import get_plate_order_context
from app.core.http_errors import (
    MSG_ARCHIVE_NOT_FOUND,
    MSG_VALIDATION,
    raise_bad_request_client_error,
    raise_not_found_client_error,
)
from app.schemas.archive import (
    ArchiveFileKind,
    ArchiveOfferDetails,
    ArchiveOfferListItem,
    ArchiveSearchResponse,
    ArchiveSection,
    MoveToProductionRequest,
    UpdateDiscountRequest,
    UpdateLogisticsCostRequest,
)
from app.services.archive_service import (
    ArchiveNotFoundError,
    ArchiveService,
    ArchiveValidationError,
)
from core.plate_order_context import PlateOrderContext


logger = logging.getLogger(__name__)

router = APIRouter(prefix="/commercial/archive", tags=["commercial-archive"])


def get_archive_service() -> ArchiveService:
    return ArchiveService()


@router.get("", response_model=list[ArchiveOfferListItem])
def list_archive_offers(
    section: ArchiveSection = Query(default="archived"),
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> list[ArchiveOfferListItem]:
    return service.list_offers(section, user=user)


@router.get("/search", response_model=ArchiveSearchResponse)
def search_archive_offers(
    kp_id: int | None = Query(default=None, ge=1, description="Номер КП"),
    customer: str | None = Query(default=None, max_length=128, description="Имя заказчика"),
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> ArchiveSearchResponse:
    if kp_id is not None:
        return service.search(user=user, kp_id=kp_id)

    if customer is None or not customer.strip():
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Укажите номер КП или имя заказчика.",
        )

    trimmed = customer.strip()
    if len(trimmed) < 2:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Имя заказчика должно содержать не менее 2 символов.",
        )

    return service.search(user=user, customer=trimmed)


@router.get("/current-plan/gantt")
async def download_current_plan_gantt(
    _user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> FileResponse:
    try:
        path = await service.build_current_plan_gantt()
    except ArchiveValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.download_current_plan_gantt",
            detail=MSG_VALIDATION,
        )
    except Exception as exc:
        logger.exception("Ошибка сборки сводного Gantt")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Не удалось собрать диаграмму Ганта",
        ) from exc
    return FileResponse(
        path=path,
        filename=path.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@router.get("/{kp_id}", response_model=ArchiveOfferDetails)
def get_archive_offer(
    kp_id: int,
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> ArchiveOfferDetails:
    try:
        return service.get_details(kp_id, user=user)
    except ArchiveNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.get_archive_offer",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )


@router.get("/{kp_id}/files/{kind}")
async def download_archive_document(
    kp_id: int,
    kind: ArchiveFileKind,
    plate_order_ctx: PlateOrderContext = Depends(get_plate_order_context),
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> FileResponse:
    try:
        path = await service.generate_document(
            kp_id,
            kind,
            user=user,
            plate_order_ctx=plate_order_ctx,
        )
    except ArchiveNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.get_archive_offer",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except ArchiveValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.download_archive_document",
            detail=MSG_VALIDATION,
        )
    except Exception as exc:
        logger.exception("Ошибка генерации %s для КП %s", kind, kp_id)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Не удалось сгенерировать {kind.upper()}",
        ) from exc

    media_type = (
        "application/pdf"
        if kind in {"pdf", "schema"}
        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    return FileResponse(path=path, filename=path.name, media_type=media_type)


@router.patch("/{kp_id}/discount", response_model=ArchiveOfferDetails)
def update_archive_discount(
    kp_id: int,
    payload: UpdateDiscountRequest,
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> ArchiveOfferDetails:
    try:
        return service.update_discount(kp_id, payload.discount, user=user)
    except ArchiveNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.get_archive_offer",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except ArchiveValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.download_archive_document",
            detail=MSG_VALIDATION,
        )


@router.patch("/{kp_id}/logistics-cost", response_model=ArchiveOfferDetails)
def update_archive_logistics_cost(
    kp_id: int,
    payload: UpdateLogisticsCostRequest,
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> ArchiveOfferDetails:
    try:
        return service.update_logistics_cost(kp_id, payload.logistics_cost, user=user)
    except ArchiveNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.get_archive_offer",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except ArchiveValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.download_archive_document",
            detail=MSG_VALIDATION,
        )


@router.delete("/{kp_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_archive_offer(
    kp_id: int,
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> Response:
    try:
        service.delete_offer(kp_id, user=user)
    except ArchiveNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.get_archive_offer",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    return Response(status_code=status.HTTP_204_NO_CONTENT)


@router.post("/{kp_id}/move-to-production", response_model=ArchiveOfferDetails)
def move_archive_offer_to_production(
    kp_id: int,
    payload: MoveToProductionRequest,
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> ArchiveOfferDetails:
    try:
        return service.move_to_production(kp_id, payload.execution_terms, user=user)
    except ArchiveNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.get_archive_offer",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except ArchiveValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.download_archive_document",
            detail=MSG_VALIDATION,
        )


@router.get("/{kp_id}/production-estimate")
def get_production_estimate(
    kp_id: int,
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> dict:
    try:
        return service.estimate_production(kp_id, user=user)
    except ArchiveNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.get_archive_offer",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
