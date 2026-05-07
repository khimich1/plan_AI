from __future__ import annotations

import logging

from fastapi import APIRouter, Depends, HTTPException, Query, status
from fastapi.responses import FileResponse, Response

from app.dependencies.auth import require_roles
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


logger = logging.getLogger(__name__)

router = APIRouter(prefix="/commercial/archive", tags=["commercial-archive"])


def get_archive_service() -> ArchiveService:
    return ArchiveService()


@router.get("", response_model=list[ArchiveOfferListItem])
def list_archive_offers(
    section: ArchiveSection = Query(default="archived"),
    _user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> list[ArchiveOfferListItem]:
    return service.list_offers(section)


@router.get("/search", response_model=ArchiveSearchResponse)
def search_offer_by_number(
    query: int = Query(..., ge=1, description="Номер КП"),
    _user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> ArchiveSearchResponse:
    offer = service.search_by_number(query)
    return ArchiveSearchResponse(found=offer is not None, offer=offer)


@router.get("/current-plan/gantt")
async def download_current_plan_gantt(
    _user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> FileResponse:
    try:
        path = await service.build_current_plan_gantt()
    except ArchiveValidationError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=str(exc)) from exc
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
    _user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> ArchiveOfferDetails:
    try:
        return service.get_details(kp_id)
    except ArchiveNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=str(exc)) from exc


@router.get("/{kp_id}/files/{kind}")
async def download_archive_document(
    kp_id: int,
    kind: ArchiveFileKind,
    _user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> FileResponse:
    try:
        path = await service.generate_document(kp_id, kind)
    except ArchiveNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=str(exc)) from exc
    except ArchiveValidationError as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc
    except Exception as exc:
        logger.exception("Ошибка генерации %s для КП %s", kind, kp_id)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Не удалось сгенерировать {kind.upper()}",
        ) from exc

    media_type = (
        "application/pdf"
        if kind == "pdf"
        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    return FileResponse(path=path, filename=path.name, media_type=media_type)


@router.patch("/{kp_id}/discount", response_model=ArchiveOfferDetails)
def update_archive_discount(
    kp_id: int,
    payload: UpdateDiscountRequest,
    _user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> ArchiveOfferDetails:
    try:
        return service.update_discount(kp_id, payload.discount)
    except ArchiveNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=str(exc)) from exc
    except ArchiveValidationError as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc


@router.patch("/{kp_id}/logistics-cost", response_model=ArchiveOfferDetails)
def update_archive_logistics_cost(
    kp_id: int,
    payload: UpdateLogisticsCostRequest,
    _user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> ArchiveOfferDetails:
    try:
        return service.update_logistics_cost(kp_id, payload.logistics_cost)
    except ArchiveNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=str(exc)) from exc
    except ArchiveValidationError as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc


@router.delete("/{kp_id}", status_code=status.HTTP_204_NO_CONTENT)
def delete_archive_offer(
    kp_id: int,
    _user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> Response:
    try:
        service.delete_offer(kp_id)
    except ArchiveNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=str(exc)) from exc
    return Response(status_code=status.HTTP_204_NO_CONTENT)


@router.post("/{kp_id}/move-to-production", response_model=ArchiveOfferDetails)
def move_archive_offer_to_production(
    kp_id: int,
    payload: MoveToProductionRequest,
    _user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> ArchiveOfferDetails:
    try:
        return service.move_to_production(kp_id, payload.execution_terms)
    except ArchiveNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=str(exc)) from exc
    except ArchiveValidationError as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)) from exc


@router.get("/{kp_id}/production-estimate")
def get_production_estimate(
    kp_id: int,
    _user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> dict:
    try:
        return service.estimate_production(kp_id)
    except ArchiveNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=str(exc)) from exc
