from __future__ import annotations

from fastapi import APIRouter, Depends, HTTPException, Query
from fastapi.responses import Response

from app.dependencies.auth import require_roles
from app.dependencies.services import get_offers_service
from app.schemas.offers import (
    CreateOfferRequest,
    CreateOfferResponse,
    DeleteOfferResponse,
    MoveOfferToProductionRequest,
    MoveToProductionResponse,
    OfferDetails,
    OfferListResponse,
    UpdateOfferDiscountRequest,
)
from app.services.offers_service import OffersService

router = APIRouter(prefix="/offers", tags=["offers"])


@router.get("", response_model=OfferListResponse)
def list_offers(
    status: str = Query(default="all"),
    limit: int = Query(default=200, ge=1, le=1000),
    kp_id: int | None = Query(default=None, ge=1),
    user: dict = Depends(require_roles("admin", "manager")),
    service: OffersService = Depends(get_offers_service),
) -> OfferListResponse:
    items = service.list_offers(status=status, limit=limit, kp_id=kp_id, user=user)
    return OfferListResponse.model_validate({"items": items, "count": len(items)})


@router.get("/{kp_id}", response_model=OfferDetails)
def get_offer(
    kp_id: int,
    user: dict = Depends(require_roles("admin", "manager")),
    service: OffersService = Depends(get_offers_service),
) -> OfferDetails:
    item = service.get_offer(kp_id, user=user)
    if not item:
        raise HTTPException(status_code=404, detail="Offer not found")
    return OfferDetails.model_validate(item)


@router.post("", response_model=CreateOfferResponse)
def create_offer(
    payload: CreateOfferRequest,
    user: dict = Depends(require_roles("admin", "manager")),
    service: OffersService = Depends(get_offers_service),
) -> CreateOfferResponse:
    return CreateOfferResponse.model_validate(service.create_offer(payload, user=user))


@router.patch("/{kp_id}/discount", response_model=OfferDetails)
def update_discount(
    kp_id: int,
    payload: UpdateOfferDiscountRequest,
    user: dict = Depends(require_roles("admin", "manager")),
    service: OffersService = Depends(get_offers_service),
) -> OfferDetails:
    item = service.update_discount(kp_id, payload.discount_percent, user=user)
    if not item:
        raise HTTPException(status_code=404, detail="Offer not found or discount was not updated")
    return OfferDetails.model_validate(item)


@router.patch("/{kp_id}/move-to-production", response_model=MoveToProductionResponse)
def move_to_production(
    kp_id: int,
    payload: MoveOfferToProductionRequest,
    user: dict = Depends(require_roles("admin", "manager")),
    service: OffersService = Depends(get_offers_service),
) -> MoveToProductionResponse:
    try:
        result = service.move_to_production(kp_id, payload.execution_terms_input, user=user)
    except ValueError as exc:
        if str(exc) == "not_found":
            raise HTTPException(status_code=404, detail="Offer not found") from exc
        if str(exc) == "invalid_status":
            raise HTTPException(status_code=400, detail="Only archived offers can be moved to production") from exc
        if str(exc) == "move_to_production_failed":
            raise HTTPException(
                status_code=500,
                detail="Failed to move offer to production",
            ) from exc
        raise HTTPException(status_code=400, detail="Failed to move offer to production") from exc
    return MoveToProductionResponse.model_validate(result)


@router.delete("/{kp_id}", response_model=DeleteOfferResponse)
def delete_offer(
    kp_id: int,
    user: dict = Depends(require_roles("admin", "manager")),
    service: OffersService = Depends(get_offers_service),
) -> DeleteOfferResponse:
    deleted = service.delete_offer(kp_id, user=user)
    if not deleted:
        raise HTTPException(status_code=404, detail="Offer not found")
    return DeleteOfferResponse(ok=True, kp_id=kp_id)


@router.get("/{kp_id}/pdf")
def download_offer_pdf(
    kp_id: int,
    user: dict = Depends(require_roles("admin", "manager")),
    service: OffersService = Depends(get_offers_service),
) -> Response:
    try:
        filename, data = service.generate_pdf(kp_id, user=user)
    except ValueError as exc:
        if str(exc) == "not_found":
            raise HTTPException(status_code=404, detail="Offer not found") from exc
        raise HTTPException(status_code=400, detail="Failed to generate PDF") from exc
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return Response(content=data, media_type="application/pdf", headers=headers)


@router.get("/{kp_id}/xlsx")
def download_offer_xlsx(
    kp_id: int,
    user: dict = Depends(require_roles("admin", "manager")),
    service: OffersService = Depends(get_offers_service),
) -> Response:
    try:
        filename, data = service.generate_xlsx(kp_id, user=user)
    except ValueError as exc:
        if str(exc) == "not_found":
            raise HTTPException(status_code=404, detail="Offer not found") from exc
        raise HTTPException(status_code=400, detail="Failed to generate XLSX") from exc
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )
