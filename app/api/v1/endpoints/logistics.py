from __future__ import annotations

import logging
from typing import Any, Literal, NoReturn, Optional

from fastapi import APIRouter, Depends, HTTPException, Query, Response, status
from openpyxl.utils.exceptions import InvalidFileException

from app.core.http_errors import raise_structured_error
from app.dependencies.auth import REQUIRE_LOGISTICS
from app.dependencies.services import (
    get_carrier_service,
    get_shipment_service,
)
from app.schemas.logistics import (
    CarrierListResponse,
    CarrierMergeRequest,
    CarrierMergeResponse,
    LogisticsKpSearchResponse,
    PileCatalogResponse,
    ShipmentCard,
    ShipmentCreateRequest,
    ShipmentItemsPutRequest,
    ShipmentListResponse,
    ShipmentMutationResponse,
    ShipmentPatchRequest,
    ShipmentProposeResponse,
)
from app.services.carrier_service import CarrierError, CarrierService
from app.services.shipment_service import ShipmentError, ShipmentService

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/logistics", tags=["logistics"])

_ERROR_4XX: dict[int, dict] = {
    403: {"description": "Недостаточно прав"},
    404: {"description": "Не найдено"},
    422: {"description": "Ошибка валидации доменной логики"},
}


def _actor(user: dict) -> Optional[str]:
    return (
        user.get("username")
        or user.get("email")
        or user.get("login")
        or user.get("user_id")
    )


def _raise_domain_error(
    exc: ShipmentError | CarrierError, *, where: str
) -> NoReturn:
    status = 404 if exc.code.endswith("_not_found") else 422
    raise_structured_error(
        status_code=status,
        code=exc.code,
        message=str(exc),
        details={},
        where=where,
    )


@router.get(
    "/shipments",
    response_model=ShipmentListResponse,
    summary="Список отгрузок (фильтры)",
    responses=_ERROR_4XX,
)
def list_shipments(
    date_from: Optional[str] = Query(None, pattern=r"^\d{4}-\d{2}-\d{2}$"),
    date_to: Optional[str] = Query(None, pattern=r"^\d{4}-\d{2}-\d{2}$"),
    kp_id: Optional[int] = Query(None, ge=1),
    carrier_id: Optional[int] = Query(None, ge=1),
    delivery_type: Optional[Literal["delivery", "pickup"]] = None,
    status: Optional[Literal["in_work", "done"]] = None,
    no_upd: bool = Query(False),
    attention: bool = Query(False),
    limit: int = Query(200, ge=1, le=1000),
    offset: int = Query(0, ge=0),
    _user: dict = Depends(REQUIRE_LOGISTICS),
    service: ShipmentService = Depends(get_shipment_service),
) -> ShipmentListResponse:
    try:
        return service.list_shipments(
            date_from=date_from,
            date_to=date_to,
            kp_id=kp_id,
            carrier_id=carrier_id,
            delivery_type=delivery_type,
            status=status,
            no_upd=no_upd,
            attention=attention,
            limit=limit,
            offset=offset,
        )
    except ShipmentError as exc:
        _raise_domain_error(exc, where="logistics.list_shipments")


@router.post(
    "/shipments",
    response_model=ShipmentCard,
    summary="Создать отгрузку",
    responses=_ERROR_4XX,
)
def create_shipment(
    payload: ShipmentCreateRequest,
    user: dict = Depends(REQUIRE_LOGISTICS),
    service: ShipmentService = Depends(get_shipment_service),
) -> ShipmentCard:
    try:
        return service.create(
            shipment_date=payload.shipment_date,
            delivery_type=payload.delivery_type,
            kp_ids=payload.kp_ids,
            actor=_actor(user),
        )
    except ShipmentError as exc:
        _raise_domain_error(exc, where="logistics.create_shipment")


@router.post(
    "/shipments/{source_id}/reuse-transport",
    response_model=ShipmentCard,
    summary="Создать рейс с копией транспорта из другого рейса",
    responses=_ERROR_4XX,
)
def reuse_transport(
    source_id: int,
    payload: ShipmentCreateRequest,
    user: dict = Depends(REQUIRE_LOGISTICS),
    service: ShipmentService = Depends(get_shipment_service),
) -> ShipmentCard:
    try:
        return service.reuse_transport(
            source_id,
            shipment_date=payload.shipment_date,
            delivery_type=payload.delivery_type,
            kp_ids=payload.kp_ids,
            actor=_actor(user),
        )
    except ShipmentError as exc:
        _raise_domain_error(exc, where="logistics.reuse_transport")


@router.get(
    "/shipments/{shipment_id}",
    response_model=ShipmentCard,
    summary="Карточка отгрузки",
    responses=_ERROR_4XX,
)
def get_shipment(
    shipment_id: int,
    _user: dict = Depends(REQUIRE_LOGISTICS),
    service: ShipmentService = Depends(get_shipment_service),
) -> ShipmentCard:
    try:
        return service.get(shipment_id)
    except ShipmentError as exc:
        _raise_domain_error(exc, where="logistics.get_shipment")


@router.patch(
    "/shipments/{shipment_id}",
    response_model=ShipmentCard,
    summary="Обновить шапку/заказы отгрузки",
    responses=_ERROR_4XX,
)
def patch_shipment(
    shipment_id: int,
    payload: ShipmentPatchRequest,
    user: dict = Depends(REQUIRE_LOGISTICS),
    service: ShipmentService = Depends(get_shipment_service),
) -> ShipmentCard:
    fields: dict[str, Any] = payload.model_dump(exclude_unset=True, exclude={"orders"})
    if "attention" in fields:
        fields["attention"] = 1 if fields["attention"] else 0
    orders = payload.orders if "orders" in payload.model_fields_set else None
    try:
        return service.patch(
            shipment_id,
            fields=fields,
            orders=orders,
            actor=_actor(user),
        )
    except ShipmentError as exc:
        _raise_domain_error(exc, where="logistics.patch_shipment")


@router.post(
    "/shipments/{shipment_id}/propose",
    response_model=ShipmentProposeResponse,
    summary="Автонабор рейса по FIFO с лимитом тоннажа",
    responses=_ERROR_4XX,
)
def propose_shipment(
    shipment_id: int,
    vehicle_class: Optional[str] = Query(None),
    _user: dict = Depends(REQUIRE_LOGISTICS),
    service: ShipmentService = Depends(get_shipment_service),
) -> ShipmentProposeResponse:
    try:
        return service.propose(shipment_id, vehicle_class=vehicle_class)
    except ShipmentError as exc:
        _raise_domain_error(exc, where="logistics.propose_shipment")


@router.put(
    "/shipments/{shipment_id}/items",
    response_model=ShipmentCard,
    summary="Подтвердить состав (полная замена строк)",
    responses=_ERROR_4XX,
)
def put_shipment_items(
    shipment_id: int,
    payload: ShipmentItemsPutRequest,
    user: dict = Depends(REQUIRE_LOGISTICS),
    service: ShipmentService = Depends(get_shipment_service),
) -> ShipmentCard:
    try:
        return service.put_items(shipment_id, payload.items, actor=_actor(user))
    except ShipmentError as exc:
        _raise_domain_error(exc, where="logistics.put_shipment_items")


@router.post(
    "/shipments/{shipment_id}/complete",
    response_model=ShipmentMutationResponse,
    summary="Отгрузить (списание СГП, аудит, DONE-проверка КП)",
    responses=_ERROR_4XX,
)
def complete_shipment(
    shipment_id: int,
    user: dict = Depends(REQUIRE_LOGISTICS),
    service: ShipmentService = Depends(get_shipment_service),
) -> ShipmentMutationResponse:
    try:
        return service.complete(shipment_id, actor=_actor(user))
    except ShipmentError as exc:
        _raise_domain_error(exc, where="logistics.complete_shipment")


@router.post(
    "/shipments/{shipment_id}/cancel",
    response_model=ShipmentMutationResponse,
    summary="Отменить рейс (резерв снимается, без аудита)",
    responses=_ERROR_4XX,
)
def cancel_shipment(
    shipment_id: int,
    user: dict = Depends(REQUIRE_LOGISTICS),
    service: ShipmentService = Depends(get_shipment_service),
) -> ShipmentMutationResponse:
    try:
        return service.cancel(shipment_id, actor=_actor(user))
    except ShipmentError as exc:
        _raise_domain_error(exc, where="logistics.cancel_shipment")


@router.get(
    "/shipments/{shipment_id}/sheet.xlsx",
    summary="Лист отгрузки (XLSX)",
    responses=_ERROR_4XX,
)
def download_shipment_sheet(
    shipment_id: int,
    _user: dict = Depends(REQUIRE_LOGISTICS),
    service: ShipmentService = Depends(get_shipment_service),
) -> Response:
    try:
        content = service.export_shipment_sheet_xlsx(shipment_id)
    except ShipmentError as exc:
        _raise_domain_error(exc, where="logistics.download_shipment_sheet")
    except InvalidFileException:
        raise_structured_error(
            status_code=500,
            code="shipment_sheet_export_failed",
            message="Ошибка формирования листа отгрузки",
            details={"shipment_id": shipment_id},
            where="logistics.download_shipment_sheet",
        )
    logger.info(
        f"Сформирован лист отгрузки для рейса #{shipment_id} ({len(content)} байт)"
    )
    return Response(
        content=content,
        media_type=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition": (
                f'attachment; filename="shipment_{shipment_id}_sheet.xlsx"'
            )
        },
    )


@router.get(
    "/carriers",
    response_model=CarrierListResponse,
    summary="Справочник перевозчиков (автокомплит)",
    responses=_ERROR_4XX,
)
def list_carriers(
    q: Optional[str] = Query(None, max_length=200),
    active: bool = Query(True),
    limit: int = Query(50, ge=1, le=500),
    _user: dict = Depends(REQUIRE_LOGISTICS),
    service: CarrierService = Depends(get_carrier_service),
) -> CarrierListResponse:
    try:
        return service.list_carriers(q=q, active=active, limit=limit)
    except CarrierError as exc:
        _raise_domain_error(exc, where="logistics.list_carriers")


@router.post(
    "/carriers/{carrier_id}/merge",
    response_model=CarrierMergeResponse,
    summary="Слияние перевозчика-дубликата",
    responses=_ERROR_4XX,
)
def merge_carrier(
    carrier_id: int,
    payload: CarrierMergeRequest,
    user: dict = Depends(REQUIRE_LOGISTICS),
    service: CarrierService = Depends(get_carrier_service),
) -> CarrierMergeResponse:
    try:
        return service.merge(carrier_id, payload.into_id, actor=_actor(user))
    except CarrierError as exc:
        _raise_domain_error(exc, where="logistics.merge_carrier")


@router.get(
    "/kp-search",
    response_model=LogisticsKpSearchResponse,
    summary="Поиск КП для рейса (в работе / На СГП)",
    responses=_ERROR_4XX,
)
def search_kp(
    kp_id: Optional[int] = Query(None, ge=1, description="Номер КП"),
    customer: Optional[str] = Query(None, max_length=128, description="Имя заказчика"),
    _user: dict = Depends(REQUIRE_LOGISTICS),
    service: ShipmentService = Depends(get_shipment_service),
) -> LogisticsKpSearchResponse:
    if kp_id is not None:
        return service.search_kp(kp_id=kp_id)

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

    return service.search_kp(customer=trimmed)


@router.get(
    "/pile-catalog",
    response_model=PileCatalogResponse,
    summary="Каталог свай (автокомплит)",
    responses=_ERROR_4XX,
)
def search_pile_catalog(
    q: str = Query("", max_length=100),
    limit: int = Query(20, ge=1, le=100),
    _user: dict = Depends(REQUIRE_LOGISTICS),
    service: ShipmentService = Depends(get_shipment_service),
) -> PileCatalogResponse:
    try:
        return service.search_pile_catalog(q, limit=limit)
    except ShipmentError as exc:
        _raise_domain_error(exc, where="logistics.search_pile_catalog")
