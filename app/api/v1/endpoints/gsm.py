"""GSM module HTTP endpoints (transactions import + registry CRUD + generation)."""

from __future__ import annotations

from datetime import date
from typing import NoReturn

from fastapi import APIRouter, Depends, File, Query, UploadFile, status
from fastapi.responses import Response

from app.core.http_errors import raise_structured_error
from app.dependencies.auth import REQUIRE_ACCOUNTING
from app.dependencies.services import (
    get_gsm_export_service,
    get_gsm_generation_service,
    get_gsm_registry_service,
    get_gsm_transaction_service,
)
from app.schemas.gsm import (
    CardCreateRequest,
    CardOut,
    CardPatchRequest,
    DriverCreateRequest,
    DriverOut,
    DriverPatchRequest,
    GsmSettings,
    StationCreateRequest,
    StationOut,
    StationPatchRequest,
    TransactionImportReport,
    RouteOut,
    VehicleCreateRequest,
    VehicleOut,
    VehiclePatchRequest,
    WaybillCreateRequest,
    WaybillExportRequest,
    WaybillGenerateRequest,
    WaybillGenerateResult,
    WaybillOut,
    WaybillPatchRequest,
)
from app.services.commercial_upload_validation import read_upload_file_capped
from app.services.gsm_export_service import GsmExportError, GsmExportService
from app.services.gsm_generation_service import GsmGenerationError, GsmGenerationService
from app.services.gsm_registry_service import GsmRegistryError, GsmRegistryService
from app.services.gsm_transaction_service import GsmTransactionService

router = APIRouter(prefix="/gsm", tags=["gsm"])

_ERROR_4XX: dict[int, dict] = {
    403: {"description": "Недостаточно прав"},
    404: {"description": "Не найдено"},
    409: {"description": "Конфликт (confirmed/exported без force)"},
    422: {"description": "Ошибка валидации доменной логики"},
}


def _raise_registry_error(exc: GsmRegistryError, *, where: str) -> NoReturn:
    http_status = (
        status.HTTP_404_NOT_FOUND
        if exc.code.endswith("_not_found")
        else status.HTTP_422_UNPROCESSABLE_ENTITY
    )
    raise_structured_error(
        status_code=http_status,
        code=exc.code,
        message=str(exc),
        details={},
        where=where,
    )


def _raise_generation_error(exc: GsmGenerationError, *, where: str) -> NoReturn:
    if exc.code.endswith("_not_found"):
        http_status = status.HTTP_404_NOT_FOUND
    elif exc.code in {"gsm_confirmed_conflict", "gsm_waybill_conflict"}:
        http_status = status.HTTP_409_CONFLICT
    elif exc.code == "gsm_invalid_period":
        http_status = status.HTTP_400_BAD_REQUEST
    else:
        http_status = status.HTTP_422_UNPROCESSABLE_ENTITY
    raise_structured_error(
        status_code=http_status,
        code=exc.code,
        message=str(exc),
        details=exc.details,
        where=where,
    )


def _raise_export_error(exc: GsmExportError, *, where: str) -> NoReturn:
    if exc.code.endswith("_not_found") or exc.code == "gsm_export_empty":
        http_status = status.HTTP_404_NOT_FOUND
    elif exc.code in {"gsm_invalid_period"}:
        http_status = status.HTTP_400_BAD_REQUEST
    elif exc.code.startswith("gsm_export_soffice") or exc.code == "gsm_export_template_missing":
        http_status = status.HTTP_500_INTERNAL_SERVER_ERROR
    else:
        http_status = status.HTTP_422_UNPROCESSABLE_ENTITY
    raise_structured_error(
        status_code=http_status,
        code=exc.code,
        message=str(exc),
        details=exc.details,
        where=where,
    )


# ---------------------------------------------------------------------------
# Transactions
# ---------------------------------------------------------------------------


@router.post("/transactions/import", response_model=TransactionImportReport)
async def import_transactions(
    files: list[UploadFile] = File(...),
    user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmTransactionService = Depends(get_gsm_transaction_service),
) -> TransactionImportReport:
    payload: list[tuple[str, bytes]] = []
    for upload in files:
        content = await read_upload_file_capped(upload)
        name = upload.filename or "upload.xls"
        payload.append((name, content))
    uploaded_by = (
        user.get("username")
        or user.get("email")
        or user.get("login")
        or None
    )
    if uploaded_by is not None:
        uploaded_by = str(uploaded_by)
    return service.import_files(payload, uploaded_by=uploaded_by)


# ---------------------------------------------------------------------------
# Vehicles
# ---------------------------------------------------------------------------


@router.get("/vehicles", response_model=list[VehicleOut], responses=_ERROR_4XX)
def list_vehicles(
    active_only: bool = Query(True),
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmRegistryService = Depends(get_gsm_registry_service),
) -> list[VehicleOut]:
    return service.list_vehicles(active_only=active_only)


@router.post(
    "/vehicles",
    response_model=VehicleOut,
    status_code=status.HTTP_201_CREATED,
    responses=_ERROR_4XX,
)
def create_vehicle(
    payload: VehicleCreateRequest,
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmRegistryService = Depends(get_gsm_registry_service),
) -> VehicleOut:
    try:
        return service.create_vehicle(**payload.model_dump())
    except GsmRegistryError as exc:
        _raise_registry_error(exc, where="gsm.create_vehicle")


@router.patch(
    "/vehicles/{vehicle_id}",
    response_model=VehicleOut,
    responses=_ERROR_4XX,
)
def patch_vehicle(
    vehicle_id: int,
    payload: VehiclePatchRequest,
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmRegistryService = Depends(get_gsm_registry_service),
) -> VehicleOut:
    try:
        return service.patch_vehicle(
            vehicle_id, **payload.model_dump(exclude_unset=True)
        )
    except GsmRegistryError as exc:
        _raise_registry_error(exc, where="gsm.patch_vehicle")


# ---------------------------------------------------------------------------
# Drivers
# ---------------------------------------------------------------------------


@router.get("/drivers", response_model=list[DriverOut], responses=_ERROR_4XX)
def list_drivers(
    active_only: bool = Query(True),
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmRegistryService = Depends(get_gsm_registry_service),
) -> list[DriverOut]:
    return service.list_drivers(active_only=active_only)


@router.post(
    "/drivers",
    response_model=DriverOut,
    status_code=status.HTTP_201_CREATED,
    responses=_ERROR_4XX,
)
def create_driver(
    payload: DriverCreateRequest,
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmRegistryService = Depends(get_gsm_registry_service),
) -> DriverOut:
    try:
        return service.create_driver(**payload.model_dump())
    except GsmRegistryError as exc:
        _raise_registry_error(exc, where="gsm.create_driver")


@router.patch(
    "/drivers/{driver_id}",
    response_model=DriverOut,
    responses=_ERROR_4XX,
)
def patch_driver(
    driver_id: int,
    payload: DriverPatchRequest,
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmRegistryService = Depends(get_gsm_registry_service),
) -> DriverOut:
    try:
        return service.patch_driver(
            driver_id, **payload.model_dump(exclude_unset=True)
        )
    except GsmRegistryError as exc:
        _raise_registry_error(exc, where="gsm.patch_driver")


# ---------------------------------------------------------------------------
# Cards
# ---------------------------------------------------------------------------


@router.get("/cards", response_model=list[CardOut], responses=_ERROR_4XX)
def list_cards(
    include_archived: bool = Query(False),
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmRegistryService = Depends(get_gsm_registry_service),
) -> list[CardOut]:
    return service.list_cards(include_archived=include_archived)


@router.post(
    "/cards",
    response_model=CardOut,
    status_code=status.HTTP_201_CREATED,
    responses=_ERROR_4XX,
)
def create_card(
    payload: CardCreateRequest,
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmRegistryService = Depends(get_gsm_registry_service),
) -> CardOut:
    try:
        return service.create_card(**payload.model_dump())
    except GsmRegistryError as exc:
        _raise_registry_error(exc, where="gsm.create_card")


@router.patch(
    "/cards/{card_id}",
    response_model=CardOut,
    responses=_ERROR_4XX,
)
def patch_card(
    card_id: int,
    payload: CardPatchRequest,
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmRegistryService = Depends(get_gsm_registry_service),
) -> CardOut:
    try:
        return service.patch_card(card_id, **payload.model_dump(exclude_unset=True))
    except GsmRegistryError as exc:
        _raise_registry_error(exc, where="gsm.patch_card")


# ---------------------------------------------------------------------------
# Stations
# ---------------------------------------------------------------------------


@router.get("/stations", response_model=list[StationOut], responses=_ERROR_4XX)
def list_stations(
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmRegistryService = Depends(get_gsm_registry_service),
) -> list[StationOut]:
    return service.list_stations()


@router.post(
    "/stations",
    response_model=StationOut,
    status_code=status.HTTP_201_CREATED,
    responses=_ERROR_4XX,
)
def create_station(
    payload: StationCreateRequest,
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmRegistryService = Depends(get_gsm_registry_service),
) -> StationOut:
    try:
        return service.create_station(**payload.model_dump())
    except GsmRegistryError as exc:
        _raise_registry_error(exc, where="gsm.create_station")


@router.patch(
    "/stations/{station_id}",
    response_model=StationOut,
    responses=_ERROR_4XX,
)
def patch_station(
    station_id: int,
    payload: StationPatchRequest,
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmRegistryService = Depends(get_gsm_registry_service),
) -> StationOut:
    try:
        return service.patch_station(
            station_id, **payload.model_dump(exclude_unset=True)
        )
    except GsmRegistryError as exc:
        _raise_registry_error(exc, where="gsm.patch_station")


# ---------------------------------------------------------------------------
# Routes (vehicle library)
# ---------------------------------------------------------------------------


@router.get("/routes", response_model=list[RouteOut], responses=_ERROR_4XX)
def list_routes(
    vehicle_id: int = Query(...),
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmRegistryService = Depends(get_gsm_registry_service),
) -> list[RouteOut]:
    try:
        return service.list_routes(vehicle_id=vehicle_id)
    except GsmRegistryError as exc:
        _raise_registry_error(exc, where="gsm.list_routes")


# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------


@router.get("/settings", response_model=GsmSettings, responses=_ERROR_4XX)
def get_settings(
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmRegistryService = Depends(get_gsm_registry_service),
) -> GsmSettings:
    return service.get_settings()


@router.put("/settings", response_model=GsmSettings, responses=_ERROR_4XX)
def put_settings(
    payload: GsmSettings,
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmRegistryService = Depends(get_gsm_registry_service),
) -> GsmSettings:
    try:
        return service.put_settings(
            winter_start=payload.winter_start,
            hook_threshold_km=payload.hook_threshold_km,
            max_daily_km=payload.max_daily_km,
        )
    except GsmRegistryError as exc:
        _raise_registry_error(exc, where="gsm.put_settings")


# ---------------------------------------------------------------------------
# Waybills / generation
# ---------------------------------------------------------------------------


@router.post(
    "/waybills/generate",
    response_model=WaybillGenerateResult,
    responses=_ERROR_4XX,
)
def generate_waybills(
    payload: WaybillGenerateRequest,
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmGenerationService = Depends(get_gsm_generation_service),
) -> WaybillGenerateResult:
    try:
        return service.generate(
            vehicle_id=payload.vehicle_id,
            period_from=payload.period_from,
            period_to=payload.period_to,
            force=payload.force,
            fuel_start=payload.fuel_start,
            odometer_start=payload.odometer_start,
        )
    except GsmGenerationError as exc:
        _raise_generation_error(exc, where="gsm.generate_waybills")


@router.get(
    "/waybills",
    response_model=list[WaybillOut],
    responses=_ERROR_4XX,
)
def list_waybills(
    vehicle_id: int = Query(...),
    period_from: date = Query(..., alias="from"),
    period_to: date = Query(..., alias="to"),
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmGenerationService = Depends(get_gsm_generation_service),
) -> list[WaybillOut]:
    try:
        return service.list_waybills(
            vehicle_id=vehicle_id,
            period_from=period_from,
            period_to=period_to,
        )
    except GsmGenerationError as exc:
        _raise_generation_error(exc, where="gsm.list_waybills")


@router.post(
    "/waybills",
    response_model=WaybillOut,
    status_code=status.HTTP_201_CREATED,
    responses=_ERROR_4XX,
)
def create_waybill(
    payload: WaybillCreateRequest,
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmGenerationService = Depends(get_gsm_generation_service),
) -> WaybillOut:
    try:
        return service.create_waybill(
            vehicle_id=payload.vehicle_id,
            day=payload.date,
            driver_id=payload.driver_id,
            route=payload.route,
            fuel_issued=payload.fuel_issued,
            fuel_start=payload.fuel_start,
            odometer_start=payload.odometer_start,
        )
    except GsmGenerationError as exc:
        _raise_generation_error(exc, where="gsm.create_waybill")


@router.patch(
    "/waybills/{waybill_id}",
    response_model=WaybillOut,
    responses=_ERROR_4XX,
)
def patch_waybill(
    waybill_id: int,
    payload: WaybillPatchRequest,
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmGenerationService = Depends(get_gsm_generation_service),
) -> WaybillOut:
    try:
        return service.patch_waybill(
            waybill_id,
            driver_id=payload.driver_id,
            km=payload.km,
            route=payload.route,
        )
    except GsmGenerationError as exc:
        _raise_generation_error(exc, where="gsm.patch_waybill")


@router.post(
    "/waybills/{waybill_id}/confirm",
    response_model=WaybillOut,
    responses=_ERROR_4XX,
)
def confirm_waybill(
    waybill_id: int,
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmGenerationService = Depends(get_gsm_generation_service),
) -> WaybillOut:
    try:
        return service.confirm_waybill(waybill_id)
    except GsmGenerationError as exc:
        _raise_generation_error(exc, where="gsm.confirm_waybill")


@router.post(
    "/waybills/export",
    responses={
        **_ERROR_4XX,
        500: {"description": "Ошибка LibreOffice / экспорта бланка"},
    },
)
def export_waybills(
    payload: WaybillExportRequest,
    _user: dict = Depends(REQUIRE_ACCOUNTING),
    service: GsmExportService = Depends(get_gsm_export_service),
) -> Response:
    try:
        data, filename = service.export_zip(
            vehicle_ids=payload.vehicle_ids,
            period_from=payload.period_from,
            period_to=payload.period_to,
        )
    except GsmExportError as exc:
        _raise_export_error(exc, where="gsm.export_waybills")
    headers = {
        "Content-Disposition": f'attachment; filename="{filename}"',
    }
    return Response(content=data, media_type="application/zip", headers=headers)
