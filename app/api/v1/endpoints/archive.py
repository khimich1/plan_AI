from __future__ import annotations

import logging
from datetime import date
from typing import Literal
from urllib.parse import quote

from fastapi import APIRouter, Depends, File, HTTPException, Query, Request, UploadFile, status
from fastapi.responses import FileResponse, Response

from app.dependencies.auth import require_roles
from app.dependencies.plate_context import get_plate_order_context
from app.dependencies.services import get_archive_service, get_supply_contract_service
from app.core.http_errors import (
    MSG_ARCHIVE_NOT_FOUND,
    MSG_UNPROCESSABLE,
    MSG_VALIDATION,
    is_ai_provider_error,
    raise_ai_provider_unavailable_error,
    raise_bad_request_client_error,
    raise_client_error,
    raise_not_found_client_error,
    raise_unprocessable_client_error,
    raise_unexpected_server_error,
)
from app.schemas.archive import (
    ArchiveFileKind,
    ArchiveOfferDetails,
    ArchiveOfferListItem,
    ArchiveProductTypeFilter,
    ArchiveSearchResponse,
    ArchiveSection,
    BindCounterpartyRequest,
    CreateAndBindCounterpartyRequest,
    CapacitySnapshotResponse,
    KpReadinessPositionsResponse,
    ArchivePaymentRequest,
    InvoiceExportRequest,
    MoveToProductionRequest,
    SpecificationChoiceIn,
    SpecificationView,
    PromiseHoldResponse,
    PromiseQuoteResponse,
    PromiseTracksPerDayRequest,
    PromiseTracksPerDayResponse,
    PromiseWeekOccupantsResponse,
    UpdateDiscountRequest,
    UpdateLogisticsCostRequest,
)
from app.schemas.commercial import CommercialDraftDetailsResponse
from app.schemas.supply_contract import (
    SupplyContractCreate,
    SupplyContractImportRow,
    SupplyContractLinkRequest,
    SupplyContractOut,
    SupplyContractParseOut,
    SupplyContractPatch,
    SupplyContractRegistryItem,
)
from app.services.commercial_upload_validation import read_upload_file_capped
from app.services.archive_service import (
    ArchiveError,
    ArchiveNotFoundError,
    ArchiveService,
    ArchiveValidationError,
)
from app.services.supply_contract_service import (
    SupplyContractFieldError,
    SupplyContractNotFoundError,
    SupplyContractPdfError,
    SupplyContractService,
    SupplyContractValidationError,
)
from app.services.promise_service import (
    PromiseHoldForbiddenError,
    PromiseHoldNotFoundError,
    PromiseHoldUnavailableError,
    PromiseKnobInvalidError,
    PromiseNotFoundError,
    PromiseService,
    PromiseWeekInvalidError,
)
from core.plate_order_context import PlateOrderContext
from core.production.promise_buckets import OccupancyUnavailableError


logger = logging.getLogger(__name__)

router = APIRouter(prefix="/commercial/archive", tags=["commercial-archive"])
settings_router = APIRouter(prefix="/commercial/settings", tags=["commercial-settings"])

MSG_OCCUPANCY_UNAVAILABLE = (
    "Недоступна занятость плана — котировка остановлена (fail-closed)."
)


def get_promise_service() -> PromiseService:
    return PromiseService()


@router.get("", response_model=list[ArchiveOfferListItem])
def list_archive_offers(
    section: ArchiveSection = Query(default="archived"),
    product_type: ArchiveProductTypeFilter = Query(default="all"),
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> list[ArchiveOfferListItem]:
    return service.list_offers(section, product_type=product_type, user=user)


@router.get(
    "/search",
    response_model=ArchiveSearchResponse,
)
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


async def _read_supply_contract_patch(request: Request) -> tuple[SupplyContractPatch, bytes | None, str]:
    content_type = request.headers.get("content-type", "")
    if "multipart/form-data" in content_type:
        form = await request.form()
        raw: dict[str, str] = {}
        for key in ("status", "scan_note", "contract_date"):
            value = form.get(key)
            if isinstance(value, str):
                raw[key] = value
        upload = form.get("file")
        data: bytes | None = None
        filename = ""
        if upload is not None and hasattr(upload, "read"):
            data = await read_upload_file_capped(upload)
            filename = str(getattr(upload, "filename", "") or "")
        return SupplyContractPatch.model_validate(raw), data or None, filename
    payload = await request.json()
    if not isinstance(payload, dict):
        payload = {}
    return SupplyContractPatch.model_validate(payload), None, ""


@router.post("/supply-contracts/import", response_model=list[SupplyContractImportRow])
async def import_archive_supply_contracts(
    file: UploadFile = File(...),
    user: dict = Depends(require_roles("admin")),
    service: SupplyContractService = Depends(get_supply_contract_service),
) -> list[dict]:
    data = await read_upload_file_capped(file)
    try:
        return service.import_lawyer_sheet(data, user=user)
    except SupplyContractValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.import_supply_contracts",
            detail=str(exc) or MSG_VALIDATION,
        )


@router.post("/supply-contracts/{contract_id}/link", response_model=SupplyContractOut)
def link_archive_supply_contract(
    contract_id: int,
    payload: SupplyContractLinkRequest,
    user: dict = Depends(require_roles("admin", "manager")),
    service: SupplyContractService = Depends(get_supply_contract_service),
) -> dict:
    try:
        return service.link_imported_contract(
            contract_id,
            payload.counterparty_id,
            user=user,
        )
    except SupplyContractNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.link_supply_contract",
            detail=str(exc) or MSG_ARCHIVE_NOT_FOUND,
        )
    except SupplyContractValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.link_supply_contract",
            detail=str(exc) or MSG_VALIDATION,
        )


@router.get("/supply-contracts", response_model=list[SupplyContractRegistryItem])
def list_archive_supply_contracts(
    user: dict = Depends(require_roles("admin", "manager")),
    service: SupplyContractService = Depends(get_supply_contract_service),
) -> list[dict]:
    return service.list_registry(user=user)


@router.patch("/supply-contracts/{contract_id}", response_model=SupplyContractOut)
async def patch_archive_supply_contract(
    contract_id: int,
    request: Request,
    user: dict = Depends(require_roles("admin", "manager")),
    service: SupplyContractService = Depends(get_supply_contract_service),
) -> dict:
    fields, scan_bytes, scan_filename = await _read_supply_contract_patch(request)
    try:
        return service.update_contract(
            contract_id,
            user=user,
            status=fields.status,
            scan_note=fields.scan_note,
            scan_note_set="scan_note" in fields.model_fields_set,
            contract_date=fields.contract_date,
            scan_bytes=scan_bytes,
            scan_filename=scan_filename,
        )
    except SupplyContractNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.patch_supply_contract",
            detail=str(exc) or MSG_ARCHIVE_NOT_FOUND,
        )
    except SupplyContractValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.patch_supply_contract",
            detail=str(exc) or MSG_VALIDATION,
        )
    except SupplyContractFieldError as exc:
        raise_unprocessable_client_error(
            exc,
            where="archive.patch_supply_contract",
            detail=str(exc) or MSG_UNPROCESSABLE,
        )


@router.post("/supply-contracts/{contract_id}/replace", response_model=SupplyContractOut)
def replace_archive_supply_contract(
    contract_id: int,
    payload: SupplyContractCreate,
    user: dict = Depends(require_roles("admin", "manager")),
    service: SupplyContractService = Depends(get_supply_contract_service),
) -> dict:
    try:
        return service.replace_contract(contract_id, payload, user=user)
    except SupplyContractNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.replace_supply_contract",
            detail=str(exc) or MSG_ARCHIVE_NOT_FOUND,
        )
    except SupplyContractValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.replace_supply_contract",
            detail=str(exc) or MSG_VALIDATION,
        )
    except SupplyContractFieldError as exc:
        raise_unprocessable_client_error(
            exc,
            where="archive.replace_supply_contract",
            detail=str(exc) or MSG_UNPROCESSABLE,
        )


@router.get("/{kp_id}/readiness/positions", response_model=KpReadinessPositionsResponse)
def get_kp_readiness_positions(
    kp_id: int,
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> KpReadinessPositionsResponse:
    try:
        return service.get_readiness_positions(kp_id, user=user)
    except ArchiveNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.get_kp_readiness_positions",
            detail=MSG_ARCHIVE_NOT_FOUND,
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


@router.post("/{kp_id}/resume", response_model=CommercialDraftDetailsResponse)
def resume_archive_offer_as_draft(
    kp_id: int,
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> CommercialDraftDetailsResponse:
    try:
        result = service.resume_as_draft(kp_id, user=user)
    except ArchiveNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.resume_archive_offer_as_draft",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except ArchiveValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.resume_archive_offer_as_draft",
            detail=MSG_VALIDATION,
        )
    return CommercialDraftDetailsResponse.model_validate(result)


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


@router.get("/{kp_id}/supply-contract", response_model=SupplyContractOut | None)
def get_archive_supply_contract(
    kp_id: int,
    user: dict = Depends(require_roles("admin", "manager")),
    service: SupplyContractService = Depends(get_supply_contract_service),
) -> SupplyContractOut | None:
    try:
        return service.get_for_kp(kp_id, user=user)
    except SupplyContractNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.get_supply_contract",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except SupplyContractValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.get_supply_contract",
            detail=str(exc) or MSG_VALIDATION,
        )


@router.post("/{kp_id}/supply-contract", response_model=SupplyContractOut)
def create_archive_supply_contract(
    kp_id: int,
    payload: SupplyContractCreate,
    user: dict = Depends(require_roles("admin", "manager")),
    service: SupplyContractService = Depends(get_supply_contract_service),
) -> dict:
    try:
        return service.create_for_kp(kp_id, payload, user=user)
    except SupplyContractNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.create_supply_contract",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except SupplyContractValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.create_supply_contract",
            detail=str(exc) or MSG_VALIDATION,
        )
    except SupplyContractFieldError as exc:
        raise_unprocessable_client_error(
            exc,
            where="archive.create_supply_contract",
            detail=str(exc) or MSG_UNPROCESSABLE,
        )


_DOCX_MEDIA = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
_PDF_MEDIA = "application/pdf"


def _supply_document_response(payload: bytes, filename: str) -> Response:
    media = _PDF_MEDIA if filename.lower().endswith(".pdf") else _DOCX_MEDIA
    return Response(
        content=payload,
        media_type=media,
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.get("/{kp_id}/supply-contract/document")
def download_archive_supply_contract(
    kp_id: int,
    file_format: Literal["docx", "pdf"] = Query(default="docx", alias="format"),
    user: dict = Depends(require_roles("admin", "manager")),
    service: SupplyContractService = Depends(get_supply_contract_service),
) -> Response:
    try:
        payload, filename = service.document_for_kp(
            kp_id,
            user=user,
            file_format=file_format,
        )
    except SupplyContractNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.download_supply_contract",
            detail=str(exc) or MSG_ARCHIVE_NOT_FOUND,
        )
    except SupplyContractValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.download_supply_contract",
            detail=str(exc) or MSG_VALIDATION,
        )
    except SupplyContractPdfError as exc:
        raise_client_error(
            exc,
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail=str(exc),
            where="archive.download_supply_contract",
        )
    return _supply_document_response(payload, filename)


@router.get("/{kp_id}/supply-contract/edo-agreement")
def download_archive_edo_agreement(
    kp_id: int,
    file_format: Literal["docx", "pdf"] = Query(default="docx", alias="format"),
    user: dict = Depends(require_roles("admin", "manager")),
    service: SupplyContractService = Depends(get_supply_contract_service),
) -> Response:
    try:
        payload, filename = service.edo_agreement_for_kp(
            kp_id,
            user=user,
            file_format=file_format,
        )
    except SupplyContractNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.download_edo_agreement",
            detail=str(exc) or MSG_ARCHIVE_NOT_FOUND,
        )
    except SupplyContractValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.download_edo_agreement",
            detail=str(exc) or MSG_VALIDATION,
        )
    except SupplyContractPdfError as exc:
        raise_client_error(
            exc,
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail=str(exc),
            where="archive.download_edo_agreement",
        )
    return _supply_document_response(payload, filename)


@router.post("/{kp_id}/supply-contract/parse", response_model=SupplyContractParseOut)
async def parse_archive_supply_contract(
    kp_id: int,
    file: UploadFile = File(...),
    user: dict = Depends(require_roles("admin", "manager")),
    service: SupplyContractService = Depends(get_supply_contract_service),
) -> dict:
    data = await read_upload_file_capped(file)
    try:
        return await service.parse_for_kp(
            kp_id,
            data,
            filename=file.filename or "",
            user=user,
        )
    except SupplyContractNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.parse_supply_contract",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except SupplyContractValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.parse_supply_contract",
            detail=str(exc) or MSG_VALIDATION,
        )
    except Exception as exc:
        if is_ai_provider_error(exc):
            raise_ai_provider_unavailable_error(exc, where="archive.parse_supply_contract")
        raise_unexpected_server_error(exc, where="archive.parse_supply_contract")


@router.patch("/{kp_id}/counterparty", response_model=ArchiveOfferDetails)
def bind_archive_counterparty(
    kp_id: int,
    payload: BindCounterpartyRequest,
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> ArchiveOfferDetails:
    try:
        return service.bind_counterparty(kp_id, payload.counterparty_id, user=user)
    except ArchiveNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.bind_counterparty",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except ArchiveValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.bind_counterparty",
            detail=str(exc) or MSG_VALIDATION,
        )


@router.post("/{kp_id}/counterparty", response_model=ArchiveOfferDetails)
def create_and_bind_archive_counterparty(
    kp_id: int,
    payload: CreateAndBindCounterpartyRequest,
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> ArchiveOfferDetails:
    try:
        return service.create_and_bind_counterparty(
            kp_id,
            name=payload.name,
            code_1c=payload.code_1c,
            inn=payload.inn,
            kpp=payload.kpp,
            user=user,
        )
    except ArchiveNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.create_and_bind_counterparty",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except ArchiveValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.create_and_bind_counterparty",
            detail=str(exc) or MSG_VALIDATION,
        )


@router.patch("/{kp_id}/logistics-cost", response_model=ArchiveOfferDetails)
def update_archive_logistics_cost(
    kp_id: int,
    payload: UpdateLogisticsCostRequest,
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> ArchiveOfferDetails:
    try:
        logistics_kwargs: dict = {
            "pile_logistics_cost": payload.pile_logistics_cost,
            "pile_trip_overrides": payload.pile_trip_overrides,
        }
        if payload.long_pile_delivery is not None:
            logistics_kwargs["long_pile_delivery"] = payload.long_pile_delivery
        return service.update_logistics_cost(
            kp_id,
            payload.logistics_cost,
            user=user,
            **logistics_kwargs,
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


@router.post("/{kp_id}/invoice-export", response_model=ArchiveOfferDetails)
def export_archive_invoice(
    kp_id: int,
    payload: InvoiceExportRequest | None = None,
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> ArchiveOfferDetails:
    try:
        warehouse = None if payload is None else payload.warehouse
        return service.export_invoice(kp_id, user=user, warehouse=warehouse)
    except ArchiveNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.export_invoice",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except ArchiveValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.export_invoice",
            detail=str(exc) or MSG_VALIDATION,
        )
    except ArchiveError as exc:
        raise_unexpected_server_error(exc, where="archive.export_invoice")


@router.post("/{kp_id}/invoice-correction", response_model=ArchiveOfferDetails)
def export_archive_invoice_correction(
    kp_id: int,
    payload: InvoiceExportRequest | None = None,
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> ArchiveOfferDetails:
    try:
        warehouse = None if payload is None else payload.warehouse
        return service.export_invoice_correction(kp_id, user=user, warehouse=warehouse)
    except ArchiveNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.export_invoice_correction",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except ArchiveValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.export_invoice_correction",
            detail=str(exc) or MSG_VALIDATION,
        )
    except ArchiveError as exc:
        raise_unexpected_server_error(exc, where="archive.export_invoice_correction")


@router.get("/{kp_id}/specification/file")
def download_archive_specification(
    kp_id: int,
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> Response:
    try:
        downloaded = service.download_specification(kp_id, user=user)
    except ArchiveNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.download_specification",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except ArchiveValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.download_specification",
            detail=str(exc) or MSG_VALIDATION,
        )
    filename = quote(downloaded.filename)
    return Response(
        content=downloaded.content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{filename}"},
    )


@router.get("/{kp_id}/specification", response_model=SpecificationView)
def get_archive_specification(
    kp_id: int,
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> SpecificationView:
    try:
        return service.get_specification(kp_id, user=user)
    except ArchiveNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.get_specification",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except ArchiveValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.get_specification",
            detail=str(exc) or MSG_VALIDATION,
        )


@router.put("/{kp_id}/specification", response_model=SpecificationView)
def put_archive_specification(
    kp_id: int,
    payload: SpecificationChoiceIn,
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> SpecificationView:
    try:
        return service.save_specification(kp_id, payload, user=user)
    except ArchiveNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.save_specification",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except ArchiveValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.save_specification",
            detail=str(exc) or MSG_VALIDATION,
        )


@router.post("/{kp_id}/payment", response_model=ArchiveOfferDetails)
def set_archive_payment(
    kp_id: int,
    payload: ArchivePaymentRequest,
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> ArchiveOfferDetails:
    try:
        return service.set_payment(kp_id, paid=payload.paid, user=user)
    except ArchiveNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.set_payment",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except ArchiveValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.set_payment",
            detail=str(exc) or MSG_VALIDATION,
        )


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
            where="archive.move_to_production",
            detail=str(exc) or MSG_VALIDATION,
        )
    except ArchiveError as exc:
        raise_unexpected_server_error(exc, where="archive.move_to_production")


@router.get("/{kp_id}/promise-quote", response_model=PromiseQuoteResponse)
def get_promise_quote(
    kp_id: int,
    user: dict = Depends(require_roles("admin", "manager")),
    service: PromiseService = Depends(get_promise_service),
) -> PromiseQuoteResponse:
    try:
        return service.get_quote(kp_id, user=user)
    except PromiseNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.get_promise_quote",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except OccupancyUnavailableError as exc:
        logger.exception("promise-quote occupancy unavailable for kp_id=%s", kp_id)
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail=str(exc) or MSG_OCCUPANCY_UNAVAILABLE,
        ) from exc


@router.get(
    "/{kp_id}/promise-weeks/{week_start}/occupants",
    response_model=PromiseWeekOccupantsResponse,
)
def get_promise_week_occupants(
    kp_id: int,
    week_start: date,
    user: dict = Depends(require_roles("admin", "manager")),
    service: PromiseService = Depends(get_promise_service),
) -> PromiseWeekOccupantsResponse:
    try:
        return service.list_week_occupants(kp_id, week_start, user=user)
    except PromiseNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.get_promise_week_occupants",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except PromiseWeekInvalidError as exc:
        raise_unprocessable_client_error(
            exc,
            where="archive.get_promise_week_occupants",
            detail=str(exc) or MSG_VALIDATION,
        )
    except OccupancyUnavailableError as exc:
        logger.exception(
            "promise-week occupants occupancy unavailable for kp_id=%s", kp_id
        )
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail=str(exc) or MSG_OCCUPANCY_UNAVAILABLE,
        ) from exc


@router.post(
    "/{kp_id}/promise-hold",
    response_model=PromiseHoldResponse,
    status_code=status.HTTP_201_CREATED,
)
def create_promise_hold(
    kp_id: int,
    user: dict = Depends(require_roles("admin", "manager")),
    service: PromiseService = Depends(get_promise_service),
) -> PromiseHoldResponse:
    try:
        return service.create_hold(kp_id, user=user)
    except PromiseNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.create_promise_hold",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except PromiseHoldUnavailableError as exc:
        raise_unprocessable_client_error(
            exc,
            where="archive.create_promise_hold",
            detail=str(exc) or MSG_VALIDATION,
        )
    except OccupancyUnavailableError as exc:
        logger.exception("promise-hold occupancy unavailable for kp_id=%s", kp_id)
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail=str(exc) or MSG_OCCUPANCY_UNAVAILABLE,
        ) from exc


@router.get("/{kp_id}/promise-hold", response_model=PromiseHoldResponse)
def get_promise_hold(
    kp_id: int,
    user: dict = Depends(require_roles("admin", "manager")),
    service: PromiseService = Depends(get_promise_service),
) -> PromiseHoldResponse:
    try:
        hold = service.get_hold(kp_id, user=user)
    except PromiseNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.get_promise_hold",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    if hold is None or hold.status != "active":
        raise_not_found_client_error(
            PromiseHoldNotFoundError("Активный холд не найден."),
            where="archive.get_promise_hold",
            detail="Активный холд не найден.",
        )
    return hold


@router.delete("/{kp_id}/promise-hold", response_model=PromiseHoldResponse)
def delete_promise_hold(
    kp_id: int,
    user: dict = Depends(require_roles("admin", "manager")),
    service: PromiseService = Depends(get_promise_service),
) -> PromiseHoldResponse:
    try:
        return service.release_hold(kp_id, user=user)
    except PromiseNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.delete_promise_hold",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except PromiseHoldNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.delete_promise_hold",
            detail="Активный холд не найден.",
        )
    except PromiseHoldForbiddenError as exc:
        raise_client_error(
            exc,
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Снять холд может только владелец или администратор.",
            where="archive.delete_promise_hold",
        )


@router.get("/{kp_id}/capacity-snapshot", response_model=CapacitySnapshotResponse)
def get_capacity_snapshot(
    kp_id: int,
    target: str | None = Query(
        default=None,
        description="ISO YYYY-MM-DD дедлайн производства; иначе из срока КП",
    ),
    user: dict = Depends(require_roles("admin", "manager")),
    service: ArchiveService = Depends(get_archive_service),
) -> CapacitySnapshotResponse:
    try:
        return service.get_capacity_snapshot(kp_id, user=user, target=target)
    except ArchiveNotFoundError as exc:
        raise_not_found_client_error(
            exc,
            where="archive.get_capacity_snapshot",
            detail=MSG_ARCHIVE_NOT_FOUND,
        )
    except ArchiveValidationError as exc:
        raise_bad_request_client_error(
            exc,
            where="archive.get_capacity_snapshot",
            detail=str(exc) or MSG_VALIDATION,
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


@settings_router.get(
    "/promise-tracks-per-day",
    response_model=PromiseTracksPerDayResponse,
)
def get_promise_tracks_per_day(
    user: dict = Depends(require_roles("admin", "manager")),
    service: PromiseService = Depends(get_promise_service),
) -> PromiseTracksPerDayResponse:
    return service.get_tracks_per_day(user=user)


@settings_router.put(
    "/promise-tracks-per-day",
    response_model=PromiseTracksPerDayResponse,
)
def put_promise_tracks_per_day(
    payload: PromiseTracksPerDayRequest,
    user: dict = Depends(require_roles("admin", "manager")),
    service: PromiseService = Depends(get_promise_service),
) -> PromiseTracksPerDayResponse:
    try:
        return service.set_tracks_per_day(payload.tracks_per_day, user=user)
    except PromiseKnobInvalidError as exc:
        raise_unprocessable_client_error(
            exc,
            where="archive.put_promise_tracks_per_day",
            detail=str(exc) or MSG_VALIDATION,
        )


_archive_router = router
router = APIRouter()
router.include_router(_archive_router)
router.include_router(settings_router)
