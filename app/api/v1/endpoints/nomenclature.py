from __future__ import annotations

from fastapi import APIRouter, Depends, File, Form, Query, UploadFile

from app.core.http_errors import (
    raise_client_error,
    raise_unexpected_server_error,
    raise_unprocessable_client_error,
)
from app.dependencies.auth import REQUIRE_PRICES
from app.dependencies.services import (
    get_nomenclature_import_service,
    get_nomenclature_queue_service,
)
from app.schemas.nomenclature import (
    DuplicateResolveRequest,
    DuplicateResolveResponse,
    GuidTasksResponse,
    Import1cResponse,
    PriceQueueResolveRequest,
    PriceQueueResolveResponse,
)
from app.services.commercial_upload_validation import read_upload_file_capped
from app.services.nomenclature_import_service import (
    NomenclatureImportError,
    NomenclatureImportService,
)
from app.services.nomenclature_queue_service import (
    NomenclatureQueueError,
    NomenclatureQueueService,
)
from core.pricelist_1c_parser import Pricelist1CFormatError
from core.universal_report_1c_parser import UniversalReport1CFormatError

router = APIRouter(prefix="/nomenclature", tags=["nomenclature"])

_KIND_HELP = (
    "Группа изделий для смешанных выгрузок (файл ЛМ): "
    "pile, bridge_pile, fbs, stair_flight, stair_step. "
    "Если не задана — определяется по имени файла так же, как в сверке: "
    "сваи → pile, мостовые → bridge_pile, блоки → fbs. "
    "Плиты, составные и смешанный ЛМ без параметра → 422."
)


@router.post("/import-1c", response_model=Import1cResponse)
async def import_1c_pricelist(
    file: UploadFile = File(..., description="Выгрузка 1С: «Прайс-лист» (.xls) или универсальный отчёт (.xlsx)"),
    product_kind: str | None = Query(default=None, description=_KIND_HELP),
    product_kind_form: str | None = Form(default=None, alias="product_kind"),
    _user: dict = Depends(REQUIRE_PRICES),
    service: NomenclatureImportService = Depends(get_nomenclature_import_service),
) -> Import1cResponse:
    """Загрузить выгрузку 1С и сверить GUID с nomenclature_guid в pb.db."""
    file_bytes = await read_upload_file_capped(file)
    kind = _first_nonempty(product_kind_form, product_kind)
    try:
        return service.import_file(
            file_bytes,
            file.filename,
            product_kind=kind,
        )
    except (Pricelist1CFormatError, UniversalReport1CFormatError, NomenclatureImportError) as exc:
        raise_unprocessable_client_error(
            exc,
            where="nomenclature.import_1c_pricelist",
            detail=str(exc),
        )
    except ValueError as exc:
        raise_unprocessable_client_error(
            exc,
            where="nomenclature.import_1c_pricelist",
            detail=str(exc),
        )
    except Exception as exc:
        raise_unexpected_server_error(
            exc,
            where="nomenclature.import_1c_pricelist",
        )


@router.get("/tasks", response_model=GuidTasksResponse)
def list_guid_tasks(
    _user: dict = Depends(REQUIRE_PRICES),
    service: NomenclatureQueueService = Depends(get_nomenclature_queue_service),
) -> GuidTasksResponse:
    try:
        return service.list_tasks()
    except Exception as exc:
        raise_unexpected_server_error(exc, where="nomenclature.list_guid_tasks")


@router.get("/duplicates", response_model=GuidTasksResponse)
def list_guid_duplicates(
    _user: dict = Depends(REQUIRE_PRICES),
    service: NomenclatureQueueService = Depends(get_nomenclature_queue_service),
) -> GuidTasksResponse:
    try:
        snap = service.list_tasks()
    except Exception as exc:
        raise_unexpected_server_error(exc, where="nomenclature.list_guid_duplicates")
    return GuidTasksResponse(duplicates=snap.duplicates)


@router.post("/price-queue/resolve", response_model=PriceQueueResolveResponse)
def resolve_price_queue(
    payload: PriceQueueResolveRequest,
    user: dict = Depends(REQUIRE_PRICES),
    service: NomenclatureQueueService = Depends(get_nomenclature_queue_service),
) -> PriceQueueResolveResponse:
    try:
        return service.resolve_price(
            payload.guid,
            payload.price,
            resolved_by=str(user.get("username") or user.get("id") or ""),
        )
    except NomenclatureQueueError as exc:
        raise_client_error(
            exc,
            status_code=exc.status_code,
            detail=exc.message,
            where="nomenclature.resolve_price_queue",
        )
    except Exception as exc:
        raise_unexpected_server_error(exc, where="nomenclature.resolve_price_queue")


@router.post("/duplicates/resolve", response_model=DuplicateResolveResponse)
def resolve_guid_duplicate(
    payload: DuplicateResolveRequest,
    user: dict = Depends(REQUIRE_PRICES),
    service: NomenclatureQueueService = Depends(get_nomenclature_queue_service),
) -> DuplicateResolveResponse:
    try:
        return service.resolve_duplicate(
            scope=payload.scope,
            key=payload.key,
            chosen_guid=payload.chosen_guid,
            note=payload.note,
            decided_by=str(user.get("username") or user.get("id") or ""),
            product_kind=payload.product_kind,
        )
    except NomenclatureQueueError as exc:
        raise_client_error(
            exc,
            status_code=exc.status_code,
            detail=exc.message,
            where="nomenclature.resolve_guid_duplicate",
        )
    except Exception as exc:
        raise_unexpected_server_error(exc, where="nomenclature.resolve_guid_duplicate")


def _first_nonempty(*values: str | None) -> str | None:
    for raw in values:
        if raw is not None and str(raw).strip():
            return str(raw).strip()
    return None
