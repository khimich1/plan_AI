from __future__ import annotations

from fastapi import APIRouter, Depends, File, Form, Query, UploadFile

from app.core.http_errors import (
    raise_bad_request_client_error,
    raise_client_error,
    raise_unexpected_server_error,
    raise_unprocessable_client_error,
)
from app.dependencies.auth import REQUIRE_PRICES
from app.dependencies.services import (
    get_guid_explain_service,
    get_guid_replace_service,
    get_nomenclature_import_service,
    get_nomenclature_queue_service,
)
from app.schemas.guid_explain import (
    GuidExplainResponse,
    GuidReplaceMoneyOut,
    GuidReplacePreviewItemOut,
    GuidReplacePreviewResponse,
    GuidReplaceRequest,
    GuidReplaceResponse,
    GuidReplaceSkipOut,
    GuidReplaceWrittenOut,
    GuidSearchResponse,
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
from app.services.guid_explain_service import GuidExplainError, GuidExplainService
from app.services.guid_replace_service import GuidReplaceService
from core.guid_replace import GuidReplaceError
from app.services.nomenclature_import_service import (
    NomenclatureImportError,
    NomenclatureImportService,
)
from app.services.nomenclature_queue_service import (
    NomenclatureQueueError,
    NomenclatureQueueService,
)
from core.guid_explain import ROLE_ECONOMIST
from core.pricelist_1c_parser import Pricelist1CFormatError
from core.universal_report_1c_parser import UniversalReport1CFormatError

router = APIRouter(prefix="/nomenclature", tags=["nomenclature"])

_KIND_HELP = (
    "Группа изделий для смешанных выгрузок (файл ЛМ): "
    "pile, bridge_pile, fbs, stair_flight, stair_step. "
    "Если не задана — определяется по имени файла так же, как в сверке: "
    "сваи → pile, мостовые → bridge_pile, блоки → fbs. "
    "Плиты (имя файла или марки) пишутся в prays_plity. "
    "Составные и смешанный ЛМ без параметра → 422."
)


@router.post("/import-1c", response_model=Import1cResponse)
async def import_1c_pricelist(
    file: UploadFile = File(..., description="Выгрузка 1С: «Прайс-лист» (.xls) или универсальный отчёт (.xlsx)"),
    product_kind: str | None = Query(default=None, description=_KIND_HELP),
    product_kind_form: str | None = Form(default=None, alias="product_kind"),
    _user: dict = Depends(REQUIRE_PRICES),
    service: NomenclatureImportService = Depends(get_nomenclature_import_service),
) -> Import1cResponse:
    """Загрузить выгрузку 1С: не-плиты сверить с nomenclature_guid, плиты записать в prays_plity."""
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


@router.get("/guid-explain", response_model=GuidExplainResponse)
def explain_guid(
    product_kind: str = Query(..., min_length=1),
    mark: str = Query(..., min_length=1),
    _user: dict = Depends(REQUIRE_PRICES),
    service: GuidExplainService = Depends(get_guid_explain_service),
) -> GuidExplainResponse:
    """Почему нет GUID у марки: та же карточка и срез своего вида (экономист)."""
    try:
        return service.explain(product_kind, mark, role=ROLE_ECONOMIST)
    except GuidExplainError as exc:
        raise_unprocessable_client_error(
            exc, where="nomenclature.explain_guid", detail=str(exc)
        )
    except Exception as exc:
        raise_unexpected_server_error(exc, where="nomenclature.explain_guid")


@router.get("/guid-search", response_model=GuidSearchResponse)
def search_guid(
    product_kind: str = Query(..., min_length=1),
    q: str = Query(default=""),
    _user: dict = Depends(REQUIRE_PRICES),
    service: GuidExplainService = Depends(get_guid_explain_service),
) -> GuidSearchResponse:
    """Поиск марки своего вида в справочнике и прайсе (экономист)."""
    try:
        return service.search(product_kind, q, role=ROLE_ECONOMIST)
    except GuidExplainError as exc:
        raise_unprocessable_client_error(
            exc, where="nomenclature.search_guid", detail=str(exc)
        )
    except Exception as exc:
        raise_unexpected_server_error(exc, where="nomenclature.search_guid")


def _money(view) -> GuidReplaceMoneyOut:
    return GuidReplaceMoneyOut(
        unit_price=view.unit_price,
        unit_discounted=view.unit_discounted,
        line_sum=view.line_sum,
        total=view.total,
        subtotal=view.subtotal,
        vat_amount=view.vat_amount,
    )


def _preview_out(result) -> GuidReplacePreviewResponse:
    return GuidReplacePreviewResponse(
        items=[
            GuidReplacePreviewItemOut(
                kp_id=item.kp_id,
                qty=item.preview.qty,
                was=_money(item.preview.was),
                will=_money(item.preview.will),
                delivery_was=item.preview.delivery_was,
                delivery_will=item.preview.delivery_will,
                oneoff_notice=item.preview.oneoff_notice,
            )
            for item in result.items
        ],
        skipped=[
            GuidReplaceSkipOut(kp_id=item.kp_id, reason=item.reason) for item in result.skipped
        ],
    )


def _replace_out(result) -> GuidReplaceResponse:
    return GuidReplaceResponse(
        written=[
            GuidReplaceWrittenOut(kp_id=item.kp_id, total=item.total) for item in result.written
        ],
        skipped=[
            GuidReplaceSkipOut(kp_id=item.kp_id, reason=item.reason) for item in result.skipped
        ],
        queue_notice=result.queue_notice,
    )


@router.post("/guid-replace/preview", response_model=GuidReplacePreviewResponse)
def preview_guid_replace(
    payload: GuidReplaceRequest,
    _user: dict = Depends(REQUIRE_PRICES),
    service: GuidReplaceService = Depends(get_guid_replace_service),
) -> GuidReplacePreviewResponse:
    """Окно «было / станет» по КП строки очереди. Ничего не записывает."""
    try:
        result = service.preview_in_queue(
            product_kind=payload.product_kind,
            mark=payload.mark,
            target_mark=payload.target_mark,
            nomenclature_id=payload.nomenclature_id,
            replace_oneoff=payload.replace_oneoff,
        )
    except GuidReplaceError as exc:
        raise_bad_request_client_error(
            exc, where="nomenclature.preview_guid_replace", detail=exc.detail
        )
    except Exception as exc:
        raise_unexpected_server_error(exc, where="nomenclature.preview_guid_replace")
    return _preview_out(result)


@router.post("/guid-replace", response_model=GuidReplaceResponse)
def replace_guid(
    payload: GuidReplaceRequest,
    user: dict = Depends(REQUIRE_PRICES),
    service: GuidReplaceService = Depends(get_guid_replace_service),
) -> GuidReplaceResponse:
    """Заменить марку во всех КП «в архиве» открытой строки очереди."""
    try:
        result = service.replace_in_queue(
            product_kind=payload.product_kind,
            mark=payload.mark,
            target_mark=payload.target_mark,
            nomenclature_id=payload.nomenclature_id,
            changed_by=str(user.get("username") or "").strip() or None,
            replace_oneoff=payload.replace_oneoff,
        )
    except GuidReplaceError as exc:
        raise_bad_request_client_error(
            exc, where="nomenclature.replace_guid", detail=exc.detail
        )
    except Exception as exc:
        raise_unexpected_server_error(exc, where="nomenclature.replace_guid")
    return _replace_out(result)


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
