from __future__ import annotations

import logging
from datetime import date, timedelta

from typing import Literal

from fastapi import APIRouter, BackgroundTasks, Depends, HTTPException, Query, status
from fastapi.responses import FileResponse

from app.dependencies.auth import require_roles
from app.dependencies.plate_context import get_plate_order_context
from app.dependencies.services import (
    get_production_capacity_service,
    get_production_service,
    get_sgp_service,
)
from app.services.day_documents_service import (
    DayDocumentsError,
    generate_day_breakdown,
    generate_day_formovka,
    generate_day_schema,
    make_cleanup_callback,
)
from app.schemas.production import (
    AnalyzeSubstratesRequest,
    AnalyzeSubstratesResponse,
    BuildPlanRequest,
    BuildPlanResponse,
    CompleteProductionDayRequest,
    CreatePlanRequest,
    DayCapacityMapResponse,
    DayOccupancyResponse,
    DayViewDetailResponse,
    DeletePlanResponse,
    KpCandidatesResponse,
    RemoveTrackResponse,
    SaveDayCapacityRequest,
    SaveDayCapacityResponse,
    SaveWorkCalendarRequest,
)
from app.schemas.sgp import (
    SgpFreePlatesResponse,
    SgpMutationResponse,
    SgpPlatesResponse,
    SgpRelinkRequest,
    SgpUnlinkRequest,
)
from app.concurrency.cpu_bound import run_cpu_bound
from app.core.http_errors import (
    MSG_DAY_NOT_FOUND,
    MSG_PLAN_VERSION_CONFLICT,
    raise_bad_request_client_error,
    raise_not_found_client_error,
    raise_structured_error,
    raise_track_removal_client_error,
    raise_unexpected_server_error,
    raise_unprocessable_client_error,
)
from core.plate_order_context import PlateOrderContext
from app.repositories.plan_errors import PlanVersionConflict
from app.schemas.errors import ERROR_CODE_PLAN_VERSION_CONFLICT, ERROR_CODE_REST_VALIDATION_FAILED
from app.services.production_capacity_service import (
    ProductionCapacityError,
    ProductionCapacityService,
)
from app.services.production_completion_service import (
    ProductionCompletionError,
    ProductionRestDbError,
    ProductionRestValidationError,
)
from app.services.production_planning_service import ProductionPlanBuildError
from app.services.production_service import (
    ProductionAnalyzeBadRequest,
    ProductionAnalyzeEmptyBacklog,
    ProductionService,
    ProductionTrackRemovalError,
)
from app.services.promise_service import PromiseExclusionError, PromiseService
from app.services.sgp_service import SgpError, SgpService

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/production", tags=["production"])


@router.get("/plans")
def list_plans(
    _user: dict = Depends(require_roles("admin", "production")),
    service: ProductionService = Depends(get_production_service),
) -> dict:
    return service.list_plans()


@router.post("/plans")
async def create_plan(
    payload: CreatePlanRequest,
    _user: dict = Depends(require_roles("admin", "production")),
    service: ProductionService = Depends(get_production_service),
) -> dict:
    return await run_cpu_bound(lambda: service.create_plan(**payload.model_dump()))


@router.post("/plans/build", response_model=BuildPlanResponse)
async def build_plan_from_filters(
    payload: BuildPlanRequest,
    plate_order_ctx: PlateOrderContext = Depends(get_plate_order_context),
    user: dict = Depends(require_roles("admin", "production")),
    service: ProductionService = Depends(get_production_service),
) -> BuildPlanResponse:
    try:
        result = await run_cpu_bound(
            lambda: service.build_plan_from_filters(
                start_date=payload.start_date,
                tracks_count=payload.tracks_count,
                filter_method=payload.filter_method,
                selected_kp_ids=payload.selected_kp_ids or None,
                selected_plate_ids=payload.selected_plate_ids or None,
                selected_plate_qty=payload.selected_plate_qty or None,
                active_plan_id=payload.active_plan_id,
                plan_name=payload.plan_name,
                fill_targets=(
                    [item.model_dump(mode="json") for item in payload.fill_targets]
                    if payload.fill_targets
                    else None
                ),
                layout_reinforcement_order=payload.layout_reinforcement_order,
                plate_order_ctx=plate_order_ctx,
                sgp_reservations=(
                    [item.model_dump() for item in payload.sgp_reservations]
                    if payload.sgp_reservations
                    else None
                ),
            ),
            plate_order_ctx=plate_order_ctx,
        )
    except ProductionPlanBuildError as exc:
        raise_unprocessable_client_error(
            exc,
            where="production.build_plan",
            detail=str(exc),
        )
    except SgpError as exc:
        raise_structured_error(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            code=exc.code,
            message=str(exc),
            details={},
            where="production.build_plan.sgp_reservations",
        )
    except PlanVersionConflict as exc:
        raise_structured_error(
            status_code=status.HTTP_409_CONFLICT,
            code=ERROR_CODE_PLAN_VERSION_CONFLICT,
            message=MSG_PLAN_VERSION_CONFLICT,
            details={
                "plan_id": exc.plan_id,
                "expected_version": exc.expected_version,
            },
            where="production.build_plan",
        )
    except Exception as exc:
        logger.exception("[production/build] Непредвиденная ошибка: %s", exc)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Не удалось построить план производства.",
        ) from exc
    try:
        _record_build_exclusions(payload, result, user, service)
    except PromiseExclusionError as exc:
        raise_unprocessable_client_error(
            exc,
            where="production.build_plan.exclusions",
            detail=str(exc),
        )
    except Exception:
        logger.exception(
            "[production/build] Журнал исключений не записан (план уже собран)."
        )
    return BuildPlanResponse(**result)


@router.post("/analyze-substrates", response_model=AnalyzeSubstratesResponse)
async def analyze_substrates(
    payload: AnalyzeSubstratesRequest,
    user: dict = Depends(require_roles("admin", "production")),
    service: ProductionService = Depends(get_production_service),
) -> AnalyzeSubstratesResponse:
    try:
        result = await run_cpu_bound(
            lambda: service.analyze_substrates(
                fill_targets=[
                    item.model_dump(mode="json") for item in payload.fill_targets
                ],
                deadline_until=payload.deadline_until,
                user=user,
            )
        )
    except ProductionAnalyzeBadRequest as exc:
        raise_unprocessable_client_error(
            exc,
            where="production.analyze_substrates",
            detail=str(exc),
        )
    except ProductionAnalyzeEmptyBacklog as exc:
        raise_unprocessable_client_error(
            exc,
            where="production.analyze_substrates",
            detail=str(exc),
        )
    except Exception as exc:
        logger.exception("[production/analyze-substrates] Непредвиденная ошибка: %s", exc)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Не удалось выполнить анализ подложек.",
        ) from exc
    return AnalyzeSubstratesResponse(**result)


@router.get("/plans/{plan_id}")
def get_plan(
    plan_id: str,
    _user: dict = Depends(require_roles("admin", "production")),
    service: ProductionService = Depends(get_production_service),
) -> dict:
    plan = service.get_plan(plan_id)
    if not plan:
        raise HTTPException(status_code=404, detail="Plan not found")
    return plan


@router.get("/plans/{plan_id}/sgp-export")
def export_plan_sgp(
    plan_id: str,
    _user: dict = Depends(require_roles("admin", "production")),
    service: ProductionService = Depends(get_production_service),
    sgp: SgpService = Depends(get_sgp_service),
):
    """XLSX «Со склада» — позиции плана, закрытые со СГП (не в схеме/формовке)."""
    from fastapi.responses import Response

    plan = service.get_plan(plan_id)
    if not plan:
        raise HTTPException(status_code=404, detail="Plan not found")
    data = sgp.export_plan_sgp_xlsx(plan_id, plan)
    filename = f"SGP_{plan_id}.xlsx"
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.delete("/plans/{plan_id}", response_model=DeletePlanResponse)
def delete_plan(
    plan_id: str,
    _user: dict = Depends(require_roles("admin", "production")),
    service: ProductionService = Depends(get_production_service),
) -> DeletePlanResponse:
    result = service.delete_plan(plan_id)
    if not result.get("deleted"):
        raise HTTPException(status_code=404, detail="Plan not found")
    return DeletePlanResponse(**result)


@router.post("/plans/{plan_id}/activate")
def activate_plan(
    plan_id: str,
    _user: dict = Depends(require_roles("admin", "production")),
    service: ProductionService = Depends(get_production_service),
) -> dict:
    result = service.activate_plan(plan_id)
    if not result:
        raise HTTPException(status_code=404, detail="Plan not found")
    return result


@router.delete(
    "/plans/{plan_id}/days/{date}/tracks/{track_index}",
    response_model=RemoveTrackResponse,
)
def remove_track_from_plan(
    plan_id: str,
    date: date,
    track_index: int,
    expected_version: int | None = Query(default=None),
    user: dict = Depends(require_roles("admin", "production")),
    service: ProductionService = Depends(get_production_service),
) -> RemoveTrackResponse:
    actor = user.get("email") or user.get("login") or user.get("user_id")
    day_iso = date.isoformat()
    try:
        result = service.remove_track(
            plan_id=plan_id,
            date=day_iso,
            track_index=track_index,
            actor=str(actor) if actor else None,
            expected_version=expected_version,
        )
    except ProductionTrackRemovalError as exc:
        raise_track_removal_client_error(
            exc,
            where="production.remove_track",
            status_code=exc.status_code,
            code=exc.code,
        )
    except PlanVersionConflict as exc:
        raise_structured_error(
            status_code=status.HTTP_409_CONFLICT,
            code=ERROR_CODE_PLAN_VERSION_CONFLICT,
            message=MSG_PLAN_VERSION_CONFLICT,
            details={
                "plan_id": exc.plan_id,
                "expected_version": exc.expected_version,
            },
            where="production.remove_track",
        )
    return RemoveTrackResponse(**result)


@router.get("/calendar")
def get_calendar(
    _user: dict = Depends(require_roles("admin", "production")),
    service: ProductionService = Depends(get_production_service),
) -> dict:
    calendar = service.get_calendar()
    if not calendar:
        return {"plans_count": 0, "days_info": {}, "completed_days": []}
    return calendar


@router.get("/day-occupancy", response_model=DayOccupancyResponse)
def get_day_occupancy(
    exclude_plan_id: str | None = None,
    _user: dict = Depends(require_roles("admin", "production")),
    service: ProductionService = Depends(get_production_service),
) -> DayOccupancyResponse:
    result = service.get_day_occupancy(exclude_plan_id=exclude_plan_id)
    return DayOccupancyResponse(**result)


@router.get("/kp-candidates", response_model=KpCandidatesResponse)
def get_kp_candidates(
    _user: dict = Depends(require_roles("admin", "production")),
    service: ProductionService = Depends(get_production_service),
    scope: Literal["plan", "in_work"] = Query("plan"),
    date_from: date | None = Query(None, alias="from"),
    date_to: date | None = Query(None, alias="to"),
) -> KpCandidatesResponse:
    if (date_from is None) != (date_to is None):
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Параметры from и to нужно передавать вместе.",
        )
    if date_from is not None and date_to is not None:
        if date_from > date_to:
            raise_bad_request_client_error(
                ValueError("from must be <= to"),
                where="production.get_kp_candidates",
                detail="Параметр from не может быть позже to.",
            )
        if (date_to - date_from).days > 365:
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail="Диапазон дат не может превышать 366 дней.",
            )
    result = service.list_kp_candidates(
        scope=scope, from_date=date_from, to_date=date_to
    )
    return KpCandidatesResponse(**result)


@router.get("/days/{target_date}", response_model=DayViewDetailResponse)
def get_day_view(
    target_date: date,
    _user: dict = Depends(require_roles("admin", "production")),
    service: ProductionService = Depends(get_production_service),
) -> DayViewDetailResponse:
    data = service.get_day_view_detailed(target_date.isoformat())
    if not data:
        raise HTTPException(status_code=404, detail="Day not found")
    return DayViewDetailResponse(**data)


@router.post("/days/{target_date}/complete")
def complete_day(
    target_date: date,
    payload: CompleteProductionDayRequest,
    user: dict = Depends(require_roles("admin", "production")),
    service: ProductionService = Depends(get_production_service),
) -> dict:
    actor = user.get("email") or user.get("login") or user.get("user_id")
    day_iso = target_date.isoformat()
    try:
        return service.complete_day(
            plan_id=payload.plan_id,
            target_date=day_iso,
            rejected_plates=[item.model_dump() for item in payload.rejected_plates],
            actor=str(actor) if actor else None,
            expected_version=payload.expected_version,
        )
    except ProductionRestValidationError as exc:
        raise_structured_error(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            code=ERROR_CODE_REST_VALIDATION_FAILED,
            message=str(exc),
            details={"plan_id": exc.plan_id, "plate": exc.plate_context},
            where="production.complete_day",
        )
    except ProductionRestDbError as exc:
        raise_unexpected_server_error(exc, where="production.complete_day")
    except PlanVersionConflict as exc:
        raise_structured_error(
            status_code=status.HTTP_409_CONFLICT,
            code=ERROR_CODE_PLAN_VERSION_CONFLICT,
            message=MSG_PLAN_VERSION_CONFLICT,
            details={
                "plan_id": exc.plan_id,
                "expected_version": exc.expected_version,
            },
            where="production.complete_day",
        )
    except ProductionCompletionError as exc:
        if getattr(exc, "code", None) == "day_already_completed":
            raise_structured_error(
                status_code=status.HTTP_409_CONFLICT,
                code="day_already_completed",
                message=str(exc),
                where="production.complete_day",
            )
        raise_unprocessable_client_error(exc, where="production.complete_day")


@router.get("/days/{target_date}/documents/schema")
async def download_day_schema(
    target_date: date,
    background_tasks: BackgroundTasks,
    plate_order_ctx: PlateOrderContext = Depends(get_plate_order_context),
    _user: dict = Depends(require_roles("admin", "production")),
) -> FileResponse:
    day_iso = target_date.isoformat()
    try:
        pdf_path, cleanup_dir = await generate_day_schema(
            day_iso,
            plate_order_ctx=plate_order_ctx,
        )
    except DayDocumentsError as exc:
        raise_not_found_client_error(
            exc,
            where="production.download_day_schema",
            detail=MSG_DAY_NOT_FOUND,
        )
    except Exception as exc:
        logger.exception("[production/day-schema] ошибка: %s", exc)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Не удалось сформировать схему дорожек.",
        ) from exc
    background_tasks.add_task(make_cleanup_callback(cleanup_dir))
    return FileResponse(
        path=str(pdf_path),
        media_type="application/pdf",
        filename=f"Схема_{day_iso}.pdf",
    )


@router.get("/days/{target_date}/documents/breakdown")
async def download_day_breakdown(
    target_date: date,
    background_tasks: BackgroundTasks,
    plate_order_ctx: PlateOrderContext = Depends(get_plate_order_context),
    _user: dict = Depends(require_roles("admin", "production")),
) -> FileResponse:
    day_iso = target_date.isoformat()
    try:
        xlsx_path, cleanup_dir = await generate_day_breakdown(
            day_iso,
            plate_order_ctx=plate_order_ctx,
        )
    except DayDocumentsError as exc:
        raise_not_found_client_error(
            exc,
            where="production.download_day_breakdown",
            detail=MSG_DAY_NOT_FOUND,
        )
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Не удалось сформировать детальную разбивку.",
        ) from exc
    background_tasks.add_task(make_cleanup_callback(cleanup_dir))
    return FileResponse(
        path=str(xlsx_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"Детальная_разбивка_{day_iso}.xlsx",
    )


@router.get("/days/{target_date}/documents/formovka")
async def download_day_formovka(
    target_date: date,
    background_tasks: BackgroundTasks,
    _user: dict = Depends(require_roles("admin", "production")),
) -> FileResponse:
    day_iso = target_date.isoformat()
    try:
        zip_path, cleanup_dir = await generate_day_formovka(day_iso)
    except DayDocumentsError as exc:
        raise_not_found_client_error(
            exc,
            where="production.download_day_formovka",
            detail=MSG_DAY_NOT_FOUND,
        )
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Не удалось сформировать файлы формовки.",
        ) from exc
    background_tasks.add_task(make_cleanup_callback(cleanup_dir))
    return FileResponse(
        path=str(zip_path),
        media_type="application/zip",
        filename=f"Формовка_{day_iso}.zip",
    )


@router.get("/candidates")
def list_candidates(
    limit: int = Query(500, ge=1, le=500),
    _user: dict = Depends(require_roles("admin", "production")),
    service: ProductionService = Depends(get_production_service),
) -> dict:
    items = service.load_candidates_for_plan(limit=limit)
    return {"items": items, "count": len(items)}


@router.get("/sgp/plates", response_model=SgpPlatesResponse)
def list_sgp_plates(
    filter: str = "all",
    _user: dict = Depends(require_roles("admin", "production")),
    sgp: SgpService = Depends(get_sgp_service),
) -> SgpPlatesResponse:
    if filter not in ("all", "linked", "unlinked"):
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="filter must be all|linked|unlinked",
        )
    return sgp.list_plates(filter=filter)  # type: ignore[arg-type]


@router.get("/sgp/free-plates", response_model=SgpFreePlatesResponse)
def list_sgp_free_plates(
    plate_name: str | None = None,
    length_m: float | None = None,
    width_m: float | None = None,
    load_class: int | None = None,
    _user: dict = Depends(require_roles("admin", "production")),
    sgp: SgpService = Depends(get_sgp_service),
) -> SgpFreePlatesResponse:
    return sgp.free_plates(
        plate_name=plate_name,
        length_m=length_m,
        width_m=width_m,
        load_class=load_class,
    )


@router.post("/sgp/plates/{sgp_id}/unlink", response_model=SgpMutationResponse)
def unlink_sgp_plate(
    sgp_id: int,
    payload: SgpUnlinkRequest,
    user: dict = Depends(require_roles("admin", "production")),
    sgp: SgpService = Depends(get_sgp_service),
) -> SgpMutationResponse:
    actor = user.get("email") or user.get("login") or user.get("user_id")
    try:
        return sgp.unlink(sgp_id, payload.qty, actor=str(actor) if actor else None)
    except SgpError as exc:
        raise_structured_error(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            code=exc.code,
            message=str(exc),
            details={"sgp_id": sgp_id},
            where="production.unlink_sgp_plate",
        )


@router.post("/sgp/plates/{sgp_id}/relink", response_model=SgpMutationResponse)
def relink_sgp_plate(
    sgp_id: int,
    payload: SgpRelinkRequest,
    user: dict = Depends(require_roles("admin", "production")),
    sgp: SgpService = Depends(get_sgp_service),
) -> SgpMutationResponse:
    actor = user.get("email") or user.get("login") or user.get("user_id")
    try:
        return sgp.relink(
            sgp_id,
            target_kp_id=payload.target_kp_id,
            qty=payload.qty,
            actor=str(actor) if actor else None,
        )
    except SgpError as exc:
        raise_structured_error(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            code=exc.code,
            message=str(exc),
            details={"sgp_id": sgp_id, "target_kp_id": payload.target_kp_id},
            where="production.relink_sgp_plate",
        )


@router.get("/work-calendar")
def get_work_calendar(
    _user: dict = Depends(require_roles("admin", "production")),
    service: ProductionService = Depends(get_production_service),
) -> dict:
    return service.get_work_calendar()


@router.put("/work-calendar")
def save_work_calendar(
    payload: SaveWorkCalendarRequest,
    _user: dict = Depends(require_roles("admin", "production")),
    service: ProductionService = Depends(get_production_service),
) -> dict:
    return service.save_work_calendar(payload.model_dump())


def _record_build_exclusions(
    payload: BuildPlanRequest,
    result: dict,
    user: dict,
    service: ProductionService,
) -> None:
    """Persist wizard exclusions after a successful build. Does not block ILP."""
    if not payload.exclusions:
        return
    plan = result.get("plan") if isinstance(result, dict) else None
    plan_id = plan.get("id") if isinstance(plan, dict) else None
    if not plan_id:
        return
    actor = user.get("username") or user.get("id")
    PromiseService(db_path=service.kp_repository.db_path).record_plan_exclusions(
        plan_id=str(plan_id),
        exclusions=payload.exclusions,
        excluded_by="" if actor is None else str(actor),
        user=user,
    )


def _date_range_inclusive(start: date, end: date) -> list[date]:
    days: list[date] = []
    current = start
    while current <= end:
        days.append(current)
        current += timedelta(days=1)
    return days


@router.get("/day-capacity", response_model=DayCapacityMapResponse)
def get_day_capacity(
    date_from: date = Query(..., alias="from"),
    date_to: date = Query(..., alias="to"),
    _user: dict = Depends(require_roles("admin", "production")),
    capacity: ProductionCapacityService = Depends(get_production_capacity_service),
) -> DayCapacityMapResponse:
    if date_from > date_to:
        raise_bad_request_client_error(
            ValueError("from must be <= to"),
            where="production.get_day_capacity",
            detail="Параметр from не может быть позже to.",
        )
    # Inclusive span ≤ 366 days ⇔ (to − from).days ≤ 365.
    if (date_to - date_from).days > 365:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Диапазон дат не может превышать 366 дней.",
        )
    capacity_map = capacity.get_capacity_map(_date_range_inclusive(date_from, date_to))
    return DayCapacityMapResponse(
        capacity={day.isoformat(): max_tracks for day, max_tracks in capacity_map.items()}
    )


@router.put("/day-capacity", response_model=SaveDayCapacityResponse)
def save_day_capacity(
    payload: SaveDayCapacityRequest,
    user: dict = Depends(require_roles("admin", "production")),
    capacity: ProductionCapacityService = Depends(get_production_capacity_service),
) -> SaveDayCapacityResponse:
    try:
        capacity.set_day_capacity(payload.date, payload.max_tracks, user=user)
    except ProductionCapacityError as exc:
        raise_bad_request_client_error(
            exc,
            where="production.save_day_capacity",
            detail=str(exc),
        )
    return SaveDayCapacityResponse(date=payload.date, max_tracks=payload.max_tracks)
