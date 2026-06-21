from __future__ import annotations

import logging

from fastapi import APIRouter, Depends, HTTPException, Query, status

from app.core.http_errors import raise_destructive_db_blocked_error
from app.dependencies.auth import get_auth_repository, require_roles
from app.dependencies.services import get_admin_service
from app.repositories.auth_repository import AuthRepository
from app.schemas.admin import DbResetReport, DbStatsResponse, RecoverPlatesResponse
from app.schemas.auth import UsersPageResponse
from app.services.admin_service import AdminService
from core.destructive_db_guard import (
    DestructiveDbOperationBlocked,
    require_destructive_db_reset,
)

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/admin", tags=["admin"])


def _enforce_destructive_db_reset() -> None:
    try:
        require_destructive_db_reset()
    except DestructiveDbOperationBlocked as exc:
        raise_destructive_db_blocked_error(exc, where="admin.destructive_guard")


@router.get("/users", response_model=UsersPageResponse)
def list_users(
    limit: int = Query(default=50, ge=1, le=200),
    offset: int = Query(default=0, ge=0),
    _user: dict = Depends(require_roles("admin")),
    repository: AuthRepository = Depends(get_auth_repository),
) -> UsersPageResponse:
    page = repository.get_users_page(limit=limit, offset=offset)
    return UsersPageResponse.model_validate(page)


@router.get("/db/stats", response_model=DbStatsResponse)
def get_db_stats(
    _user: dict = Depends(require_roles("admin")),
    service: AdminService = Depends(get_admin_service),
) -> DbStatsResponse:
    return service.get_stats()


@router.post("/db/reset/full", response_model=DbResetReport)
def reset_full(
    _user: dict = Depends(require_roles("admin")),
    _guard: None = Depends(_enforce_destructive_db_reset),
    service: AdminService = Depends(get_admin_service),
) -> DbResetReport:
    try:
        return service.reset_full()
    except DestructiveDbOperationBlocked as exc:
        raise_destructive_db_blocked_error(exc, where="admin.reset_full")
    except Exception as exc:
        logger.exception("[admin/reset-full] ошибка полного обнуления")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Не удалось выполнить полное обнуление базы данных.",
        ) from exc


@router.post("/db/reset/kp-only", response_model=DbResetReport)
def reset_kp_only(
    _user: dict = Depends(require_roles("admin")),
    _guard: None = Depends(_enforce_destructive_db_reset),
    service: AdminService = Depends(get_admin_service),
) -> DbResetReport:
    try:
        return service.reset_kp_only()
    except DestructiveDbOperationBlocked as exc:
        raise_destructive_db_blocked_error(exc, where="admin.reset_kp_only")
    except Exception as exc:
        logger.exception("[admin/reset-kp-only] ошибка обнуления таблиц КП")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Не удалось обнулить таблицы коммерческих предложений.",
        ) from exc


@router.post("/db/reset/plans-only", response_model=DbResetReport)
def reset_plans_only(
    _user: dict = Depends(require_roles("admin")),
    _guard: None = Depends(_enforce_destructive_db_reset),
    service: AdminService = Depends(get_admin_service),
) -> DbResetReport:
    try:
        return service.reset_plans_only()
    except DestructiveDbOperationBlocked as exc:
        raise_destructive_db_blocked_error(exc, where="admin.reset_plans_only")
    except Exception as exc:
        logger.exception("[admin/reset-plans-only] ошибка обнуления планов")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Не удалось удалить файлы планов производства.",
        ) from exc


@router.post("/db/reset/calendar-only", response_model=DbResetReport)
def reset_calendar_only(
    _user: dict = Depends(require_roles("admin")),
    _guard: None = Depends(_enforce_destructive_db_reset),
    service: AdminService = Depends(get_admin_service),
) -> DbResetReport:
    try:
        return service.reset_calendar_only()
    except DestructiveDbOperationBlocked as exc:
        raise_destructive_db_blocked_error(exc, where="admin.reset_calendar_only")
    except Exception as exc:
        logger.exception("[admin/reset-calendar-only] ошибка сброса календаря")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Не удалось сбросить производственный календарь.",
        ) from exc


@router.post("/db/recover-plates", response_model=RecoverPlatesResponse)
def recover_stuck_plates(
    _user: dict = Depends(require_roles("admin")),
    service: AdminService = Depends(get_admin_service),
) -> RecoverPlatesResponse:
    try:
        return service.recover_stuck_plates()
    except Exception as exc:
        logger.exception("[admin/recover-plates] ошибка восстановления плит")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Не удалось восстановить застрявшие плиты.",
        ) from exc
