from __future__ import annotations

import logging
import sys
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from app.api.v1.router import router as api_v1_router
from app.core.settings import get_settings
from app.repositories.auth_repository import AuthRepository
from app.security.login_rate_limit import (
    rate_limit_deployment_info,
    warn_if_multi_worker_without_shared_store,
)
from app.services.draft_store import DraftStoreLockTimeout
from app.web.router import router as web_router
from core.logging_config import setup_logging


@asynccontextmanager
async def lifespan(app: FastAPI):
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    settings = get_settings()
    setup_logging(level=logging.INFO, log_dir=settings.logs_dir, log_filename="backend.log")
    from core import kp_db

    kp_db.ensure_schema(str(settings.plita_db_path))
    AuthRepository(str(settings.plita_db_path)).init_schema()
    warn_if_multi_worker_without_shared_store()
    app.state.settings = settings
    yield


def create_app() -> FastAPI:
    settings = get_settings()
    app = FastAPI(
        title=settings.app_name,
        debug=settings.app_debug,
        lifespan=lifespan,
    )
    app.add_middleware(
        CORSMiddleware,
        allow_origins=settings.cors_allowed_origins,
        allow_credentials=True,
        allow_methods=["GET", "POST", "PATCH", "DELETE", "OPTIONS"],
        allow_headers=["*"],
    )
    # После CORS: внешний слой на входе — отдельное состояние заказа на каждый запрос (S1).
    from app.middleware.plate_runtime_isolation import (
        PlateMutableRuntimeIsolationMiddleware,
    )

    app.add_middleware(PlateMutableRuntimeIsolationMiddleware)

    from app.middleware.csrf import CsrfMiddleware
    from app.middleware.security_headers import SecurityHeadersMiddleware

    app.add_middleware(SecurityHeadersMiddleware)
    app.add_middleware(CsrfMiddleware)

    @app.exception_handler(DraftStoreLockTimeout)
    async def _draft_store_lock_handler(
        _request: Request, _exc: DraftStoreLockTimeout
    ) -> JSONResponse:
        return JSONResponse(
            status_code=503,
            content={
                "detail": "Черновик временно занят другим запросом. Повторите попытку."
            },
        )

    @app.get("/health", tags=["health"])
    def root_health() -> dict:
        return {
            "status": "ok",
            "app": settings.app_name,
            "rate_limiting": rate_limit_deployment_info(),
        }

    app.include_router(api_v1_router, prefix="/api/v1")
    app.include_router(web_router)
    return app


app = create_app()

