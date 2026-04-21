from __future__ import annotations

import logging
import sys
from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.api.v1.router import router as api_v1_router
from app.core.settings import get_settings
from app.repositories.auth_repository import AuthRepository
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
    AuthRepository(str(settings.plita_db_path)).init_schema()
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

    @app.get("/health", tags=["health"])
    def root_health() -> dict:
        return {"status": "ok", "app": settings.app_name}

    app.include_router(api_v1_router, prefix="/api/v1")
    app.include_router(web_router)
    return app


app = create_app()

