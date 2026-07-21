from __future__ import annotations

from fastapi import APIRouter

from app.dependencies.commercial_draft import check_draft_ownership  # noqa: F401 — test monkeypatch hook
from app.web.legacy_routes import router as legacy_router
from app.web.spa_routes import router as spa_router

router = APIRouter(include_in_schema=False)
router.include_router(legacy_router)
router.include_router(spa_router)
