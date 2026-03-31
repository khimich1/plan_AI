from __future__ import annotations

from fastapi import APIRouter

from app.api.v1.endpoints import auth, commercial, health, managers, production

router = APIRouter()
router.include_router(health.router)
router.include_router(auth.router)
router.include_router(managers.router)
router.include_router(commercial.router)
router.include_router(production.router)

