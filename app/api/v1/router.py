from __future__ import annotations

from fastapi import APIRouter

from app.api.v1.endpoints import (
    admin,
    archive,
    auth,
    commercial,
    delivery_schedule,
    gsm,
    health,
    logistics,
    managers,
    notifications,
    offers,
    production,
)

router = APIRouter()
router.include_router(health.router)
router.include_router(auth.router)
router.include_router(managers.router)
router.include_router(commercial.router)
router.include_router(offers.router)
router.include_router(archive.router)
router.include_router(delivery_schedule.router)
router.include_router(production.router)
router.include_router(logistics.router)
router.include_router(gsm.router)
router.include_router(admin.router)
router.include_router(notifications.router)

