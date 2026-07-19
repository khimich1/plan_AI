"""Регистрация всех обработчиков бота"""
from aiogram import Dispatcher, Router

from bot.middleware.auth import BotAuthMiddleware
from bot.middleware.plate_runtime_isolation import PlateMutableRuntimeIsolationMiddleware
from bot.middleware.role import RoleMiddleware
from . import (
    main, instructions, kp, comparison, commercial, export, admin, pb_info, archive,
    production_planning, production_plans_list, production_create, production_calendar,
    production_execution, production_day_view, production_export, production_completion,
    work_calendar_manager
)

_ALL_ROLES = ("admin", "manager", "production")
_ADMIN_MANAGER = ("admin", "manager")
_ADMIN_PRODUCTION = ("admin", "production")
_ADMIN_ONLY = ("admin",)


def _attach_role_middleware(router: Router, *roles: str) -> None:
    middleware = RoleMiddleware(*roles)
    router.message.middleware(middleware)
    router.callback_query.middleware(middleware)


def register_all_handlers(dp: Dispatcher):
    """Регистрируем все роутеры в правильном порядке"""
    dp.update.middleware(BotAuthMiddleware())
    dp.update.middleware(PlateMutableRuntimeIsolationMiddleware())

    _attach_role_middleware(main.router, *_ALL_ROLES)
    _attach_role_middleware(instructions.router, *_ALL_ROLES)

    for router in (
        kp.router,
        comparison.router,
        commercial.router,
        archive.router,
        pb_info.router,
        export.router,
    ):
        _attach_role_middleware(router, *_ADMIN_MANAGER)

    for router in (
        production_planning.router,
        production_plans_list.router,
        production_create.router,
        production_calendar.router,
        production_execution.router,
        production_day_view.router,
        production_export.router,
        production_completion.router,
        work_calendar_manager.router,
    ):
        _attach_role_middleware(router, *_ADMIN_PRODUCTION)

    _attach_role_middleware(admin.router, *_ADMIN_ONLY)

    dp.include_router(main.router)
    dp.include_router(instructions.router)
    dp.include_router(kp.router)
    dp.include_router(comparison.router)
    dp.include_router(commercial.router)
    dp.include_router(archive.router)
    dp.include_router(production_planning.router)
    dp.include_router(production_plans_list.router)
    dp.include_router(production_create.router)
    dp.include_router(production_calendar.router)
    dp.include_router(production_execution.router)
    dp.include_router(production_day_view.router)
    dp.include_router(production_export.router)
    dp.include_router(production_completion.router)
    dp.include_router(work_calendar_manager.router)
    dp.include_router(pb_info.router)
    dp.include_router(export.router)
    dp.include_router(admin.router)
