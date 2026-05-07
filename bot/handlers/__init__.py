"""Регистрация всех обработчиков бота"""
from aiogram import Dispatcher

from bot.middleware.plate_runtime_isolation import PlateMutableRuntimeIsolationMiddleware
from . import (
    main, instructions, kp, comparison, commercial, export, admin, pb_info, archive,
    production_planning, production_plans_list, production_create, production_calendar,
    production_execution, production_day_view, production_export, production_completion,
    work_calendar_manager
)


def register_all_handlers(dp: Dispatcher):
    """Регистрируем все роутеры в правильном порядке"""
    dp.update.middleware(PlateMutableRuntimeIsolationMiddleware())
    dp.include_router(main.router)
    dp.include_router(instructions.router)  # Роутер инструкций
    dp.include_router(kp.router)
    dp.include_router(comparison.router)
    dp.include_router(commercial.router)
    dp.include_router(archive.router)  # Роутер архива
    # Роутеры планирования производства (разделены на модули)
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

