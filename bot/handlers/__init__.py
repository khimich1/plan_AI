"""Регистрация всех обработчиков бота"""
from aiogram import Dispatcher
from . import main, kp, comparison, commercial, export, admin, production_new


def register_all_handlers(dp: Dispatcher):
    """Регистрируем все роутеры в правильном порядке"""
    dp.include_router(main.router)
    dp.include_router(kp.router)
    dp.include_router(comparison.router)
    dp.include_router(commercial.router)
    dp.include_router(production_new.router)
    dp.include_router(export.router)
    dp.include_router(admin.router)

