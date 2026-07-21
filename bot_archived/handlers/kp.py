"""Обработчики команд для работы с планами"""
import os
import logging
import sys
from pathlib import Path

from aiogram import Router
from aiogram.types import Message, FSInputFile
from aiogram.filters import Command

# Добавляем корень проекта в sys.path
BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core.plate_order_context import PlateOrderContext, run_in_order_context
from core.visualization import visualize_plan
from ..bot_config import OUTPUTS_DIR_STR
from ..dependencies.plate_context import PlateOrderContextDep

logger = logging.getLogger(__name__)

router = Router()

@router.message(Command("build_plan"), PlateOrderContextDep())
async def cmd_build_plan(message: Message, plate_order_ctx: PlateOrderContext):
    """Обработчик команды /build_plan"""
    await message.answer("⏳ Выполняю расчёт дорожки, подожди немного...")
    
    try:
        result_paths = await run_in_order_context(
            plate_order_ctx,
            visualize_plan,
            OUTPUTS_DIR_STR,
        )
        
        if isinstance(result_paths, tuple) and len(result_paths) >= 2:
            png_path, pdf_path = result_paths
            
            # Ищем дополнительные файлы
            # Извлекаем timestamp (всё после "КЗ_")
            base = os.path.basename(png_path)
            if 'КЗ_' in base:
                timestamp = base.split('КЗ_', 1)[-1].replace('.png', '')
            else:
                # Fallback: последняя часть после последнего подчеркивания
                timestamp = base.rsplit('_', 1)[-1].replace('.png', '')
            
            csv_path = os.path.join(OUTPUTS_DIR_STR, f'Раскладка_Дорожка_1_{timestamp}.csv')
            xlsx_path = os.path.join(OUTPUTS_DIR_STR, f'Ведомость_Дорожка_1_{timestamp}.xlsx')
            breakdown_path = os.path.join(OUTPUTS_DIR_STR, f'Детальная_разбивка_Дорожка_1_{timestamp}.xlsx')
            xlsx_smeta_path = os.path.join(OUTPUTS_DIR_STR, f'Смета_Дорожка_1_{timestamp}.xlsx')
            
            await message.answer("✅ Готово! Отправляю файлы:")
            
            # Отправляем изображение как документ, чтобы избежать PHOTO_INVALID_DIMENSIONS
            if os.path.exists(png_path):
                await message.answer_document(FSInputFile(png_path))
            
            # Отправляем документы
            if os.path.exists(pdf_path):
                await message.answer_document(FSInputFile(pdf_path))
            
            if os.path.exists(xlsx_path):
                await message.answer_document(FSInputFile(xlsx_path))
            
            if os.path.exists(xlsx_smeta_path):
                await message.answer_document(FSInputFile(xlsx_smeta_path))
            
            if os.path.exists(breakdown_path):
                await message.answer_document(FSInputFile(breakdown_path))
            
            if os.path.exists(csv_path):
                await message.answer_document(FSInputFile(csv_path))
            
            await message.answer(
                "📋 **Результаты расчёта готовы!**\n\n"
                "• Схема раскладки сохранена\n"
                "• Ведомость материалов готова\n"
                "• Смета стоимости рассчитана\n"
                "• Все файлы экспортированы"
            )
        else:
            await message.answer("❌ Ошибка при расчёте плана")
            
    except Exception as e:
        logger.exception(f"Ошибка в /build_plan: {e}")
        await message.answer(
            "❌ Не удалось построить план.\n\n"
            "Попробуйте ещё раз через минуту.\n"
            "Если повторяется — откройте файл logs/bot.log и пришлите последние строки.",
            parse_mode=None
        )

