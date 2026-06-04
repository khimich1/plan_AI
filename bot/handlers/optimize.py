"""Обработчики оптимизации резов"""
import sys
from pathlib import Path

from aiogram import Router, F
from aiogram.types import Message
from aiogram.filters import Command

# Добавляем корень проекта в sys.path
BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

import core.config_and_data as cfg
from core.optimization import optimize_with_cascading_longitudinal_cuts
from core.plate_order_context import PlateOrderContext, run_in_order_context

from ..dependencies.plate_context import PlateOrderContextDep
from ..keyboards import main_menu_kb

router = Router()


@router.message(Command("optimize"), PlateOrderContextDep())
@router.message(F.text == "Оптимизация резов", PlateOrderContextDep())
async def cmd_optimize(message: Message, plate_order_ctx: PlateOrderContext):
    """Оптимизация раскроя с каскадными продольными резами"""
    await message.answer("⏳ Выполняю оптимизацию раскроя с учётом вторичных резов...")
    
    try:
        # Собираем заказы из текущей конфигурации
        orders = {}
        if cfg.PLATES_0_32:
            orders[320] = len(cfg.PLATES_0_32)
        if cfg.PLATES_0_46:
            orders[460] = len(cfg.PLATES_0_46)
        if cfg.PLATES_0_70:
            orders[700] = len(cfg.PLATES_0_70)
        if cfg.PLATES_0_72:
            orders[720] = len(cfg.PLATES_0_72)
        if cfg.PLATES_0_86:
            orders[860] = len(cfg.PLATES_0_86)
        if cfg.PLATES_0_88:
            orders[880] = len(cfg.PLATES_0_88)
        if cfg.PLATES_0_74:
            orders[740] = len(cfg.PLATES_0_74)
        if cfg.PLATES_0_48:
            orders[480] = len(cfg.PLATES_0_48)
        if cfg.PLATES_0_50:
            orders[500] = len(cfg.PLATES_0_50)
        if cfg.PLATES_0_34:
            orders[340] = len(cfg.PLATES_0_34)
        
        if not orders:
            await message.answer(
                "⚠️ Нет данных для оптимизации.\n"
                "Сначала создайте заказ через '📝 Создать КП'.",
                reply_markup=main_menu_kb()
            )
            return
        
        result = await run_in_order_context(
            plate_order_ctx,
            optimize_with_cascading_longitudinal_cuts,
            orders,
        )
        
        if result and result.get('total_plates', 0) > 0:
            # Формируем красивый ответ
            response = "✅ **Оптимизация завершена!**\n\n"
            response += f"📊 **Результат:**\n"
            response += f"• Плит потребуется: **{result['total_plates']} шт**\n"
            response += f"• Стоимость: **{result['total_cost']:,} ₽**\n".replace(',', ' ')
            response += f"• Отходы по ширине: **{result.get('waste_width', 0)} мм**\n\n"
            
            if result.get('primary_cuts'):
                response += "🔹 **Первичные резы:**\n"
                for cut in result['primary_cuts']:
                    response += f"  • {cut['qty']} плит → {cut['width']} мм + остаток {cut['rest']} мм\n"
            
            if result.get('secondary_cuts'):
                response += f"\n🔸 **Вторичные резы (из остатков):**\n"
                for cut in result['secondary_cuts']:
                    if cut.get('pieces', 1) > 1:
                        response += f"  • {cut['qty']} остатков {cut['source']} мм → {cut['pieces']} частей по {cut['cuts'][0]} мм\n"
                    else:
                        cuts_str = ' + '.join(str(c) for c in cut['cuts'])
                        response += f"  • {cut['qty']} остатков {cut['source']} мм → {cuts_str} мм\n"
            
            response += "\n💡 **Преимущества:**\n"
            response += "• Минимум плит\n"
            response += "• Остатки используются повторно\n"
            response += "• Меньше отходов\n"
            
            await message.answer(response, parse_mode="Markdown", reply_markup=main_menu_kb())
        else:
            await message.answer(
                "❌ Не удалось выполнить оптимизацию.\n"
                "Проверьте корректность данных.",
                reply_markup=main_menu_kb()
            )
    
    except Exception as e:
        await message.answer(
            f"❌ Ошибка при оптимизации: {str(e)}\n\n"
            f"Убедитесь, что библиотека PuLP установлена.",
            reply_markup=main_menu_kb()
        )

