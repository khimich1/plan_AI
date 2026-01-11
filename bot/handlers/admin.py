"""Обработчики административных команд (удаление КП и т.д.)"""
import sys
from pathlib import Path

from aiogram import Router, F
from aiogram.types import Message, CallbackQuery
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

# Добавляем корень проекта в sys.path
BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent
sys.path.insert(0, str(PROJECT_ROOT))

from core import kp_db
from ..keyboards import main_menu_kb, db_management_kb, db_clear_confirm_kb

router = Router()


# Состояния для подтверждения удаления
class AdminStates(StatesGroup):
    waiting_delete_confirmation = State()
    waiting_clear_all_confirmation = State()


@router.message(Command("delete_kp"))
async def cmd_delete_kp(message: Message, state: FSMContext):
    """
    Удаляет КП из базы данных по номеру.
    Использование: /delete_kp 5
    """
    args = message.text.split()
    
    if len(args) != 2:
        await message.answer(
            "❌ Неверный формат команды!\n\n"
            "Использование: /delete_kp <номер_КП>\n"
            "Пример: /delete_kp 5",
            reply_markup=main_menu_kb()
        )
        return
    
    try:
        kp_id = int(args[1])
    except ValueError:
        await message.answer(
            "❌ Номер КП должен быть числом!",
            reply_markup=main_menu_kb()
        )
        return
    
    # Получаем информацию о КП перед удалением (для подтверждения)
    kp_info = kp_db.get_kp_by_id(kp_id)
    
    if not kp_info:
        await message.answer(
            f"❌ КП #{kp_id} не найдено в базе данных.",
            reply_markup=main_menu_kb()
        )
        return
    
    # Сохраняем kp_id в состояние для подтверждения
    await state.update_data(delete_kp_id=kp_id)
    await state.set_state(AdminStates.waiting_delete_confirmation)
    
    # Показываем информацию о КП
    plates_count = len(kp_info.get('plates', []))
    total_amount = kp_info.get('total_amount', 0)
    
    await message.answer(
        f"⚠️ **ВНИМАНИЕ!** Вы собираетесь удалить КП #{kp_id}\n\n"
        f"📋 Информация:\n"
        f"  • Дата: {kp_info.get('creation_date', 'Не указана')}\n"
        f"  • Клиент: {kp_info.get('customer_name', 'Не указан')}\n"
        f"  • Менеджер: {kp_info.get('manager_name', 'Не указан')}\n"
        f"  • Сумма: {total_amount:,.2f} ₽ (с НДС)\n"
        f"  • Плит: {plates_count} шт\n"
        f"  • Статус: {kp_info.get('status', 'не указан')}\n\n"
        f"⚠️ **Это действие необратимо!**\n"
        f"Все данные (плиты, файлы, метаданные) будут удалены.\n\n"
        f"Для подтверждения отправьте: **ДА**\n"
        f"Для отмены отправьте: **НЕТ**",
        parse_mode="Markdown"
    )


@router.message(AdminStates.waiting_delete_confirmation)
async def confirm_delete_kp(message: Message, state: FSMContext):
    """Подтверждение удаления КП"""
    confirmation = message.text.strip().upper()
    
    if confirmation not in ['ДА', 'YES', 'Y']:
        await state.clear()
        await message.answer(
            "❌ Удаление отменено.",
            reply_markup=main_menu_kb()
        )
        return
    
    # Получаем kp_id из состояния
    data = await state.get_data()
    kp_id = data.get('delete_kp_id')
    
    if not kp_id:
        await state.clear()
        await message.answer(
            "❌ Ошибка: не удалось получить номер КП.",
            reply_markup=main_menu_kb()
        )
        return
    
    # Удаляем КП
    success = kp_db.delete_kp_by_id(kp_id)
    
    await state.clear()
    
    if success:
        await message.answer(
            f"✅ **КП #{kp_id} успешно удалено из базы данных!**\n\n"
            f"Все связанные данные (плиты, файлы, метаданные) также удалены.",
            parse_mode="Markdown",
            reply_markup=main_menu_kb()
        )
    else:
        await message.answer(
            f"❌ Не удалось удалить КП #{kp_id}.\n"
            f"Возможно, оно уже было удалено ранее.",
            reply_markup=main_menu_kb()
        )


@router.message(Command("list_kp"))
async def cmd_list_kp(message: Message):
    """
    Показывает список всех КП в базе данных.
    Использование: /list_kp
    """
    try:
        # Получаем все КП со статусом "в работе"
        kp_list_active = kp_db.get_all_kp_by_status('в работе')
        kp_list_done = kp_db.get_all_kp_by_status('выполнено')
        kp_list_rejected = kp_db.get_all_kp_by_status('отклонено')
        
        response = "📋 **Список КП в базе данных:**\n\n"
        
        if kp_list_active:
            response += "🟢 **В работе:**\n"
            for kp in kp_list_active:
                response += (
                    f"  • КП #{kp['kp_id']} - {kp.get('customer_name', 'Без имени')} - "
                    f"{kp.get('total_amount', 0):,.2f} ₽\n"
                )
            response += "\n"
        
        if kp_list_done:
            response += "✅ **Выполнено:**\n"
            for kp in kp_list_done:
                response += (
                    f"  • КП #{kp['kp_id']} - {kp.get('customer_name', 'Без имени')} - "
                    f"{kp.get('total_amount', 0):,.2f} ₽\n"
                )
            response += "\n"
        
        if kp_list_rejected:
            response += "❌ **Отклонено:**\n"
            for kp in kp_list_rejected:
                response += (
                    f"  • КП #{kp['kp_id']} - {kp.get('customer_name', 'Без имени')} - "
                    f"{kp.get('total_amount', 0):,.2f} ₽\n"
                )
            response += "\n"
        
        if not (kp_list_active or kp_list_done or kp_list_rejected):
            response += "📭 База данных пуста.\n\n"
        
        response += "💡 Для удаления КП используйте: /delete_kp <номер>"
        
        await message.answer(response, parse_mode="Markdown", reply_markup=main_menu_kb())
    
    except Exception as e:
        await message.answer(
            f"❌ Ошибка при получении списка КП: {str(e)}",
            reply_markup=main_menu_kb()
        )


@router.message(Command("clear_all_kp"))
async def cmd_clear_all_kp(message: Message, state: FSMContext):
    """
    Полностью очищает все КП из базы данных.
    Использование: /clear_all_kp
    """
    try:
        # Подсчитываем количество КП перед удалением
        kp_list_active = kp_db.get_all_kp_by_status('в работе')
        kp_list_done = kp_db.get_all_kp_by_status('выполнено')
        kp_list_rejected = kp_db.get_all_kp_by_status('отклонено')
        kp_list_waiting = kp_db.get_all_kp_by_status('в ожидании')
        
        total_count = len(kp_list_active) + len(kp_list_done) + len(kp_list_rejected) + len(kp_list_waiting)
        
        if total_count == 0:
            await message.answer(
                "ℹ️ База данных уже пуста. Нечего удалять.",
                reply_markup=main_menu_kb()
            )
            return
        
        # Сохраняем информацию в состояние
        await state.update_data(clear_all_count=total_count)
        await state.set_state(AdminStates.waiting_clear_all_confirmation)
        
        # Показываем предупреждение
        await message.answer(
            f"⚠️ **КРИТИЧЕСКОЕ ВНИМАНИЕ!**\n\n"
            f"Вы собираетесь **ПОЛНОСТЬЮ ОЧИСТИТЬ** всю базу данных с КП!\n\n"
            f"📊 Будет удалено:\n"
            f"  • КП в работе: {len(kp_list_active)}\n"
            f"  • КП выполнено: {len(kp_list_done)}\n"
            f"  • КП отклонено: {len(kp_list_rejected)}\n"
            f"  • КП в ожидании: {len(kp_list_waiting)}\n"
            f"  • **ВСЕГО: {total_count} КП**\n\n"
            f"🗑️ Будут удалены:\n"
            f"  • Все записи о КП\n"
            f"  • Все плиты в КП\n"
            f"  • Все файлы XLSX\n"
            f"  • Все метаданные\n\n"
            f"⚠️ **ЭТО ДЕЙСТВИЕ НЕОБРАТИМО!**\n"
            f"После очистки база данных будет полностью пуста.\n\n"
            f"Для подтверждения отправьте: **ДА**\n"
            f"Для отмены отправьте: **НЕТ**",
            parse_mode="Markdown"
        )
    
    except Exception as e:
        await message.answer(
            f"❌ Ошибка при проверке базы данных: {str(e)}",
            reply_markup=main_menu_kb()
        )


@router.message(AdminStates.waiting_clear_all_confirmation)
async def confirm_clear_all_kp(message: Message, state: FSMContext):
    """Подтверждение полной очистки БД"""
    confirmation = message.text.strip().upper()
    
    if confirmation not in ['ДА', 'YES', 'Y']:
        await state.clear()
        await message.answer(
            "❌ Очистка отменена.",
            reply_markup=main_menu_kb()
        )
        return
    
    # Получаем количество из состояния
    data = await state.get_data()
    expected_count = data.get('clear_all_count', 0)
    
    try:
        # Выполняем полную очистку
        result = kp_db.clear_all_kp()
        
        await state.clear()
        
        # Формируем отчёт
        total_deleted = result.get('kp_offers', 0)
        
        await message.answer(
            f"✅ **База данных полностью очищена!**\n\n"
            f"📊 Удалено:\n"
            f"  • КП: {result.get('kp_offers', 0)}\n"
            f"  • Записей плит: {result.get('kp_plates', 0)}\n"
            f"  • Файлов: {result.get('kp_files', 0)}\n"
            f"  • Метаданных: {result.get('kp_meta', 0)}\n\n"
            f"🔄 Счётчики AUTOINCREMENT сброшены.\n"
            f"Новые КП будут начинаться с номера 1.",
            parse_mode="Markdown",
            reply_markup=main_menu_kb()
        )
    
    except Exception as e:
        await state.clear()
        await message.answer(
            f"❌ Ошибка при очистке базы данных: {str(e)}\n\n"
            f"Попробуйте снова позже.",
            reply_markup=main_menu_kb()
        )
        import traceback
        traceback.print_exc()


# ==================== НОВОЕ МЕНЮ УПРАВЛЕНИЯ БД ====================

@router.message(F.text == "⚙️ Управление БД")
async def btn_db_management(message: Message):
    """Обработчик кнопки 'Управление БД' из главного меню"""
    await message.answer(
        "⚙️ **Управление базой данных**\n\n"
        "Выберите действие:",
        parse_mode="Markdown",
        reply_markup=db_management_kb()
    )


@router.callback_query(F.data == "db_stats")
async def show_db_stats(callback: CallbackQuery):
    """Показывает статистику БД"""
    try:
        stats = kp_db.get_db_stats()
        
        response = "📊 **Статистика базы данных**\n\n"
        response += f"📋 **КП:**\n"
        response += f"  • Всего: {stats['kp_total']}\n"
        response += f"  • В работе: {stats['kp_in_work']}\n"
        response += f"  • Выполнено: {stats['kp_completed']}\n\n"
        response += f"📦 **Плиты:**\n"
        response += f"  • В работе: {stats['plates_in_work']}\n"
        response += f"  • Выполнено: {stats['plates_completed']}\n\n"
        response += f"🔧 **Остатки от резки:**\n"
        response += f"  • Всего: {stats['plate_rests']}\n"
        
        await callback.message.answer(
            response,
            parse_mode="Markdown",
            reply_markup=db_management_kb()
        )
        await callback.answer()
    
    except Exception as e:
        await callback.message.answer(
            f"❌ Ошибка при получении статистики: {str(e)}",
            reply_markup=db_management_kb()
        )
        await callback.answer()


@router.callback_query(F.data == "db_clear_all")
async def request_db_clear(callback: CallbackQuery):
    """Запрос подтверждения полной очистки БД"""
    try:
        stats = kp_db.get_db_stats()
        
        total = (stats['kp_total'] + stats['plates_in_work'] + 
                stats['plates_completed'] + stats['plate_rests'])
        
        if total == 0:
            await callback.message.answer(
                "ℹ️ База данных уже пуста. Нечего удалять.",
                reply_markup=db_management_kb()
            )
            await callback.answer()
            return
        
        warning = (
            "🔴 **КРИТИЧЕСКОЕ ПРЕДУПРЕЖДЕНИЕ!**\n\n"
            "Вы собираетесь **ПОЛНОСТЬЮ УДАЛИТЬ ВСЕ ДАННЫЕ** из базы!\n\n"
            f"📊 Будет удалено:\n"
            f"  • КП: {stats['kp_total']} шт\n"
            f"  • Плиты в работе: {stats['plates_in_work']} шт\n"
            f"  • Выполненные плиты: {stats['plates_completed']} шт\n"
            f"  • Остатки: {stats['plate_rests']} шт\n"
            f"  • **ВСЕГО ЗАПИСЕЙ: {total}**\n\n"
            f"⚠️ **ЭТО ДЕЙСТВИЕ НЕОБРАТИМО!**\n"
            f"Восстановить данные будет невозможно!\n\n"
            f"Вы уверены, что хотите продолжить?"
        )
        
        await callback.message.answer(
            warning,
            parse_mode="Markdown",
            reply_markup=db_clear_confirm_kb()
        )
        await callback.answer()
    
    except Exception as e:
        await callback.message.answer(
            f"❌ Ошибка при проверке БД: {str(e)}",
            reply_markup=db_management_kb()
        )
        await callback.answer()


@router.callback_query(F.data == "db_clear_confirmed")
async def confirm_db_clear(callback: CallbackQuery):
    """Подтверждение и выполнение полной очистки БД"""
    try:
        # Выполняем полную очистку
        result = kp_db.clear_all_plates_data()
        
        report = (
            "✅ **БАЗА ДАННЫХ ПОЛНОСТЬЮ ОЧИЩЕНА!**\n\n"
            f"📊 Удалено:\n"
            f"  • КП: {result['kp_offers']}\n"
            f"  • Плиты в работе: {result['kp_plates']}\n"
            f"  • Выполненные плиты: {result['completed_plates']}\n"
            f"  • Остатки: {result['plate_rests']}\n"
            f"  • Файлы: {result['kp_files']}\n"
            f"  • Метаданные: {result['kp_meta']}\n"
            f"  • **ВСЕГО: {result['total']} записей**\n\n"
            f"🔄 Счётчики сброшены. Новые КП начнутся с #1.\n"
            f"База данных теперь пуста и готова к новым заказам."
        )
        
        await callback.message.answer(
            report,
            parse_mode="Markdown",
            reply_markup=main_menu_kb()
        )
        await callback.answer("✅ База очищена!")
    
    except Exception as e:
        await callback.message.answer(
            f"❌ Ошибка при очистке БД: {str(e)}\n\n"
            f"Попробуйте снова позже.",
            reply_markup=db_management_kb()
        )
        await callback.answer("❌ Ошибка!")
        import traceback
        traceback.print_exc()


@router.callback_query(F.data == "db_clear_cancel")
async def cancel_db_clear(callback: CallbackQuery):
    """Отмена очистки БД"""
    await callback.message.answer(
        "❌ Очистка отменена. База данных не изменена.",
        reply_markup=db_management_kb()
    )
    await callback.answer()


@router.callback_query(F.data == "db_back_to_menu")
async def back_to_menu(callback: CallbackQuery):
    """Возврат в главное меню"""
    await callback.message.answer(
        "Возврат в главное меню",
        reply_markup=main_menu_kb()
    )
    await callback.answer()

