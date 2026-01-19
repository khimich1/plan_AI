"""Обработчики для архива коммерческих предложений"""
import os
import json
from pathlib import Path
from aiogram import Router, F
from aiogram.types import Message, CallbackQuery, FSInputFile, InlineKeyboardMarkup, InlineKeyboardButton
from aiogram.fsm.context import FSMContext

from core import kp_db
from ..keyboards import main_menu_kb, archive_sections_kb, kp_details_kb

# Определяем пути к директориям проекта
BOT_DIR = Path(__file__).parent.parent
PROJECT_ROOT = BOT_DIR.parent

router = Router()


@router.message(F.text == "📁 Архив")
async def btn_archive(message: Message):
    """
    Обработчик кнопки 'Архив' в главном меню.
    
    Простыми словами:
    - Показывает меню выбора раздела архива
    - Два раздела: "📦 В архиве" и "🏭 В производстве"
    """
    await message.answer(
        "📁 Архив коммерческих предложений\n\n"
        "Выберите раздел для просмотра:",
        reply_markup=archive_sections_kb()
    )


@router.callback_query(F.data == "archive_section_archived")
async def show_archived_kp(callback: CallbackQuery):
    """
    Показать КП в архиве (статус "в архиве").
    
    Простыми словами:
    - Получает все КП со статусом "в архиве" из БД
    - Формирует список с кнопками для каждого КП
    - Сортирует по номеру КП (от меньшего к большему)
    """
    try:
        await callback.answer()
    except:
        pass  # Игнорируем ошибку, если callback устарел
    
    # Получаем список КП со статусом "в архиве"
    all_kp = kp_db.get_all_kp_list()
    archived_kp = all_kp.get('archived', [])
    
    if not archived_kp:
        await callback.message.edit_text(
            "📦 В архиве пока нет КП\n\n"
            "Чтобы добавить КП в архив, создайте его через '📝 Создать КП' "
            "и нажмите '📦 В архив' после генерации документов.",
            reply_markup=archive_sections_kb()
        )
        return
    
    # Формируем текст с информацией о КП
    text = f"📦 КП в архиве ({len(archived_kp)} шт.)\n\n"
    
    # Создаём inline кнопки для каждого КП
    buttons = []
    db_path = PROJECT_ROOT / "plita.db"
    
    for kp in archived_kp:
        kp_id = kp['kp_id']
        customer = kp.get('customer_name', 'Без имени')
        total = kp.get('total_amount', 0)
        date = kp.get('creation_date', '')
        
        # Получаем процент выполнения
        completion_info = kp_db.get_kp_completion_percentage(kp_id, str(db_path))
        percentage = completion_info['percentage']
        
        # Обрезаем длинные имена клиентов
        customer_short = customer[:20] + '...' if len(customer) > 20 else customer
        
        # Формируем текст кнопки с процентом выполнения
        buttons.append([
            InlineKeyboardButton(
                text=f"КП №{kp_id} | {customer_short} | {percentage:.0f}% | {total:,.0f}₽",
                callback_data=f"view_kp_{kp_id}"
            )
        ])
    
    # Кнопка "Назад"
    buttons.append([
        InlineKeyboardButton(text="◀️ Назад", callback_data="archive_back_to_sections")
    ])
    
    await callback.message.edit_text(
        text,
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons)
    )


@router.callback_query(F.data == "archive_section_production")
async def show_production_kp(callback: CallbackQuery):
    """
    Показать КП в производстве (статус "в работе").
    
    Простыми словами:
    - Получает все КП со статусом "в работе" из БД
    - Формирует список с кнопками для каждого КП
    - Сортирует по номеру КП (от меньшего к большему)
    """
    try:
        await callback.answer()
    except:
        pass  # Игнорируем ошибку, если callback устарел
    
    # Получаем список КП со статусом "в работе"
    all_kp = kp_db.get_all_kp_list()
    production_kp = all_kp.get('in_production', [])
    
    if not production_kp:
        await callback.message.edit_text(
            "🏭 В производстве пока нет КП\n\n"
            "Чтобы отправить КП в производство, создайте его через '📝 Создать КП' "
            "и нажмите '💾 Сохранить в БД' после генерации документов.",
            reply_markup=archive_sections_kb()
        )
        return
    
    # Формируем текст с информацией о КП
    text = f"🏭 КП в производстве ({len(production_kp)} шт.)\n\n"
    
    # Создаём inline кнопки для каждого КП
    buttons = []
    db_path = PROJECT_ROOT / "plita.db"
    
    for kp in production_kp:
        kp_id = kp['kp_id']
        customer = kp.get('customer_name', 'Без имени')
        total = kp.get('total_amount', 0)
        execution_terms = kp.get('execution_terms', '')
        
        # Получаем процент выполнения
        completion_info = kp_db.get_kp_completion_percentage(kp_id, str(db_path))
        percentage = completion_info['percentage']
        
        # Обрезаем длинные имена клиентов
        customer_short = customer[:20] + '...' if len(customer) > 20 else customer
        
        # Формируем текст кнопки с процентом
        button_text = f"КП №{kp_id} | {customer_short} | {percentage:.0f}% | {total:,.0f}₽"
        if execution_terms:
            button_text += f" | ⏰{execution_terms}"
        
        buttons.append([
            InlineKeyboardButton(
                text=button_text,
                callback_data=f"view_kp_{kp_id}"
            )
        ])
    
    # Кнопка "Назад"
    buttons.append([
        InlineKeyboardButton(text="◀️ Назад", callback_data="archive_back_to_sections")
    ])
    
    await callback.message.edit_text(
        text,
        reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons)
    )


@router.callback_query(F.data.startswith("view_kp_"))
async def view_kp_details(callback: CallbackQuery):
    """
    Показать детали конкретного КП.
    
    Простыми словами:
    - Получает полную информацию о КП из БД
    - Показывает: клиента, менеджера, дату, сумму, статус, список плит
    - Показывает кнопки: Скачать PDF, Скачать XLSX, Удалить КП
    """
    await callback.answer()
    
    kp_id = int(callback.data.split("_")[-1])
    kp_info = kp_db.get_kp_by_id(kp_id)
    
    if not kp_info:
        await callback.message.edit_text(
            "❌ КП не найдено в базе данных",
            reply_markup=archive_sections_kb()
        )
        return
    
    # Формируем детальное описание
    text = f"📋 Коммерческое предложение № {kp_id}\n\n"
    text += f"👤 Клиент: {kp_info.get('customer_name', 'Не указан')}\n"
    text += f"👨‍💼 Менеджер: {kp_info.get('manager_name', 'Не указан')}\n"
    text += f"📅 Дата создания: {kp_info.get('creation_date', 'Не указана')}\n"
    
    # Статус с эмодзи
    status = kp_info.get('status', 'Неизвестен')
    status_emoji = {
        'в архиве': '📦',
        'в работе': '🏭',
        'выполнено': '✅',
        'отклонено': '❌'
    }.get(status, '❓')
    text += f"📊 Статус: {status_emoji} {status}\n"
    
    if kp_info.get('execution_terms'):
        text += f"⏰ Срок выполнения: {kp_info['execution_terms']}\n"
    
    text += f"\n💰 Финансы:\n"
    text += f"  • Сумма без НДС: {kp_info.get('subtotal', 0):,.2f} ₽\n"
    text += f"  • НДС (22%): {kp_info.get('vat_amount', 0):,.2f} ₽\n"
    text += f"  • Итого с НДС: {kp_info.get('total_amount', 0):,.2f} ₽\n"
    
    if kp_info.get('discount_percent', 0) > 0:
        text += f"  • Скидка: {kp_info['discount_percent']}%\n"
    
    text += f"\n📦 Состав заказа ({len(kp_info.get('plates', []))} позиций):\n"
    
    # Список плит (ограничиваем до 10 позиций, чтобы не переполнить сообщение)
    plates = kp_info.get('plates', [])
    for i, plate in enumerate(plates[:10], 1):
        text += f"  {i}. {plate['plate_name']} — {plate['qty']} шт\n"
    
    if len(plates) > 10:
        text += f"  ... и ещё {len(plates) - 10} позиций\n"
    
    await callback.message.edit_text(
        text,
        reply_markup=kp_details_kb(kp_id)
    )


@router.callback_query(F.data.startswith("download_pdf_"))
async def download_pdf(callback: CallbackQuery):
    """
    Скачать PDF файл КП.
    
    Простыми словами:
    - Ищет PDF файл КП на диске или в БД
    - Отправляет его пользователю
    """
    await callback.answer("⏳ Подготавливаю PDF...")
    
    kp_id = int(callback.data.split("_")[-1])
    kp_info = kp_db.get_kp_by_id(kp_id)
    
    if not kp_info or 'file' not in kp_info:
        await callback.message.answer("❌ Информация о файлах не найдена")
        return
    
    # Пытаемся найти PDF файл
    file_path = kp_info['file'].get('file_path')
    
    if file_path:
        # Конвертируем путь XLSX в PDF (меняем расширение)
        pdf_path = file_path.replace('.xlsx', '.pdf')
        
        if os.path.exists(pdf_path):
            await callback.message.answer_document(
                FSInputFile(pdf_path),
                caption=f"📄 КП № {kp_id} (PDF)"
            )
        else:
            await callback.message.answer(
                f"❌ PDF файл не найден по пути:\n{pdf_path}\n\n"
                "Возможно, файл был удалён с диска."
            )
    else:
        await callback.message.answer("❌ Путь к файлу не сохранён в БД")


@router.callback_query(F.data.startswith("download_xlsx_"))
async def download_xlsx(callback: CallbackQuery):
    """
    Скачать XLSX файл КП.
    
    Простыми словами:
    - Извлекает XLSX файл из БД (BLOB)
    - Отправляет его пользователю
    """
    await callback.answer("⏳ Подготавливаю XLSX...")
    
    kp_id = int(callback.data.split("_")[-1])
    
    # Извлекаем XLSX из БД
    xlsx_data = kp_db.get_xlsx_file(kp_id)
    
    if xlsx_data:
        # Сохраняем во временный файл и отправляем
        temp_path = f"/tmp/КП_{kp_id}.xlsx"
        with open(temp_path, 'wb') as f:
            f.write(xlsx_data)
        
        await callback.message.answer_document(
            FSInputFile(temp_path),
            caption=f"📊 КП № {kp_id} (XLSX с формулами)"
        )
        
        # Удаляем временный файл
        try:
            os.remove(temp_path)
        except:
            pass
    else:
        await callback.message.answer(
            "❌ XLSX файл не найден в базе данных\n\n"
            "Возможно, файл не был сохранён при создании КП."
        )


@router.callback_query(F.data.startswith("delete_kp_"))
async def delete_kp_confirm(callback: CallbackQuery, state: FSMContext):
    """
    Подтверждение удаления КП.
    
    Простыми словами:
    - Показывает предупреждение об удалении
    - Просит подтвердить действие
    """
    try:
        await callback.answer()
    except:
        pass  # Игнорируем ошибку, если callback устарел
    
    kp_id = int(callback.data.split("_")[-1])
    
    # Получаем информацию о КП для показа в подтверждении
    kp_info = kp_db.get_kp_by_id(kp_id)
    customer = kp_info.get('customer_name', 'Неизвестен') if kp_info else 'Неизвестен'
    
    # Показываем подтверждение
    await callback.message.edit_text(
        f"⚠️ Вы уверены, что хотите удалить КП № {kp_id}?\n\n"
        f"Клиент: {customer}\n\n"
        f"Это действие НЕОБРАТИМО!\n"
        f"Будут удалены:\n"
        f"  • Информация о КП\n"
        f"  • Список плит\n"
        f"  • Файлы (PDF и XLSX)\n"
        f"  • Метаданные",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Да, удалить", callback_data=f"delete_kp_confirmed_{kp_id}")],
            [InlineKeyboardButton(text="❌ Отмена", callback_data=f"view_kp_{kp_id}")]
        ])
    )


@router.callback_query(F.data.startswith("delete_kp_confirmed_"))
async def delete_kp_execute(callback: CallbackQuery):
    """
    Выполнить удаление КП из базы данных.
    
    Простыми словами:
    - Удаляет КП из БД (все связанные записи удаляются автоматически)
    - Показывает результат операции
    """
    try:
        await callback.answer()
    except:
        pass  # Игнорируем ошибку, если callback устарел
    
    kp_id = int(callback.data.split("_")[-1])
    
    # Выполняем удаление
    success = kp_db.delete_kp_by_id(kp_id)
    
    if success:
        await callback.message.edit_text(
            f"✅ КП № {kp_id} успешно удалён из базы данных\n\n"
            f"Удалены:\n"
            f"  • Основная информация о КП\n"
            f"  • Все позиции (плиты)\n"
            f"  • Файлы из БД\n"
            f"  • Метаданные и статус",
            reply_markup=archive_sections_kb()
        )
    else:
        await callback.message.edit_text(
            f"❌ Ошибка при удалении КП № {kp_id}\n\n"
            f"Возможно, КП уже был удалён ранее.",
            reply_markup=archive_sections_kb()
        )


@router.callback_query(F.data == "view_current_plan")
async def view_current_plan(callback: CallbackQuery):
    """
    Просмотр актуального плана производства.
    
    Простыми словами:
    - Читает сохранённый файл с информацией об актуальном плане
    - Проверяет существование файла диаграммы Ганта
    - Отправляет файл пользователю с информацией о плане
    """
    try:
        await callback.answer()
    except:
        pass  # Игнорируем ошибку, если callback устарел
    
    # Определяем путь к файлу с данными плана
    # Берём путь к bot/handlers и поднимаемся на уровень выше к bot/
    bot_dir = Path(__file__).parent.parent
    json_path = bot_dir / "data" / "current_plan.json"
    
    # Проверяем существование файла
    if not json_path.exists():
        await callback.message.edit_text(
            "📊 Актуальный план не найден\n\n"
            "❌ Сохранённого актуального плана пока нет.\n\n"
            "💡 Чтобы создать актуальный план:\n"
            "1. Перейдите в раздел '🏭 Производство'\n"
            "2. Создайте план производства\n"
            "3. После просмотра дня нажмите кнопку '💾 Актуальный план'\n\n"
            "После этого диаграмма Ганта будет доступна здесь.",
            reply_markup=archive_sections_kb()
        )
        return
    
    try:
        # Читаем JSON файл
        with open(json_path, 'r', encoding='utf-8') as f:
            plan_data = json.load(f)
        
        gantt_file_path = plan_data.get('gantt_file_path')
        saved_at = plan_data.get('saved_at', 'неизвестно')
        total_days = plan_data.get('total_days', 0)
        tracks_count = plan_data.get('tracks_count', 0)
        
        # Проверяем существование файла диаграммы Ганта
        if not gantt_file_path or not os.path.exists(gantt_file_path):
            await callback.message.edit_text(
                "⚠️ Файл диаграммы Ганта не найден\n\n"
                f"❌ Диаграмма была удалена с диска.\n\n"
                f"📋 Информация о плане:\n"
                f"  • Дата сохранения: {saved_at}\n"
                f"  • Дней производства: {total_days}\n"
                f"  • Дорожек в день: {tracks_count}\n\n"
                f"💡 Создайте новый актуальный план через раздел Производство.",
                reply_markup=archive_sections_kb()
            )
            return
        
        # Отправляем файл
        await callback.message.answer_document(
            FSInputFile(gantt_file_path),
            caption=(
                f"📊 Актуальный план производства\n\n"
                f"📅 Дата сохранения: {saved_at}\n"
                f"📋 Дней производства: {total_days}\n"
                f"🏭 Дорожек в день: {tracks_count}\n\n"
                f"Цветовая кодировка:\n"
                f"🟢 Зелёный — успеваем до дедлайна\n"
                f"🟡 Жёлтый — завершаем в день дедлайна\n"
                f"🔴 Красный — опаздываем!"
            )
        )
        
        # Возвращаем клавиатуру выбора раздела
        await callback.message.answer(
            "Выберите раздел для просмотра:",
            reply_markup=archive_sections_kb()
        )
        
    except Exception as e:
        await callback.message.edit_text(
            f"❌ Ошибка при чтении плана: {e}\n\n"
            f"Попробуйте создать новый актуальный план.",
            reply_markup=archive_sections_kb()
        )


@router.callback_query(F.data == "archive_back_to_sections")
async def back_to_sections(callback: CallbackQuery):
    """Вернуться к выбору раздела архива"""
    try:
        await callback.answer()
    except:
        pass  # Игнорируем ошибку, если callback устарел
    
    await callback.message.edit_text(
        "📁 Архив коммерческих предложений\n\n"
        "Выберите раздел для просмотра:",
        reply_markup=archive_sections_kb()
    )


@router.callback_query(F.data == "archive_back_to_menu")
async def back_to_menu(callback: CallbackQuery):
    """Вернуться в главное меню"""
    try:
        await callback.answer()
    except:
        pass  # Игнорируем ошибку, если callback устарел
    
    await callback.message.answer(
        "Главное меню:",
        reply_markup=main_menu_kb()
    )
