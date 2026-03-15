"""FSM состояния для диалогов бота"""
from aiogram.fsm.state import StatesGroup, State


class KPStates(StatesGroup):
    """Состояния диалога для получения КП"""
    waiting_for_plate_list = State()
    waiting_for_commercial_offer = State()
    # Состояния для пошагового создания коммерческого предложения
    waiting_manager_selection = State()
    waiting_client_name = State()
    waiting_plates_list = State()
    waiting_plates_confirm = State()
    waiting_wide_plates_replacement = State()
    waiting_discount = State()
    waiting_conditions_choice = State()
    waiting_delivery_conditions = State()
    waiting_payment_conditions = State()
    # Состояния для сохранения КП в БД
    waiting_execution_terms = State()


class CompareStates(StatesGroup):
    """Состояния диалога для сравнения КП и сметы"""
    waiting_kp = State()
    waiting_smeta = State()


class AdminStates(StatesGroup):
    """Состояния для административных команд"""
    waiting_delete_confirmation = State()
    waiting_clear_all_confirmation = State()


class ProductionStates(StatesGroup):
    """Состояния для планирования производства"""
    waiting_start_date = State()      # НОВОЕ: дата начала плана
    waiting_tracks_count = State()    # Количество дорожек
    waiting_filter_method = State()   # Выбор способа фильтрации
    waiting_date_number = State()     # Дата дедлайна
    waiting_kp_numbers = State()      # Ввод номеров КП
    waiting_customer_name = State()   # Ввод заказчика
    waiting_day_selection = State()   # Ожидание выбора дня
    marking_completion = State()      # Отметка брака при завершении дня
    viewing_calendar = State()        # НОВОЕ: просмотр календарного плана
    viewing_plans_list = State()      # Просмотр списка планов
    confirming_plan_delete = State()  # Подтверждение удаления плана
    selecting_kps = State()           # Выбор КП кнопками (мультивыбор + Подтвердить)
    selecting_plates_in_kp = State()  # Выбор плит внутри одного КП


class PBInfoStates(StatesGroup):
    """Состояния для работы с информацией о КП в производстве"""
    waiting_kp_number = State()      # Ожидание ввода номера КП
    waiting_new_date = State()       # Ожидание ввода новой даты


class ArchiveStates(StatesGroup):
    """Состояния для раздела «Архив» (поиск по номеру КП)."""
    waiting_kp_number = State()
    waiting_discount = State()