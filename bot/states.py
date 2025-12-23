"""FSM состояния для диалогов бота"""
from aiogram.fsm.state import StatesGroup, State


class KPStates(StatesGroup):
    """Состояния диалога для получения КП"""
    waiting_for_plate_list = State()
    waiting_for_commercial_offer = State()
    # Состояния для пошагового создания коммерческого предложения
    waiting_manager_name = State()
    waiting_client_name = State()
    waiting_plates_list = State()
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


class ProductionStates(StatesGroup):
    """Состояния для планирования производства"""
    waiting_tracks_count = State()    # Количество дорожек
    waiting_date_number = State()     # Дата дедлайна