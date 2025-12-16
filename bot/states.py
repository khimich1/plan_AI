"""FSM состояния для диалогов бота"""
from aiogram.fsm.state import StatesGroup, State


class KPStates(StatesGroup):
    """Состояния диалога для получения КП"""
    waiting_for_plate_list = State()
    waiting_for_commercial_offer = State()


class CompareStates(StatesGroup):
    """Состояния диалога для сравнения КП и сметы"""
    waiting_kp = State()
    waiting_smeta = State()

