"""Константы ёмкости производства (дорожки и темп производства).

Используются:
- ``app/services/archive_service.py`` (оценка срока производства КП);
- ``core/delivery_schedule_check.py`` (светофор графика поставки, docs/specs/delivery-schedule.md).

Калибруются по факту на 3–5 прошлых заказах (см. спеку, Success Criteria).
"""

# Максимальная длина одной производственной дорожки, метры.
MAX_TRACK_LENGTH_M = 101.0

# Сколько дорожек производство осваивает за день (дни = дорожки / TRACKS_PER_DAY_DEFAULT).
TRACKS_PER_DAY_DEFAULT = 5.0
