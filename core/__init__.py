"""
Основные модули проекта:
- visualization: визуализация раскладки плит
- optimization: оптимизация раскроя
- config_and_data: конфигурация и данные
- price_db: работа с базой данных цен
- commercial_offer: генерация коммерческих предложений
- kp_db: работа с базой данных КП (plita.db)
"""

from . import visualization
from . import optimization
from . import config_and_data
from . import price_db
from . import commercial_offer
from . import kp_db

__all__ = [
    'visualization',
    'optimization',
    'config_and_data',
    'price_db',
    'commercial_offer',
    'kp_db',
]

