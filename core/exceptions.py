#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Пользовательские исключения для проекта

Эти исключения помогают точно определить, где и что пошло не так,
и показать пользователю понятное сообщение об ошибке.
"""


class PlateParseError(Exception):
    """
    Ошибка парсинга заказа плит
    
    Используется когда текст заказа от пользователя невозможно распознать
    или содержит некорректные данные.
    
    Пример:
        raise PlateParseError("Не удалось распознать размер плиты: '12-abc'")
    """
    pass


class DatabaseConnectionError(Exception):
    """
    Ошибка подключения к базе данных
    
    Используется когда не удаётся подключиться к SQLite базе данных
    или выполнить запрос.
    
    Пример:
        raise DatabaseConnectionError("Не удалось открыть базу pb.db")
    """
    pass


class FileGenerationError(Exception):
    """
    Ошибка генерации файлов (PDF, Excel, PNG)
    
    Используется когда не удаётся создать выходной файл:
    - Коммерческое предложение (PDF/Excel)
    - Схема раскладки (PNG)
    - Ведомость (Excel)
    
    Пример:
        raise FileGenerationError("Не удалось создать PDF файл")
    """
    pass


class PriceNotFoundError(Exception):
    """
    Цена на плиту не найдена в базе данных
    
    Используется когда для заданных параметров плиты (длина, нагрузка)
    нет цены в прайс-листе.
    
    Пример:
        raise PriceNotFoundError("Цена для плиты ПБ 78-12-8п не найдена")
    """
    pass


class UnpricedPlatesError(Exception):
    """Одна или несколько позиций заказа не имеют цены в прайс-листе."""

    def __init__(self, positions: list[str]) -> None:
        self.positions = list(positions)
        labels = ", ".join(self.positions)
        super().__init__(f"Нет цен для позиций: {labels}")
