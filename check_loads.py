#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль проверки допустимости нагрузок и опирания плит
Серия ПБ ЖБК СТАРТ
"""
import sqlite3
from dataclasses import dataclass
from typing import Optional


# Классы нагрузок для серии ПБ ЖБК СТАРТ (в сотнях кг/м²)
LOAD_CLASSES = [6, 8, 10, 12.5, 16, 21]

# Минимальное опирание (мм)
MIN_BEARING_MM = {
    'masonry': 100,      # Кирпич/бетон
    'rc_steel': 80       # Железобетон/сталь
}


@dataclass
class LoadCheckResult:
    """Результат проверки нагрузки или опирания"""
    ok: bool                    # Проверка пройдена
    reason: str                 # Причина (пояснение)
    suggest: Optional[str]      # Предложение по исправлению


def check_load(
    con: sqlite3.Connection,
    length_m: float,
    load_class: float
) -> LoadCheckResult:
    """
    Проверяет допустимость класса нагрузки для заданной длины
    
    Args:
        con: Соединение с БД
        length_m: Длина плиты в метрах
        load_class: Класс нагрузки
        
    Returns:
        LoadCheckResult с результатом проверки
    """
    length_mm = int(round(length_m * 1000))
    
    # Проверяем наличие типоразмера в БД
    cursor = con.execute("""
        SELECT COUNT(*) FROM slab_sizes
        WHERE length_mm = ? AND load_class = ?
    """, (length_mm, load_class))
    
    count = cursor.fetchone()[0]
    
    if count > 0:
        return LoadCheckResult(
            ok=True,
            reason=f"Плита {length_m:.2f}м с классом нагрузки {load_class} есть в серии ПБ ЖБК СТАРТ",
            suggest=None
        )
    
    # Ищем доступные классы для этой длины
    cursor = con.execute("""
        SELECT DISTINCT load_class FROM slab_sizes
        WHERE length_mm = ?
        ORDER BY load_class
    """, (length_mm,))
    
    available = [row[0] for row in cursor.fetchall()]
    
    if not available:
        return LoadCheckResult(
            ok=False,
            reason=f"Длина {length_m:.2f}м не найдена в серии ПБ ЖБК СТАРТ",
            suggest="Выберите длину из диапазона 2.98м - 9.88м с шагом 0.4м"
        )
    
    # Класс не подходит, но длина есть
    if load_class < min(available):
        suggest_class = min(available)
        return LoadCheckResult(
            ok=False,
            reason=f"Класс {load_class} слишком низкий для длины {length_m:.2f}м",
            suggest=f"Минимальный класс: {suggest_class}"
        )
    
    if load_class > max(available):
        return LoadCheckResult(
            ok=False,
            reason=f"Класс {load_class} не производится для длины {length_m:.2f}м",
            suggest=f"Доступные классы: {', '.join(map(str, available))}"
        )
    
    # Класс есть, но не точное совпадение
    closest = min(available, key=lambda x: abs(x - load_class))
    return LoadCheckResult(
        ok=False,
        reason=f"Класс {load_class} недоступен для длины {length_m:.2f}м",
        suggest=f"Ближайший класс: {closest}. Доступны: {', '.join(map(str, available))}"
    )


def check_bearing(
    support_type: str,
    bearing_mm: int
) -> LoadCheckResult:
    """
    Проверяет достаточность опирания плиты
    
    Args:
        support_type: Тип опоры ('masonry' - кирпич/бетон, 'rc_steel' - ж/б/сталь)
        bearing_mm: Величина опирания в мм
        
    Returns:
        LoadCheckResult с результатом проверки
    """
    if support_type not in MIN_BEARING_MM:
        return LoadCheckResult(
            ok=False,
            reason="Неизвестный тип опоры",
            suggest="Используйте 'masonry' (кирпич/бетон) или 'rc_steel' (ж/б/сталь)"
        )
    
    min_bearing = MIN_BEARING_MM[support_type]
    
    if bearing_mm < min_bearing:
        return LoadCheckResult(
            ok=False,
            reason=f"Опирание {bearing_mm} мм меньше допустимого {min_bearing} мм для {support_type}",
            suggest=f"Увеличьте опирание до {min_bearing} мм или используйте закладные/усиление"
        )
    
    support_name = "кирпич/бетон" if support_type == 'masonry' else "железобетон/сталь"
    return LoadCheckResult(
        ok=True,
        reason=f"Опирание {bearing_mm} мм допустимо для {support_name} (мин. {min_bearing} мм)",
        suggest=None
    )


def check_length_range(
    con: sqlite3.Connection,
    length_m: float
) -> LoadCheckResult:
    """
    Проверяет, входит ли длина в диапазон серии
    
    Args:
        con: Соединение с БД
        length_m: Длина в метрах
        
    Returns:
        LoadCheckResult с результатом проверки
    """
    cursor = con.execute("""
        SELECT MIN(length_mm), MAX(length_mm) FROM slab_sizes
    """)
    
    min_mm, max_mm = cursor.fetchone()
    min_m = min_mm / 1000.0
    max_m = max_mm / 1000.0
    
    if length_m < min_m:
        return LoadCheckResult(
            ok=False,
            reason=f"Длина {length_m:.2f}м меньше минимальной {min_m:.2f}м",
            suggest=f"Используйте доборы или минимальную длину {min_m:.2f}м"
        )
    
    if length_m > max_m:
        return LoadCheckResult(
            ok=False,
            reason=f"Длина {length_m:.2f}м больше максимальной {max_m:.2f}м",
            suggest=f"Разбейте на несколько плит или используйте макс. {max_m:.2f}м"
        )
    
    return LoadCheckResult(
        ok=True,
        reason=f"Длина {length_m:.2f}м в диапазоне серии ({min_m:.2f}м - {max_m:.2f}м)",
        suggest=None
    )


def format_check_message(result: LoadCheckResult) -> str:
    """
    Форматирует результат проверки для вывода пользователю
    
    Args:
        result: Результат проверки
        
    Returns:
        Строка с сообщением
    """
    icon = "✅" if result.ok else "⚠️"
    message = f"{icon} {result.reason}"
    
    if result.suggest:
        message += f"\n💡 {result.suggest}"
    
    return message


if __name__ == "__main__":
    # Тест модуля
    import sys
    sys.path.insert(0, '.')
    
    conn = sqlite3.connect('pb.db')
    
    print('=== ПРОВЕРКА НАГРУЗОК ===\n')
    
    test_cases = [
        (5.58, 8),    # OK
        (6.68, 12.5), # OK
        (9.88, 21),   # OK
        (5.58, 25),   # Класс слишком высокий
        (3.0, 6),     # Длина не производится
    ]
    
    for length_m, load_class in test_cases:
        result = check_load(conn, length_m, load_class)
        print(f"Длина {length_m}м, класс {load_class}:")
        print(f"  {format_check_message(result)}\n")
    
    print('=== ПРОВЕРКА ОПИРАНИЯ ===\n')
    
    bearing_tests = [
        ('masonry', 100),   # OK
        ('masonry', 80),    # Недостаточно
        ('rc_steel', 80),   # OK
        ('rc_steel', 70),   # Недостаточно
    ]
    
    for support_type, bearing_mm in bearing_tests:
        result = check_bearing(support_type, bearing_mm)
        print(f"{support_type}, {bearing_mm}мм:")
        print(f"  {format_check_message(result)}\n")
    
    conn.close()

















