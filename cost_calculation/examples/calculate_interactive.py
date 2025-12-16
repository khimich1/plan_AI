#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Программа для расчета себестоимости плиты по названию
"""

from cost_calculation import calculate_plate_cost
import core.config_and_data as cfg


def main():
    """Основная функция"""
    print("=" * 80)
    print("РАСЧЕТ СЕБЕСТОИМОСТИ ПЛИТЫ")
    print("=" * 80)
    print()
    print("Формат названия: ПБ {длина_дм}-{ширина_дм}-{нагрузка}п")
    print("Примеры: ПБ 17-12-6, ПБ 20-12-8, ПБ 30-12-10")
    print()
    
    while True:
        plate_name = input("Введите название плиты (или 'exit' для выхода): ").strip()
        
        if plate_name.lower() in ['exit', 'quit', 'выход']:
            print("До свидания!")
            break
        
        if not plate_name:
            continue
        
        print()
        print("-" * 80)
        
        result = calculate_plate_cost(plate_name, cfg.PRICE_DB_PATH)
        
        if result:
            print(f"📋 Плита: {result['plate_name']}")
            print()
            print(f"   Параметры:")
            params = result['parameters']
            print(f"     - Длина: {params['length_m']:.1f} м ({params['length_dm']} дм)")
            print(f"     - Ширина: {params['width_m']:.1f} м ({params['width_dm']} дм)")
            print(f"     - Нагрузка: {params['load_code']}п")
            print(f"     - Марка бетона: {params['concrete_grade']}")
            print(f"     - Объем: {result['volume_m3']:.4f} м³")
            
            print()
            print(f"   Компоненты себестоимости:")
            components = result['components']
            for component, cost in components.items():
                component_name = {
                    'concrete': 'Бетон',
                    'reinforcement': 'Армирование',
                    'loops': 'Петли',
                    'izoform': 'Изоформ'
                }.get(component, component.capitalize())
                print(f"     - {component_name}: {cost:,.2f} руб")
            
            print()
            print(f"   💰 ИТОГО: {result['total_cost']:,.2f} руб")
            
            # Проверяем, есть ли данные из Excel для сравнения
            import sqlite3
            conn = sqlite3.connect(cfg.PRICE_DB_PATH)
            try:
                cur = conn.cursor()
                cur.execute("""
                    SELECT total_cost FROM excel_total_costs
                    WHERE length_dm=? AND width_dm=? AND load_code=?
                """, (params['length_dm'], params['width_dm'], params['load_code']))
                row = cur.fetchone()
                if row:
                    excel_cost = row[0]
                    diff = abs(result['total_cost'] - excel_cost)
                    diff_pct = (diff / excel_cost * 100) if excel_cost > 0 else 0
                    print()
                    print(f"   📊 Сравнение с Excel:")
                    print(f"     - Себестоимость из Excel: {excel_cost:,.2f} руб")
                    print(f"     - Разница: {diff:,.2f} руб ({diff_pct:.2f}%)")
                    if diff_pct < 1:
                        print(f"     ✅ Совпадает с Excel!")
                    elif diff_pct < 5:
                        print(f"     ⚠️ Небольшое расхождение")
                    else:
                        print(f"     ❌ Значительное расхождение")
            finally:
                conn.close()
            
            print()
            print("-" * 80)
            print()
        else:
            print(f"❌ Ошибка: не удалось рассчитать себестоимость для '{plate_name}'")
            print("Проверьте формат названия: ПБ {длина}-{ширина}-{нагрузка}п")
            print()
            print("-" * 80)
            print()


if __name__ == "__main__":
    main()

