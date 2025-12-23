#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Детальное сравнение наших расчетов со значениями из Excel
Выясняет причины расхождений
"""

from cost_calculation import calculate_plate_cost
import core.config_and_data as cfg
import sqlite3

def compare_plate_with_excel(plate_name: str):
    """Сравнивает расчет плиты с данными из Excel"""
    
    print("=" * 100)
    print(f"АНАЛИЗ РАСХОЖДЕНИЙ ДЛЯ ПЛИТЫ: {plate_name}")
    print("=" * 100)
    print()
    
    # Получаем наш расчет
    result = calculate_plate_cost(plate_name, cfg.PRICE_DB_PATH)
    
    if not result:
        print(f"❌ Не удалось рассчитать себестоимость для '{plate_name}'")
        return
    
    params = result['parameters']
    length_dm = params['length_dm']
    width_dm = params['width_dm']
    load_code = params['load_code']
    
    # Получаем данные из Excel
    conn = sqlite3.connect(cfg.PRICE_DB_PATH)
    try:
        cur = conn.cursor()
        
        # Объем
        cur.execute("""
            SELECT volume_m3 FROM plate_volumes
            WHERE length_dm=? AND width_dm=? AND load_code=?
        """, (length_dm, width_dm, load_code))
        excel_volume = cur.fetchone()
        
        # Бетон
        cur.execute("""
            SELECT concrete_cost FROM concrete_costs
            WHERE length_dm=? AND width_dm=? AND load_code=?
        """, (length_dm, width_dm, load_code))
        excel_concrete = cur.fetchone()
        
        # Армирование
        cur.execute("""
            SELECT wire_kg, cable_cost FROM reinforcement_costs
            WHERE length_dm=? AND width_dm=? AND load_code=?
        """, (length_dm, width_dm, load_code))
        excel_reinforcement = cur.fetchone()
        
        # Изоформ
        cur.execute("""
            SELECT izoform_kg, izoform_cost FROM izoform_costs
            WHERE length_dm=? AND width_dm=? AND load_code=?
        """, (length_dm, width_dm, load_code))
        excel_izoform = cur.fetchone()
        
        # Общая себестоимость из Excel (колонка 19)
        cur.execute("""
            SELECT total_cost FROM excel_total_costs
            WHERE length_dm=? AND width_dm=? AND load_code=?
        """, (length_dm, width_dm, load_code))
        excel_total = cur.fetchone()
        
        # КЭФ
        cur.execute("""
            SELECT kef FROM plate_kef_values
            WHERE length_dm=? AND width_dm=? AND load_code=?
        """, (length_dm, width_dm, load_code))
        excel_kef = cur.fetchone()
        
        # Получаем цену проволоки для расчета
        cur.execute("SELECT value FROM cost_constants WHERE key='wire_price_per_kg'")
        wire_price_row = cur.fetchone()
        wire_price = wire_price_row[0] if wire_price_row else 80.0
        
    finally:
        conn.close()
    
    # Выводим сравнение
    print("📊 ПАРАМЕТРЫ ПЛИТЫ:")
    print(f"   Длина: {length_dm} дм = {params['length_m']} м")
    print(f"   Ширина: {width_dm} дм = {params['width_m']} м")
    print(f"   Нагрузка: {load_code}п")
    print(f"   Марка бетона: {params['concrete_grade']}")
    print()
    
    # 1. ОБЪЕМ
    print("=" * 100)
    print("1. ОБЪЕМ ПЛИТЫ")
    print("=" * 100)
    our_volume = result['volume_m3']
    excel_volume_val = excel_volume[0] if excel_volume else None
    
    print(f"   Наш расчет:     {our_volume:.4f} м³")
    if excel_volume_val:
        print(f"   Из Excel:       {excel_volume_val:.4f} м³")
        diff = abs(our_volume - excel_volume_val)
        if diff < 0.0001:
            print(f"   ✅ СОВПАДАЕТ (разница: {diff:.6f} м³)")
        else:
            print(f"   ⚠️  РАСХОЖДЕНИЕ: {diff:.4f} м³ ({diff/our_volume*100:.2f}%)")
    else:
        print(f"   ⚠️  Нет данных в Excel")
    print()
    
    # 2. БЕТОН
    print("=" * 100)
    print("2. СТОИМОСТЬ БЕТОНА")
    print("=" * 100)
    our_concrete = result['components']['concrete']
    excel_concrete_val = excel_concrete[0] if excel_concrete else None
    
    print(f"   Наш расчет:     {our_concrete:,.2f} руб")
    if excel_concrete_val:
        print(f"   Из Excel:       {excel_concrete_val:,.2f} руб (колонка 13)")
        diff = abs(our_concrete - excel_concrete_val)
        diff_pct = (diff / excel_concrete_val * 100) if excel_concrete_val > 0 else 0
        if diff < 0.01:
            print(f"   ✅ СОВПАДАЕТ (разница: {diff:.2f} руб)")
        else:
            print(f"   ⚠️  РАСХОЖДЕНИЕ: {diff:,.2f} руб ({diff_pct:.2f}%)")
    else:
        print(f"   ⚠️  Нет данных в Excel")
    
    # Детализация бетона
    breakdown = result['breakdown']
    if 'cement_kg' in breakdown['concrete']:
        print()
        print("   Детализация (наш расчет):")
        print(f"     Цемент:     {breakdown['concrete']['cement_kg']:.2f} кг")
        print(f"     Песок:      {breakdown['concrete']['sand_kg']:.2f} кг")
        print(f"     Щебень:     {breakdown['concrete']['gravel_kg']:.2f} кг")
        if 'polyplast_kg' in breakdown['concrete']:
            print(f"     Полипласт:  {breakdown['concrete']['polyplast_kg']:.2f} кг")
    print()
    
    # 3. АРМИРОВАНИЕ
    print("=" * 100)
    print("3. СТОИМОСТЬ АРМИРОВАНИЯ")
    print("=" * 100)
    our_reinforcement = result['components']['reinforcement']
    
    print(f"   Наш расчет:     {our_reinforcement:,.2f} руб")
    
    if excel_reinforcement:
        wire_kg_excel = excel_reinforcement[0] if excel_reinforcement[0] else 0
        cable_cost_excel = excel_reinforcement[1] if excel_reinforcement[1] else 0
        
        # Рассчитываем стоимость проволоки из Excel
        wire_cost_excel = wire_kg_excel * wire_price
        total_reinforcement_excel = wire_cost_excel + cable_cost_excel
        
        print(f"   Из Excel:")
        print(f"     Проволока:  {wire_kg_excel:.3f} кг × {wire_price} руб/кг = {wire_cost_excel:,.2f} руб")
        print(f"     Канат:      {cable_cost_excel:,.2f} руб (колонка 6)")
        print(f"     ИТОГО:      {total_reinforcement_excel:,.2f} руб")
        
        diff = abs(our_reinforcement - total_reinforcement_excel)
        diff_pct = (diff / total_reinforcement_excel * 100) if total_reinforcement_excel > 0 else 0
        
        if diff < 0.01:
            print(f"   ✅ СОВПАДАЕТ (разница: {diff:.2f} руб)")
        else:
            print(f"   ⚠️  РАСХОЖДЕНИЕ: {diff:,.2f} руб ({diff_pct:.2f}%)")
            
            # Анализ расхождения
            print()
            print("   Анализ расхождения:")
            if 'wire_cost' in breakdown['reinforcement']:
                our_wire_cost = breakdown['reinforcement']['wire_cost']
                our_wire_kg = breakdown['reinforcement'].get('wire_kg', 0)
                print(f"     Наша проволока: {our_wire_kg:.3f} кг × {wire_price} руб/кг = {our_wire_cost:,.2f} руб")
                print(f"     Excel проволока: {wire_kg_excel:.3f} кг × {wire_price} руб/кг = {wire_cost_excel:,.2f} руб")
                
                if abs(our_wire_kg - wire_kg_excel) > 0.001:
                    print(f"     ⚠️  Разница в количестве проволоки: {abs(our_wire_kg - wire_kg_excel):.3f} кг")
                
                if abs(our_wire_cost - wire_cost_excel) > 0.01:
                    print(f"     ⚠️  Разница в стоимости проволоки: {abs(our_wire_cost - wire_cost_excel):,.2f} руб")
            
            if 'cable_cost' in breakdown['reinforcement']:
                our_cable_cost = breakdown['reinforcement']['cable_cost']
                print(f"     Наш канат: {our_cable_cost:,.2f} руб")
                print(f"     Excel канат: {cable_cost_excel:,.2f} руб")
                
                if abs(our_cable_cost - cable_cost_excel) > 0.01:
                    print(f"     ⚠️  Разница в стоимости каната: {abs(our_cable_cost - cable_cost_excel):,.2f} руб")
    else:
        print(f"   ⚠️  Нет данных в Excel")
    
    # Наш расчет детально
    if 'wire_cost' in breakdown['reinforcement']:
        print()
        print("   Наш расчет детально:")
        print(f"     Проволока:  {breakdown['reinforcement'].get('wire_kg', 0):.3f} кг × {wire_price} руб/кг = {breakdown['reinforcement']['wire_cost']:,.2f} руб")
        if 'cable_cost' in breakdown['reinforcement']:
            print(f"     Канат:      {breakdown['reinforcement']['cable_cost']:,.2f} руб")
    print()
    
    # 4. ПЕТЛИ
    print("=" * 100)
    print("4. СТОИМОСТЬ ПЕТЕЛЬ")
    print("=" * 100)
    our_loops = result['components']['loops']
    expected_loops = 572.0  # Стандартно 2 петли × 286 руб
    
    print(f"   Наш расчет:     {our_loops:,.2f} руб")
    print(f"   Ожидается:      {expected_loops:,.2f} руб (2 петли × 286 руб)")
    diff = abs(our_loops - expected_loops)
    if diff < 0.01:
        print(f"   ✅ СОВПАДАЕТ")
    else:
        print(f"   ⚠️  РАСХОЖДЕНИЕ: {diff:,.2f} руб")
    print()
    
    # 5. ИЗОФОРМ
    print("=" * 100)
    print("5. СТОИМОСТЬ ИЗОФОРМА")
    print("=" * 100)
    our_izoform = result['components']['izoform']
    excel_izoform_val = excel_izoform[1] if excel_izoform and excel_izoform[1] else None
    excel_izoform_kg = excel_izoform[0] if excel_izoform and excel_izoform[0] else None
    
    print(f"   Наш расчет:     {our_izoform:,.2f} руб")
    if excel_izoform_val:
        print(f"   Из Excel:       {excel_izoform_val:,.2f} руб (колонка 18)")
        diff = abs(our_izoform - excel_izoform_val)
        if diff < 0.01:
            print(f"   ✅ СОВПАДАЕТ (разница: {diff:.2f} руб)")
        else:
            print(f"   ⚠️  РАСХОЖДЕНИЕ: {diff:,.2f} руб")
    else:
        print(f"   ⚠️  Нет данных в Excel")
    
    if 'izoform_kg' in breakdown:
        print(f"   Количество:    {breakdown['izoform_kg']:.4f} кг")
        if excel_izoform_kg:
            print(f"   Excel количество: {excel_izoform_kg:.4f} кг")
    print()
    
    # 6. ПРЯМЫЕ ЗАТРАТЫ
    print("=" * 100)
    print("6. ПРЯМЫЕ ЗАТРАТЫ (лист 'Стоимость', колонка R)")
    print("=" * 100)
    our_direct = result['direct_cost']
    excel_total_val = excel_total[0] if excel_total else None
    
    print(f"   Наш расчет:     {our_direct:,.2f} руб")
    print(f"   Компоненты:")
    print(f"     Бетон:        {result['components']['concrete']:,.2f} руб")
    print(f"     Армирование:  {result['components']['reinforcement']:,.2f} руб")
    print(f"     Петли:        {result['components']['loops']:,.2f} руб")
    print(f"     Изоформ:      {result['components']['izoform']:,.2f} руб")
    print(f"     Сумма:        {sum(result['components'].values()):,.2f} руб")
    
    if excel_total_val:
        print(f"   Из Excel (колонка 19): {excel_total_val:,.2f} руб")
        diff = abs(our_direct - excel_total_val)
        diff_pct = (diff / excel_total_val * 100) if excel_total_val > 0 else 0
        
        if diff < 1:
            print(f"   ✅ СОВПАДАЕТ (разница: {diff:,.2f} руб)")
        else:
            print(f"   ⚠️  РАСХОЖДЕНИЕ: {diff:,.2f} руб ({diff_pct:.2f}%)")
            
            # Анализ расхождения
            print()
            print("   Анализ расхождения:")
            print(f"     Разница: {diff:,.2f} руб")
            print()
            print("     Возможные причины:")
            print(f"     1. Колонка 19 в Excel может содержать прямые затраты БЕЗ некоторых компонентов")
            print(f"     2. В Excel может быть другая формула для колонки 19")
            print(f"     3. Колонка 19 может быть из другого листа (не 'Стоимость')")
            print()
            print("     Проверка: что если колонка 19 = наши прямые БЕЗ чего-то?")
            
            # Проверяем разные варианты
            without_loops = our_direct - result['components']['loops']
            without_izoform = our_direct - result['components']['izoform']
            without_both = our_direct - result['components']['loops'] - result['components']['izoform']
            
            print(f"       Без петель: {without_loops:,.2f} руб (разница: {abs(without_loops - excel_total_val):,.2f})")
            print(f"       Без изоформа: {without_izoform:,.2f} руб (разница: {abs(without_izoform - excel_total_val):,.2f})")
            print(f"       Без петель и изоформа: {without_both:,.2f} руб (разница: {abs(without_both - excel_total_val):,.2f})")
            
            # Проверяем, какие компоненты могут давать расхождение
            components_diff = {
                'Бетон': abs(result['components']['concrete'] - (excel_concrete_val if excel_concrete_val else 0)),
                'Армирование': abs(result['components']['reinforcement'] - (total_reinforcement_excel if excel_reinforcement else 0)),
                'Петли': abs(result['components']['loops'] - expected_loops),
                'Изоформ': abs(result['components']['izoform'] - (excel_izoform_val if excel_izoform_val else 0))
            }
            
            print()
            print("     Проверка компонентов (все должны быть 0):")
            for comp, comp_diff in sorted(components_diff.items(), key=lambda x: x[1], reverse=True):
                if comp_diff > 0.01:
                    print(f"       ⚠️  {comp}: разница {comp_diff:,.2f} руб")
                else:
                    print(f"       ✅ {comp}: совпадает")
    else:
        print(f"   ⚠️  Нет данных в Excel (колонка 19)")
    print()
    
    # 7. КЭФ И ПОЛНАЯ СЕБЕСТОИМОСТЬ
    print("=" * 100)
    print("7. КЭФ И ПОЛНАЯ СЕБЕСТОИМОСТЬ (лист 'Себестоимость')")
    print("=" * 100)
    our_kef = result['kef']
    our_overhead = result['overhead_cost']
    our_full = result['full_cost_with_kef']
    
    excel_kef_val = excel_kef[0] if excel_kef else None
    
    print(f"   КЭФ:")
    print(f"     Наш расчет:     {our_kef:.2f}")
    if excel_kef_val:
        print(f"     Из Excel:       {excel_kef_val:.2f}")
        if abs(our_kef - excel_kef_val) < 0.01:
            print(f"     ✅ СОВПАДАЕТ")
        else:
            print(f"     ⚠️  РАСХОЖДЕНИЕ: {abs(our_kef - excel_kef_val):.2f}")
    else:
        print(f"     ⚠️  Используется дефолтное значение из констант")
    print()
    
    print(f"   Накладные расходы:")
    print(f"     Наш расчет:     {our_overhead:,.2f} руб")
    print(f"     Формула:        Прямые × КЭФ = {our_direct:,.2f} × {our_kef:.2f} = {our_direct * our_kef:,.2f} руб")
    if abs(our_overhead - (our_direct * our_kef)) < 0.01:
        print(f"     ✅ Формула правильная (M3 = L3 × КЭФ)")
    else:
        print(f"     ❌ ОШИБКА В ФОРМУЛЕ!")
    print()
    
    print(f"   Полная себестоимость:")
    print(f"     Наш расчет:     {our_full:,.2f} руб")
    print(f"     Формула:        Прямые + Накладные = {our_direct:,.2f} + {our_overhead:,.2f} = {our_direct + our_overhead:,.2f} руб")
    if abs(our_full - (our_direct + our_overhead)) < 0.01:
        print(f"     ✅ Формула правильная (E3 = Прямые + M3)")
    else:
        print(f"     ❌ ОШИБКА В ФОРМУЛЕ!")
    print()
    
    # ИТОГОВЫЙ ВЫВОД
    print("=" * 100)
    print("ИТОГОВЫЙ ВЫВОД")
    print("=" * 100)
    print()
    
    issues = []
    
    if excel_volume_val and abs(our_volume - excel_volume_val) > 0.0001:
        issues.append("Объем не совпадает с Excel")
    
    if excel_concrete_val and abs(our_concrete - excel_concrete_val) > 0.01:
        issues.append("Стоимость бетона не совпадает с Excel")
    
    if excel_reinforcement:
        total_reinforcement_excel = (excel_reinforcement[0] * wire_price if excel_reinforcement[0] else 0) + (excel_reinforcement[1] if excel_reinforcement[1] else 0)
        if abs(our_reinforcement - total_reinforcement_excel) > 0.01:
            issues.append("Стоимость армирования не совпадает с Excel")
    
    if abs(our_loops - expected_loops) > 0.01:
        issues.append("Стоимость петель не совпадает")
    
    if excel_izoform_val and abs(our_izoform - excel_izoform_val) > 0.01:
        issues.append("Стоимость изоформа не совпадает с Excel")
    
    if excel_total_val and abs(our_direct - excel_total_val) > 1:
        issues.append(f"Прямые затраты не совпадают с Excel (разница: {abs(our_direct - excel_total_val):,.2f} руб)")
    
    if abs(our_overhead - (our_direct * our_kef)) > 0.01:
        issues.append("❌ КРИТИЧНО: Формула накладных расходов неправильная!")
    
    if abs(our_full - (our_direct + our_overhead)) > 0.01:
        issues.append("❌ КРИТИЧНО: Формула полной себестоимости неправильная!")
    
    if issues:
        print("⚠️  НАЙДЕНЫ ПРОБЛЕМЫ:")
        for i, issue in enumerate(issues, 1):
            print(f"   {i}. {issue}")
    else:
        print("✅ Все проверки пройдены успешно!")
    
    print()
    print("=" * 100)
    print()


def main():
    """Основная функция"""
    import sys
    
    # Если указана плита в аргументах
    if len(sys.argv) > 1:
        plate_name = sys.argv[1].strip()
        compare_plate_with_excel(plate_name)
    else:
        # Тестируем несколько плит
        test_plates = [
            "ПБ 17-12-6",
            "ПБ 56-12-8п",
            "ПБ 30-12-10"
        ]
        
        print("=" * 100)
        print("ДЕТАЛЬНОЕ СРАВНЕНИЕ С EXCEL")
        print("=" * 100)
        print()
        print("Тестируем плиты:", ", ".join(test_plates))
        print()
        
        for plate_name in test_plates:
            compare_plate_with_excel(plate_name)
            print()


if __name__ == "__main__":
    main()

