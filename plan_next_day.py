"""
Программа планирования дорожек НА СЛЕДУЮЩИЙ ДЕНЬ
Исключает плиты, использованные вчера, и планирует из оставшихся
"""
import pandas as pd
import sqlite3
import sys
from datetime import datetime, timedelta
from typing import List, Dict, Tuple

# Настройка кодировки для Windows
if sys.platform == 'win32':
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'ignore')


def parse_plate_marking(marking: str):
    """Парсит маркировку плиты"""
    import re
    pattern = r'ПБ\s+(\d+[,.]?\d*)-(\d+[,.]?\d*)-(\d+)'
    match = re.search(pattern, marking)
    
    if not match:
        return None
    
    try:
        length_dm = float(match.group(1).replace(',', '.'))
        length_m = length_dm / 10
        
        width_dm = float(match.group(2).replace(',', '.'))
        width_m = width_dm / 10
        
        load_code = int(match.group(3))
        load_capacity = load_code * 100
        
        return {
            'length': length_m,
            'width': width_m,
            'load_capacity': load_capacity
        }
    except:
        return None


def load_used_plates_from_excel(excel_file: str) -> List[str]:
    """Загружает список использованных плит из Excel файла предыдущего дня"""
    used_markings = set()
    
    try:
        # Читаем все листы Excel файла
        excel_data = pd.read_excel(excel_file, sheet_name=None)
        
        for sheet_name, df in excel_data.items():
            if 'Дорожка' in sheet_name or 'Раскладка' in sheet_name:
                # Ищем колонку с маркировкой плит
                marking_cols = [col for col in df.columns if 'маркировка' in col.lower() or 'номенклатура' in col.lower()]
                
                for col in marking_cols:
                    if col in df.columns:
                        markings = df[col].dropna().astype(str)
                        used_markings.update(markings)
        
        print(f"[ЗАГРУЗКА] Из файла {excel_file} найдено {len(used_markings)} использованных плит")
        
    except Exception as e:
        print(f"[ОШИБКА] Не удалось загрузить использованные плиты: {e}")
        print(f"[ИНФОРМАЦИЯ] Будем использовать все плиты из базы")
    
    return list(used_markings)


def load_available_plates(db_path: str = 'pb.db', used_markings: List[str] = None, min_load: float = 800) -> pd.DataFrame:
    """Загружает доступные плиты, исключая использованные"""
    conn = sqlite3.connect(db_path)
    
    query = '''
        SELECT 
            "номенклатура пб к производству" as marking,
            "длина плиты, м" as length_db,
            "контрагент заказчик" as customer,
            "армирование по серии" as reinforcement,
            "неделя формовки" as week
        FROM plity_ex
    '''
    df = pd.read_sql(query, conn)
    conn.close()
    
    if used_markings:
        # Исключаем использованные плиты
        df = df[~df['marking'].isin(used_markings)]
        print(f"[ФИЛЬТРАЦИЯ] Исключено {len(used_markings)} использованных плит")
    
    # Парсим маркировку и фильтруем по нагрузке
    parsed = []
    for idx, row in df.iterrows():
        params = parse_plate_marking(row['marking'])
        if params and params['load_capacity'] >= min_load:
            parsed.append({
                'marking': row['marking'],
                'length': params['length'],
                'width': params['width'],
                'load_capacity': params['load_capacity'],
                'customer': row['customer'],
                'reinforcement': row['reinforcement'],
                'week': row['week']
            })
    
    result_df = pd.DataFrame(parsed)
    
    # Сортируем по срочности (неделя формовки)
    if len(result_df) > 0:
        result_df = result_df.sort_values('week', ascending=True).reset_index(drop=True)
    
    return result_df


def find_best_width_combination(plates: pd.DataFrame, target_width: float, gap_cm: float = 1.0):
    """Подбирает комбинацию плит по ширине"""
    gap_m = gap_cm / 100
    target_width_with_gaps = target_width
    
    widths = plates['width'].unique()
    widths = sorted(widths, reverse=True)
    
    best_combination = None
    min_gap = float('inf')
    
    for main_width in widths:
        plates_with_width = plates[plates['width'] == main_width]
        if len(plates_with_width) == 0:
            continue
        
        num_plates = 1
        current_width = main_width
        
        while current_width < target_width_with_gaps and num_plates < 10:
            num_plates += 1
            current_width = num_plates * main_width + (num_plates - 1) * gap_m
        
        gap = abs(current_width - target_width_with_gaps)
        
        if gap < min_gap and len(plates_with_width) >= num_plates:
            min_gap = gap
            best_combination = {
                'width': main_width,
                'count': num_plates,
                'total_width': current_width,
                'gap': gap
            }
    
    return best_combination


def subset_sum_for_length_improved(plates: List[Dict], target_length_cm: int) -> Tuple[List[int], int]:
    """Улучшенный алгоритм подбора плит по длине"""
    if not plates:
        return [], 0
    
    # Сортируем плиты по длине (от больших к меньшим)
    sorted_plates = sorted(plates, key=lambda x: x['length'], reverse=True)
    
    used_plates = []
    current_length_cm = 0
    
    # Сначала пытаемся точно попасть в целевую длину
    for plate in sorted_plates:
        length_cm = int(plate['length'] * 100)
        if current_length_cm + length_cm <= target_length_cm:
            used_plates.append(plate)
            current_length_cm += length_cm
    
    # Если не заполнили полностью, пробуем добавить плиты меньшей длины
    if current_length_cm < target_length_cm:
        remaining_cm = target_length_cm - current_length_cm
        for plate in sorted_plates:
            length_cm = int(plate['length'] * 100)
            if length_cm <= remaining_cm:
                used_plates.append(plate)
                current_length_cm += length_cm
                remaining_cm = target_length_cm - current_length_cm
                if remaining_cm <= 0:
                    break
    
    # Возвращаем индексы в оригинальном списке
    indices = []
    for used_plate in used_plates:
        for i, original_plate in enumerate(plates):
            if (original_plate['marking'] == used_plate['marking'] and 
                original_plate['length'] == used_plate['length'] and
                original_plate['customer'] == used_plate['customer']):
                indices.append(i)
                break
    
    return indices, current_length_cm


def plan_track_for_next_day(length_m: float, width_m: float, plates_df: pd.DataFrame, 
                           gap_cm: float = 1.0, track_num: int = 1):
    """Планирует одну дорожку для следующего дня"""
    
    print(f"\n{'='*70}")
    print(f"🛤️  ДОРОЖКА #{track_num} (СЛЕДУЮЩИЙ ДЕНЬ)")
    print(f"{'='*70}")
    
    if len(plates_df) == 0:
        print("❌ Нет доступных плит для планирования!")
        return None
    
    # Подбираем по ширине
    print(f"🔍 Подбираю плиты по ширине ({width_m} м)...")
    width_combo = find_best_width_combination(plates_df, width_m, gap_cm)
    
    if not width_combo:
        print("❌ Не удалось подобрать плиты по ширине!")
        return None
    
    print(f"✅ Найдена комбинация:")
    print(f"   • Ширина плиты: {width_combo['width']} м")
    print(f"   • Количество плит поперёк: {width_combo['count']} шт")
    print(f"   • Общая ширина: {width_combo['total_width']:.3f} м")
    print(f"   • Отклонение: {width_combo['gap']*100:.1f} см")
    
    # Фильтруем плиты нужной ширины
    selected_width = width_combo['width']
    plates_filtered = plates_df[plates_df['width'] == selected_width].copy()
    
    if len(plates_filtered) == 0:
        print(f"❌ Нет доступных плит шириной {selected_width} м")
        return None
    
    print(f"\n🔍 Подбираю плиты по длине ({length_m} м)...")
    print(f"   Доступно плит шириной {selected_width} м: {len(plates_filtered)}")
    
    # Группируем по армированию
    by_reinforcement = {}
    for reinf in plates_filtered['reinforcement'].unique():
        if pd.isna(reinf):
            continue
        group = plates_filtered[plates_filtered['reinforcement'] == reinf]
        by_reinforcement[str(reinf)] = group.to_dict('records')
    
    # Подбираем по длине
    best_layout = None
    best_coverage = 0
    
    for reinf, plates_list in by_reinforcement.items():
        if len(plates_list) < width_combo['count']:
            continue
        
        target_length_cm = int(length_m * 100)
        indices, achieved_cm = subset_sum_for_length_improved(plates_list, target_length_cm)
        
        coverage_ratio = achieved_cm / target_length_cm if target_length_cm > 0 else 0
        
        if (achieved_cm > best_coverage or 
            (achieved_cm == best_coverage and coverage_ratio > (best_layout.get('coverage_ratio', 0) if best_layout else 0))):
            best_coverage = achieved_cm
            best_layout = {
                'reinforcement': reinf,
                'plates': [plates_list[i] for i in indices],
                'achieved_cm': achieved_cm,
                'achieved_m': achieved_cm / 100,
                'gap_m': length_m - achieved_cm / 100,
                'coverage_ratio': coverage_ratio,
                'total_plates_available': len(plates_list)
            }
    
    if not best_layout:
        print("❌ Не удалось подобрать плиты по длине!")
        return None
    
    print(f"✅ Найдена раскладка:")
    print(f"   • Армирование: {best_layout['reinforcement']}")
    print(f"   • Плит вдоль дорожки: {len(best_layout['plates'])} шт")
    print(f"   • Покрытая длина: {best_layout['achieved_m']:.2f} м")
    print(f"   • Остаток: {best_layout['gap_m']*100:.0f} см")
    if 'coverage_ratio' in best_layout:
        print(f"   • Эффективность: {best_layout['coverage_ratio']*100:.1f}%")
    if 'total_plates_available' in best_layout:
        print(f"   • Доступно плит этого типа: {best_layout['total_plates_available']} шт")
    
    # Показываем недели формовки
    used_weeks = [p['week'] for p in best_layout['plates']]
    unique_weeks = sorted(set(used_weeks))
    print(f"   • Недели формовки: {', '.join(map(str, unique_weeks[:5]))}")
    
    total_plates = len(best_layout['plates']) * width_combo['count']
    
    result = {
        'track_num': track_num,
        'length_m': length_m,
        'width_m': width_m,
        'width_combo': width_combo,
        'length_layout': best_layout,
        'total_plates': total_plates,
    }
    
    print(f"\n✅ Дорожка #{track_num} запланирована на следующий день!")
    print(f"   Использовано плит: {total_plates} шт")
    
    return result


def main():
    """Основная функция планирования на следующий день"""
    print("=" * 70)
    print("ПЛАНИРОВАНИЕ ДОРОЖЕК НА СЛЕДУЮЩИЙ ДЕНЬ")
    print("=" * 70)
    
    # Настройки
    length_m = 101  # метры
    width_m = 3.6   # метры
    min_load = 800  # кг/м2
    gap_cm = 1.0    # сантиметры
    num_tracks = 3  # количество дорожек
    
    # Файл с использованными плитами (вчерашний день)
    yesterday_excel = "раскладка_3_дорожки_101x3.6.xlsx"
    
    print(f"📅 Планируем на: {(datetime.now() + timedelta(days=1)).strftime('%d.%m.%Y')}")
    print(f"📋 Исключаем плиты из файла: {yesterday_excel}")
    print(f"📏 Параметры дорожек: {length_m}м × {width_m}м")
    print(f"⚖️  Минимальная нагрузка: {min_load} кг/м2")
    print()
    
    # Загружаем использованные плиты
    print("[ШАГ 1] Загружаю плиты, использованные вчера...")
    used_markings = load_used_plates_from_excel(yesterday_excel)
    
    # Загружаем доступные плиты
    print("\n[ШАГ 2] Загружаю доступные плиты из базы данных...")
    available_plates = load_available_plates(
        db_path='pb.db', 
        used_markings=used_markings, 
        min_load=min_load
    )
    
    if len(available_plates) == 0:
        print("❌ Нет доступных плит для планирования на следующий день!")
        return
    
    print(f"✅ Доступно плит для следующего дня: {len(available_plates)}")
    
    # Показываем распределение по неделям
    week_counts = available_plates['week'].value_counts().sort_index()
    print("\n📅 Распределение доступных плит по неделям формовки:")
    for week, count in week_counts.head(10).items():
        print(f"   Неделя {week}: {count} плит")
    if len(week_counts) > 10:
        print(f"   ... и ещё {len(week_counts)-10} недель")
    
    # Планируем дорожки
    print(f"\n[ШАГ 3] Планирую {num_tracks} дорожек...")
    results = []
    
    for track_num in range(1, num_tracks + 1):
        if len(available_plates) == 0:
            print(f"\n❌ Не осталось плит для дорожки #{track_num}")
            break
        
        result = plan_track_for_next_day(length_m, width_m, available_plates, gap_cm, track_num)
        
        if result:
            results.append(result)
            
            # Удаляем использованные плиты из доступных
            used_markings_this_track = set(p['marking'] for p in result['length_layout']['plates'])
            available_plates = available_plates[~available_plates['marking'].isin(used_markings_this_track)].reset_index(drop=True)
            
            print(f"   Осталось доступных плит: {len(available_plates)} шт")
        else:
            print(f"\n❌ Не удалось спланировать дорожку #{track_num}")
            break
    
    # Итоговая сводка
    if results:
        print(f"\n" + "=" * 70)
        print("ИТОГОВАЯ СВОДКА - ПЛАН НА СЛЕДУЮЩИЙ ДЕНЬ")
        print("=" * 70)
        print(f"✅ Запланировано дорожек: {len(results)} из {num_tracks}")
        
        total_plates = sum(r['total_plates'] for r in results)
        total_area = sum(r['length_m'] * r['width_m'] for r in results)
        
        print(f"📦 Всего использовано плит: {total_plates} шт")
        print(f"📐 Общая площадь покрытия: {total_area:.2f} м2")
        print(f"⚖️  Примерная масса: ~{total_plates * 0.6:.0f} тонн")
        
        print(f"\n📋 Детали по дорожкам:")
        for result in results:
            layout = result['length_layout']
            print(f"   • Дорожка #{result['track_num']}: {result['total_plates']} плит, "
                  f"покрытие {layout['achieved_m']:.2f}м, "
                  f"эффективность {layout['coverage_ratio']*100:.1f}%")
        
        print(f"\n💾 Создаю Excel файл с планом на следующий день...")
        
        # Создаём Excel файл
        filename = f"план_на_следующий_день_{len(results)}_дорожки.xlsx"
        
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            # Общая сводка
            summary_rows = []
            for result in results:
                layout = result['length_layout']
                used_weeks = [p['week'] for p in layout['plates']]
                unique_weeks = sorted(set(used_weeks))
                weeks_str = ', '.join(map(str, unique_weeks[:10]))
                
                summary_rows.append({
                    'Дорожка': f"#{result['track_num']}",
                    'Длина, м': result['length_m'],
                    'Ширина, м': result['width_m'],
                    'Всего плит': result['total_plates'],
                    'Армирование': layout['reinforcement'],
                    'Покрыто, м': layout['achieved_m'],
                    'Остаток, см': round(layout['gap_m'] * 100),
                    'Эффективность, %': round(layout['coverage_ratio'] * 100, 1),
                    'Недели формовки': weeks_str
                })
            
            summary_df = pd.DataFrame(summary_rows)
            summary_df.to_excel(writer, sheet_name='📊 План на завтра', index=False)
            
            # Каждая дорожка на отдельном листе
            for result in results:
                track_num = result['track_num']
                layout = result['length_layout']
                width_combo = result['width_combo']
                
                rows = []
                for i, plate in enumerate(layout['plates'], 1):
                    for j in range(width_combo['count']):
                        rows.append({
                            '№': len(rows) + 1,
                            'Позиция вдоль': i,
                            'Позиция поперёк': j + 1,
                            'Маркировка': plate['marking'],
                            'Длина, м': plate['length'],
                            'Ширина, м': plate['width'],
                            'Нагрузка, кг/м2': plate['load_capacity'],
                            'Армирование': plate['reinforcement'],
                            'Заказчик': plate['customer'],
                            'Неделя формовки': plate['week']
                        })
                
                df = pd.DataFrame(rows)
                sheet_name = f"Дорожка #{track_num}"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"✅ Файл создан: {filename}")
        print(f"   📑 Листов: {len(results) + 1} (план + {len(results)} дорожки)")
        
    else:
        print("\n❌ Не удалось спланировать ни одной дорожки на следующий день!")
        print("   Возможно, недостаточно доступных плит или нужно изменить параметры")


if __name__ == "__main__":
    main()



