"""
Программа для визуализации раскладки плит на дорожках
Создаёт красивую схему, которую можно распечатать
"""
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.backends.backend_pdf import PdfPages
import sqlite3
import pandas as pd
from datetime import datetime


def load_track_data(db_path: str = 'pb.db') -> pd.DataFrame:
    """Загружает данные о плитах из базы данных"""
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
    
    return df


def parse_plate_marking(marking: str):
    """Парсит маркировку плиты и извлекает размеры"""
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


def get_track_layouts():
    """Получает раскладку плит для всех дорожек"""
    df = load_track_data()
    
    # Фильтруем плиты с шириной 1.2 м
    plates_1_2m = []
    for idx, row in df.iterrows():
        params = parse_plate_marking(row['marking'])
        if params and params['width'] == 1.2:
            plates_1_2m.append({
                'marking': row['marking'],
                'length': params['length'],
                'width': params['width'],
                'load_capacity': params['load_capacity'],
                'customer': row['customer'],
                'reinforcement': row['reinforcement'],
                'week': row['week']
            })
    
    # Сортируем по неделе формовки
    plates_1_2m.sort(key=lambda x: x['week'] if x['week'] is not None else 999)
    
    tracks = []
    
    # Дорожка 1 - армирование 8.0
    track1_plates = [p for p in plates_1_2m if p['reinforcement'] == 8.0]
    track1_plates.sort(key=lambda x: x['week'] if x['week'] is not None else 999)
    
    if track1_plates:
        selected_plates = []
        current_track_length = 0
        target_track_length = 101.0
        
        for plate in track1_plates:
            if current_track_length + plate['length'] <= target_track_length:
                selected_plates.append(plate)
                current_track_length += plate['length']
            else:
                break
        
        tracks.append({
            'track_num': 1,
            'plates': selected_plates,
            'reinforcement': 8.0,
            'total_length': current_track_length,
            'color': '#FF6B6B'  # Красноватый
        })
    
    # Дорожка 2 - армирование 6.0
    track2_plates = [p for p in plates_1_2m if p['reinforcement'] == 6.0]
    track2_plates.sort(key=lambda x: x['week'] if x['week'] is not None else 999)
    
    if track2_plates:
        selected_plates = []
        current_track_length = 0
        target_track_length = 101.0
        
        for plate in track2_plates:
            if current_track_length + plate['length'] <= target_track_length:
                selected_plates.append(plate)
                current_track_length += plate['length']
            else:
                break
        
        tracks.append({
            'track_num': 2,
            'plates': selected_plates,
            'reinforcement': 6.0,
            'total_length': current_track_length,
            'color': '#4ECDC4'  # Голубоватый
        })
    
    # Дорожка 3 - армирование 4.0
    track3_plates = [p for p in plates_1_2m if p['reinforcement'] == 4.0]
    track3_plates.sort(key=lambda x: x['week'] if x['week'] is not None else 999)
    
    if track3_plates:
        selected_plates = []
        current_track_length = 0
        target_track_length = 101.0
        
        for plate in track3_plates:
            if current_track_length + plate['length'] <= target_track_length:
                selected_plates.append(plate)
                current_track_length += plate['length']
            else:
                break
        
        tracks.append({
            'track_num': 3,
            'plates': selected_plates,
            'reinforcement': 4.0,
            'total_length': current_track_length,
            'color': '#95E1D3'  # Зеленоватый
        })
    
    return tracks


def visualize_track_layout(save_pdf=True, save_png=True):
    """Создаёт визуализацию раскладки плит на всех дорожках"""
    
    print("=" * 70)
    print("ВИЗУАЛИЗАЦИЯ РАСКЛАДКИ ПЛИТ")
    print("=" * 70)
    
    # Получаем данные
    print("\n[ЗАГРУЗКА] Читаю данные из базы...")
    tracks = get_track_layouts()
    print(f"[УСПЕХ] Загружено {len(tracks)} дорожек")
    
    # Параметры визуализации
    track_width = 3.6  # Ширина дорожки в метрах
    spacing_between_tracks = 2.0  # Расстояние между дорожками
    
    # Создаём большой рисунок (размер А3 ландшафт для печати)
    fig, ax = plt.subplots(figsize=(16.5, 11.7))  # А3 в дюймах
    
    # Настройки отображения
    ax.set_xlim(0, 110)  # Длина + запас
    ax.set_ylim(0, len(tracks) * (track_width + spacing_between_tracks) + 2)
    ax.set_aspect('equal')
    
    # Убираем рамки
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_visible(False)
    ax.spines['left'].set_visible(False)
    
    # Заголовок
    current_date = datetime.now().strftime("%d.%m.%Y")
    plt.title(f'Раскладка плит на дорожках (дата: {current_date})', 
              fontsize=18, fontweight='bold', pad=20)
    
    # Отрисовываем каждую дорожку
    for i, track in enumerate(tracks):
        track_num = track['track_num']
        plates = track['plates']
        reinforcement = track['reinforcement']
        total_length = track['total_length']
        color = track['color']
        
        # Позиция дорожки по вертикали (снизу вверх)
        y_base = i * (track_width + spacing_between_tracks) + 1
        
        print(f"\n[ДОРОЖКА {track_num}] Отрисовываю раскладку...")
        print(f"   Армирование: {reinforcement}")
        print(f"   Плит: {len(plates)}")
        print(f"   Длина: {total_length:.2f} м")
        
        # Рисуем границу дорожки
        track_rect = patches.Rectangle(
            (0, y_base), 
            101, 
            track_width,
            linewidth=2,
            edgecolor='black',
            facecolor='#F5F5F5',
            alpha=0.3,
            linestyle='--'
        )
        ax.add_patch(track_rect)
        
        # Подпись дорожки
        ax.text(-1, y_base + track_width/2, 
                f'Дорожка {track_num}\nАрм. {reinforcement}',
                fontsize=10, fontweight='bold',
                ha='right', va='center')
        
        # Рисуем плиты
        current_x = 0
        for j, plate in enumerate(plates):
            plate_length = plate['length']
            
            # Рисуем прямоугольник плиты (3 плиты поперёк)
            for k in range(3):  # 3 плиты поперёк (каждая по 1.2м ширины)
                y_pos = y_base + k * 1.2
                
                rect = patches.Rectangle(
                    (current_x, y_pos),
                    plate_length,
                    1.2,  # Ширина одной плиты
                    linewidth=1,
                    edgecolor='black',
                    facecolor=color,
                    alpha=0.7
                )
                ax.add_patch(rect)
                
                # Подписи на плитах (только на средней плите для читаемости)
                if k == 1:  # Средняя плита
                    # Короткое название плиты
                    marking_short = plate['marking'].replace('ПБ ', '').split(' ')[0]
                    
                    # Неделя формовки
                    try:
                        week_text = f"Нед.{int(plate['week'])}" if plate['week'] and pd.notna(plate['week']) else "Срочно"
                    except:
                        week_text = "Срочно"
                    
                    # Если плита достаточно длинная, добавляем подпись
                    if plate_length >= 2.0:  # Увеличил минимальную длину для подписей
                        ax.text(current_x + plate_length/2, y_pos + 0.6,
                                f'{marking_short}\n{week_text}',
                                fontsize=6, ha='center', va='center',
                                fontweight='bold', color='black')
                    elif plate_length >= 1.0:  # Для средних плит - только неделя
                        ax.text(current_x + plate_length/2, y_pos + 0.6,
                                week_text,
                                fontsize=5, ha='center', va='center',
                                fontweight='bold', color='black')
                    else:
                        # Для очень коротких плит - только размер
                        ax.text(current_x + plate_length/2, y_pos + 0.6,
                                f'{plate_length}м',
                                fontsize=4, ha='center', va='center',
                                fontweight='bold', color='black')
            
            current_x += plate_length
        
        # Размеры дорожки
        ax.text(total_length + 1, y_base + track_width/2,
                f'{total_length:.1f}м',
                fontsize=9, fontweight='bold',
                ha='left', va='center',
                bbox=dict(boxstyle='round,pad=0.5', facecolor='white', edgecolor='black'))
    
    # Легенда
    legend_elements = []
    for track in tracks:
        legend_elements.append(
            patches.Patch(facecolor=track['color'], edgecolor='black',
                         label=f"Дорожка {track['track_num']} - Арм.{track['reinforcement']} ({len(track['plates'])} плит, {track['total_length']:.1f}м)")
        )
    
    ax.legend(handles=legend_elements, loc='upper right', fontsize=9)
    
    # Оси
    ax.set_xlabel('Длина (метры)', fontsize=12, fontweight='bold')
    ax.set_ylabel('Дорожки', fontsize=12, fontweight='bold')
    
    # Сетка
    ax.grid(True, alpha=0.3, linestyle=':', linewidth=0.5)
    
    plt.tight_layout()
    
    # Сохраняем файлы
    print("\n[СОХРАНЕНИЕ] Создаю файлы для печати...")
    
    files_created = []
    
    if save_pdf:
        pdf_filename = f'Раскладка_плит_{datetime.now().strftime("%Y%m%d")}.pdf'
        plt.savefig(pdf_filename, format='pdf', dpi=300, bbox_inches='tight')
        files_created.append(pdf_filename)
        print(f"   [PDF] {pdf_filename}")
    
    if save_png:
        png_filename = f'Раскладка_плит_{datetime.now().strftime("%Y%m%d")}.png'
        plt.savefig(png_filename, format='png', dpi=300, bbox_inches='tight')
        files_created.append(png_filename)
        print(f"   [PNG] {png_filename}")
    
    print("\n" + "=" * 70)
    print("ГОТОВО!")
    print("=" * 70)
    print(f"[УСПЕХ] Создано файлов: {len(files_created)}")
    for f in files_created:
        print(f"   • {f}")
    
    print("\n[СОВЕТ] Открой PDF файл для печати на принтере")
    print("        или PNG для просмотра на экране")
    
    plt.show()


def main():
    """Основная функция"""
    try:
        # Создаём визуализацию (PDF + PNG)
        visualize_track_layout(save_pdf=True, save_png=True)
        
    except Exception as e:
        print(f"\n[ОШИБКА] Что-то пошло не так: {e}")
        print("         Проверь, что файл pb.db находится в этой же папке")


if __name__ == "__main__":
    main()

