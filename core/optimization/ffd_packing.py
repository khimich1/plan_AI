"""
First-Fit-Decreasing packing of plate pieces onto production tracks (bins).

Pure stdlib; no ILP / PuLP / order-dispatch dependencies.
"""

from __future__ import annotations

from dataclasses import dataclass, field

__all__ = [
    "Piece",
    "Track",
    "pack_tracks",
    "first_fit_decreasing",
    "optimize_tracks",
]


@dataclass
class Piece:
    """Кусок плиты для укладки в дорожку"""

    length_m: float
    qty: int
    kind: str  # 'standard' | 'addon'
    load_class: float
    width_m: float = 1.196


@dataclass
class Track:
    """Дорожка (линия производства)"""

    width_m: float = 1.196
    total_m: float = 0.0
    pieces: list[Piece] = field(default_factory=list)
    leftover_m: float = 0.0


def pack_tracks(
    pieces: list[Piece],
    stock_len_m: float = 9.88,
) -> list[Track]:
    """
    Алгоритм First Fit Decreasing для оптимизации раскроя
    Минимизирует количество дорожек (плит-заготовок)

    Args:
        pieces: Список Piece объектов (куски для размещения)
        stock_len_m: Длина заготовки (максимальная длина плиты)

    Returns:
        Список Track объектов (дорожек) с размещёнными кусками
    """
    pool: list[Track] = []

    # Сортируем по убыванию длины (FFD алгоритм)
    sorted_pieces = sorted(pieces, key=lambda x: x.length_m, reverse=True)

    # Развёртываем количество в отдельные элементы
    expanded: list[Piece] = []
    for p in sorted_pieces:
        for _ in range(p.qty):
            expanded.append(Piece(p.length_m, 1, p.kind, p.load_class, p.width_m))

    # Размещаем каждый кусок
    for piece in expanded:
        placed = False

        # Пробуем поместить в существующие дорожки
        for track in pool:
            if track.total_m + piece.length_m <= stock_len_m:
                track.pieces.append(piece)
                track.total_m += piece.length_m
                placed = True
                break

        # Если не поместился, создаём новую дорожку
        if not placed:
            track = Track()
            track.pieces.append(piece)
            track.total_m = piece.length_m
            pool.append(track)

    # Вычисляем остатки
    for track in pool:
        track.leftover_m = stock_len_m - track.total_m

    return pool


# Обратное имя для существующих вызовов / тестов
first_fit_decreasing = pack_tracks


def optimize_tracks(
    items: list,
    stock_len_m: float = 9.88,
) -> dict:
    """
    Оптимизирует размещение плит в дорожки

    Args:
        items: Список позиций [{'length_m': float, 'qty': int, 'kind': str, 'load_class': float}, ...]
        stock_len_m: Длина заготовки (максимальная длина)

    Returns:
        Словарь с результатами оптимизации
    """
    pieces: list[Piece] = []

    for item in items:
        pieces.append(
            Piece(
                length_m=item.get("length_m", 0),
                qty=item.get("qty", 1),
                kind=item.get("kind", "standard"),
                load_class=item.get("load_class", 8.0),
                width_m=item.get("width_m", 1.196),
            )
        )

    tracks = pack_tracks(pieces, stock_len_m)

    # Статистика
    total_tracks = len(tracks)
    total_used = sum(t.total_m for t in tracks)
    total_leftover = sum(t.leftover_m for t in tracks)
    efficiency = (total_used / (total_tracks * stock_len_m) * 100) if total_tracks > 0 else 0

    return {
        "tracks": tracks,
        "total_tracks": total_tracks,
        "total_used_m": round(total_used, 2),
        "total_leftover_m": round(total_leftover, 2),
        "efficiency_pct": round(efficiency, 1),
        "stock_length_m": stock_len_m,
    }
