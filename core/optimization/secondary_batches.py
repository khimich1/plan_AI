# -*- coding: utf-8 -*-
"""Helpers for secondary (z_sec) batch sizing in 2D optimization."""


def _batch_sizes_for_secondary_z_sec(qty: int, pieces: int) -> list[int]:
    """
    Разбиение qty строк выхода z_sec на батчи: один родительский остаток на батч
    длиной до pieces (ограничение cap_sec в ILP: sum z_sec <= x_sec * pieces).

    Examples: qty=3, pieces=2 -> [2, 1].
    """
    p = max(1, int(pieces or 1))
    sizes: list[int] = []
    offset = 0
    while offset < qty:
        b = min(p, qty - offset)
        sizes.append(b)
        offset += b
    return sizes
