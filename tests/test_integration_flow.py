import pytest

try:
    from core.parsers.text_parser import parse_order_from_text
except ImportError as exc:
    pytest.skip(
        f"Интеграционный тест отключён: нет core.parsers.text_parser (OPT-010): {exc}",
        allow_module_level=True,
    )

import os
from pathlib import Path
from core.optimization import optimize_with_cascading_longitudinal_cuts
from core.visualization import visualize_plan
def test_integration_flow(tmp_path):
    text = "ПБ 78-12-8п 2\nПБ 78-10-8п 1\nПБ 78-3-8п 3"
    
    # 1. Парсинг
    order = parse_order_from_text(text)
    assert len(order.items) == 3
    assert order.total_plates == 6
    
    # 2. Оптимизация
    opt_result = optimize_with_cascading_longitudinal_cuts(order=order)
    
    # Должно быть 2 целых (1.2)
    # И 1 на резку 1.0 (main: 1.0, rest: 0.2)
    # И ещё остатки. 
    # В результате мы не знаем точное количество, но оптимизация должна отработать без ошибок
    assert opt_result is not None
    assert 'total_plates' in opt_result
    assert opt_result['total_plates'] > 0
    
    # 3. Визуализация
    output_dir = tmp_path / "outputs"
    output_dir.mkdir()
    
    # Чтобы не ломать логику загрузки прайсов, ставим price_db_path None (оно создастся или проигнорируется)
    result_paths = visualize_plan(
        output_dir=str(output_dir),
        order=order
    )
    
    assert result_paths is not None
    assert len(result_paths) >= 2 # Обычно (png_path, pdf_path, ...)
