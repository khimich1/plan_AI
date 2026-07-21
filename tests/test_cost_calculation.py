#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тесты для расчета себестоимости плит
"""

import pytest

try:
    import unittest
    import os
    import tempfile
    import sqlite3
    from core.cost_db import (
        init_cost_schema, init_default_constants,
        get_constant, get_concrete_norms, get_reinforcement_norms, get_izoform_norm
    )
    from core.cost_calculation import (
        parse_plate_name, calculate_plate_volume,
        calculate_plate_cost, calculate_concrete_cost,
        calculate_reinforcement_cost, calculate_loops_cost,
        calculate_izoform_cost
    )
except ImportError as exc:
    pytest.skip(
        "Архивные тесты себестоимости: отсутствуют core.cost_db / core.cost_calculation "
        f"(OPT-010): {exc}",
        allow_module_level=True,
    )


class TestCostDB(unittest.TestCase):
    """Тесты для работы с БД констант"""
    
    def setUp(self):
        """Создаем временную БД для тестов"""
        self.temp_db = tempfile.NamedTemporaryFile(delete=False, suffix='.db')
        self.temp_db.close()
        self.db_path = self.temp_db.name
        init_default_constants(self.db_path)
    
    def tearDown(self):
        """Удаляем временную БД"""
        if os.path.exists(self.db_path):
            os.unlink(self.db_path)
    
    def test_init_schema(self):
        """Тест инициализации схемы БД"""
        init_cost_schema(self.db_path)
        
        conn = sqlite3.connect(self.db_path)
        try:
            cur = conn.cursor()
            cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = [row[0] for row in cur.fetchall()]
            
            self.assertIn('cost_constants', tables)
            self.assertIn('concrete_norms', tables)
            self.assertIn('reinforcement_norms', tables)
            self.assertIn('izoform_norms', tables)
        finally:
            conn.close()
    
    def test_get_constant(self):
        """Тест получения константы"""
        value = get_constant('cement_price_per_kg', self.db_path)
        self.assertIsNotNone(value)
        self.assertEqual(value, 360.0)
        
        value = get_constant('loop_d18_price', self.db_path)
        self.assertEqual(value, 286.0)
    
    def test_get_concrete_norms(self):
        """Тест получения норм бетона"""
        norms = get_concrete_norms('М400', self.db_path)
        self.assertIsNotNone(norms)
        self.assertEqual(norms['cement_kg_per_m3'], 360.0)
        self.assertEqual(norms['sand_m3_per_m3'], 0.62)
        self.assertEqual(norms['gravel_m3_per_m3'], 2.065)
        
        norms = get_concrete_norms('М500', self.db_path)
        self.assertIsNotNone(norms)
        self.assertEqual(norms['cement_kg_per_m3'], 380.0)
    
    def test_get_reinforcement_norms(self):
        """Тест получения норм армирования"""
        norms = get_reinforcement_norms(6, self.db_path)
        self.assertIsNotNone(norms)
        self.assertEqual(norms['wire_kg_per_m3'], 6.0)
        
        norms = get_reinforcement_norms(8, self.db_path)
        self.assertIsNotNone(norms)
        self.assertEqual(norms['wire_kg_per_m3'], 6.6)
    
    def test_get_izoform_norm(self):
        """Тест получения нормы изоформа"""
        norm = get_izoform_norm(0.28, self.db_path)
        self.assertIsNotNone(norm)
        self.assertEqual(norm, 0.072)
        
        norm = get_izoform_norm(0.35, self.db_path)
        self.assertIsNotNone(norm)
        self.assertEqual(norm, 0.1224)


class TestCostCalculation(unittest.TestCase):
    """Тесты для расчета себестоимости"""
    
    def setUp(self):
        """Создаем временную БД для тестов"""
        self.temp_db = tempfile.NamedTemporaryFile(delete=False, suffix='.db')
        self.temp_db.close()
        self.db_path = self.temp_db.name
        init_default_constants(self.db_path)
    
    def tearDown(self):
        """Удаляем временную БД"""
        if os.path.exists(self.db_path):
            os.unlink(self.db_path)
    
    def test_parse_plate_name(self):
        """Тест парсинга названия плиты"""
        result = parse_plate_name("ПБ 17-12-6")
        self.assertIsNotNone(result)
        self.assertEqual(result['length_dm'], 17)
        self.assertEqual(result['length_m'], 1.7)
        self.assertEqual(result['width_dm'], 12)
        self.assertEqual(result['width_m'], 1.2)
        self.assertEqual(result['load_code'], 6)
        self.assertEqual(result['concrete_grade'], 'М400')
        
        result = parse_plate_name("ПБ 20-12-8")
        self.assertIsNotNone(result)
        self.assertEqual(result['length_m'], 2.0)
        self.assertEqual(result['load_code'], 8)
        
        result = parse_plate_name("ПБ 30-12-12")
        self.assertIsNotNone(result)
        self.assertEqual(result['concrete_grade'], 'М500')
        
        result = parse_plate_name("неправильное название")
        self.assertIsNone(result)
    
    def test_calculate_plate_volume(self):
        """Тест расчета объема плиты"""
        volume = calculate_plate_volume(1.7, 1.2)
        expected = 1.7 * 1.2 * 0.22
        self.assertAlmostEqual(volume, expected, places=4)
        
        volume = calculate_plate_volume(2.0, 1.2)
        expected = 2.0 * 1.2 * 0.22
        self.assertAlmostEqual(volume, expected, places=4)
    
    def test_calculate_concrete_cost(self):
        """Тест расчета стоимости бетона"""
        volume = 0.2805
        cost = calculate_concrete_cost(volume, 'М400', self.db_path)
        
        # Проверяем, что стоимость больше нуля
        self.assertGreater(cost, 0)
        
        # Проверяем расчет для М500
        cost_m500 = calculate_concrete_cost(volume, 'М500', self.db_path)
        self.assertGreater(cost_m500, 0)
        # М500 должен быть дороже М400
        self.assertGreater(cost_m500, cost)
    
    def test_calculate_reinforcement_cost(self):
        """Тест расчета стоимости армирования"""
        volume = 0.2805
        
        cost_6 = calculate_reinforcement_cost(volume, 6, self.db_path)
        cost_8 = calculate_reinforcement_cost(volume, 8, self.db_path)
        
        self.assertGreaterEqual(cost_8, cost_6)
    
    def test_calculate_loops_cost(self):
        """Тест расчета стоимости петель"""
        cost = calculate_loops_cost(6, self.db_path)
        self.assertEqual(cost, 286.0)
    
    def test_calculate_izoform_cost(self):
        """Тест расчета стоимости изоформа"""
        volume = 0.2805
        cost = calculate_izoform_cost(volume, self.db_path)
        
        # Для объема 0.28 м³ норма изоформа = 0.072 кг
        # Цена = 95 руб/кг
        # Ожидаемая стоимость = 0.072 * 95 = 6.84 руб
        expected = 0.072 * 95.0
        self.assertAlmostEqual(cost, expected, places=2)
    
    def test_calculate_plate_cost_full(self):
        """Тест полного расчета себестоимости плиты"""
        result = calculate_plate_cost("ПБ 17-12-6", self.db_path)
        
        self.assertIsNotNone(result)
        self.assertEqual(result['plate_name'], "ПБ 17-12-6")
        self.assertIn('parameters', result)
        self.assertIn('components', result)
        self.assertIn('total_cost', result)
        self.assertIn('breakdown', result)
        
        # Проверяем структуру components
        components = result['components']
        self.assertIn('concrete', components)
        self.assertIn('reinforcement', components)
        self.assertIn('loops', components)
        self.assertIn('izoform', components)
        
        # Проверяем, что все компоненты положительные
        for component, cost in components.items():
            self.assertGreaterEqual(cost, 0, f"Компонент {component} должен быть >= 0")
        
        # Проверяем, что общая стоимость равна сумме компонентов
        total = sum(components.values())
        self.assertAlmostEqual(result['total_cost'], total, places=2)
        
        # Проверяем параметры
        params = result['parameters']
        self.assertEqual(params['length_m'], 1.7)
        self.assertEqual(params['width_m'], 1.2)
        self.assertEqual(params['load_code'], 6)
    
    def test_calculate_plate_cost_different_loads(self):
        """Тест расчета для разных нагрузок"""
        result_6 = calculate_plate_cost("ПБ 20-12-6", self.db_path)
        result_8 = calculate_plate_cost("ПБ 20-12-8", self.db_path)
        result_10 = calculate_plate_cost("ПБ 20-12-10", self.db_path)
        
        self.assertIsNotNone(result_6)
        self.assertIsNotNone(result_8)
        self.assertIsNotNone(result_10)
        
        # Проверяем, что объем одинаковый
        self.assertEqual(result_6['volume_m3'], result_8['volume_m3'])
        self.assertEqual(result_8['volume_m3'], result_10['volume_m3'])
        
        # Проверяем, что себестоимость увеличивается с нагрузкой
        # (из-за увеличения армирования)
        self.assertGreaterEqual(
            result_8['components']['reinforcement'],
            result_6['components']['reinforcement']
        )
    
    def test_calculate_plate_cost_different_lengths(self):
        """Тест расчета для разных длин"""
        result_17 = calculate_plate_cost("ПБ 17-12-6", self.db_path)
        result_20 = calculate_plate_cost("ПБ 20-12-6", self.db_path)
        result_30 = calculate_plate_cost("ПБ 30-12-6", self.db_path)
        
        self.assertIsNotNone(result_17)
        self.assertIsNotNone(result_20)
        self.assertIsNotNone(result_30)
        
        # Проверяем, что объем увеличивается с длиной
        self.assertLess(result_17['volume_m3'], result_20['volume_m3'])
        self.assertLess(result_20['volume_m3'], result_30['volume_m3'])
        
        # Проверяем, что себестоимость увеличивается с объемом
        self.assertLess(result_17['total_cost'], result_20['total_cost'])
        self.assertLess(result_20['total_cost'], result_30['total_cost'])
    
    def test_calculate_plate_cost_invalid_name(self):
        """Тест с неправильным названием плиты"""
        result = calculate_plate_cost("неправильное название", self.db_path)
        self.assertIsNone(result)
        
        result = calculate_plate_cost("ПБ 17", self.db_path)
        self.assertIsNone(result)
    
    def test_calculate_plate_cost_breakdown(self):
        """Тест детальной разбивки себестоимости"""
        result = calculate_plate_cost("ПБ 17-12-6", self.db_path)
        
        self.assertIsNotNone(result)
        breakdown = result['breakdown']
        
        self.assertIn('concrete', breakdown)
        self.assertIn('reinforcement', breakdown)
        
        concrete = breakdown['concrete']
        self.assertIn('volume_m3', concrete)
        self.assertIn('grade', concrete)
        self.assertIn('cement_kg', concrete)
        self.assertIn('sand_m3', concrete)
        self.assertIn('gravel_m3', concrete)
        
        reinforcement = breakdown['reinforcement']
        self.assertIn('load_code', reinforcement)
        self.assertIn('wire_kg', reinforcement)
        self.assertIn('cable_cost', reinforcement)


class TestIntegration(unittest.TestCase):
    """Интеграционные тесты"""
    
    def setUp(self):
        """Создаем временную БД для тестов"""
        self.temp_db = tempfile.NamedTemporaryFile(delete=False, suffix='.db')
        self.temp_db.close()
        self.db_path = self.temp_db.name
        init_default_constants(self.db_path)
    
    def tearDown(self):
        """Удаляем временную БД"""
        if os.path.exists(self.db_path):
            os.unlink(self.db_path)
    
    def test_full_workflow(self):
        """Тест полного рабочего процесса"""
        # 1. Парсим название
        params = parse_plate_name("ПБ 17-12-6")
        self.assertIsNotNone(params)
        
        # 2. Рассчитываем объем
        volume = calculate_plate_volume(params['length_m'], params['width_m'])
        self.assertGreater(volume, 0)
        
        # 3. Рассчитываем компоненты
        concrete_cost = calculate_concrete_cost(
            volume, params['concrete_grade'], self.db_path
        )
        reinforcement_cost = calculate_reinforcement_cost(
            volume, params['load_code'], self.db_path
        )
        loops_cost = calculate_loops_cost(params['load_code'], self.db_path)
        izoform_cost = calculate_izoform_cost(volume, self.db_path)
        
        # 4. Полный расчет
        result = calculate_plate_cost("ПБ 17-12-6", self.db_path)
        self.assertIsNotNone(result)
        
        # 5. Проверяем согласованность
        self.assertAlmostEqual(
            result['components']['concrete'],
            concrete_cost,
            places=2
        )
        self.assertAlmostEqual(
            result['components']['reinforcement'],
            reinforcement_cost,
            places=2
        )
        self.assertAlmostEqual(
            result['components']['loops'],
            loops_cost,
            places=2
        )
        self.assertAlmostEqual(
            result['components']['izoform'],
            izoform_cost,
            places=2
        )


if __name__ == '__main__':
    unittest.main()

