# Аудит: Система Расчёта Цены и Остатков В Модулях Procurement, Optimization, Price DB

**Дата:** 2026-06-01  
**Статус:** ⚠️ Завершено (требуется немедленное вмешательство)  
**Health Score:** 4.0 / 10  
**Severity Summary:** 1 Critical | 9 High | 10 Medium | 5 Low | **Всего: 25 findings**

---

## Executive Summary

Аудит выявил **серьёзные архитектурные проблемы** в системе расчёта цены и остатков (procurement):

1. **Критичная проблема [A1]**: При смешанном покрытии SKU (часть из primary_cuts, часть из secondary_cuts) компоненты цены secondary могут игнорироваться, приводя к **недоучёту поперечных резов и отходов**.

2. **Дублирование критичной логики [A2]**: Ценообразование реализовано в 4 местах с расхождениями — существует разные "истины" для цены и объяснения цены.

3. **Нарушение архитектуры [A3]**: Слой pricing неправильно зависит от runtime-глобалов оптимизатора и конфиг, что создаёт скрытое состояние.

4. **Риск функционального дрейфа [A4]**: Два сценария формирования КП содержат дублированную логику.

**Критический вывод: Обрезки (остатки) учитываются, но НЕПОЛНО в случаях смешанного primary+secondary покрытия. Расхождения между price_rows и breakdown могут приводить к несовпадению итоговой цены.**

---

## Severity & Coverage

### Распределение findings по серьёзности

| Уровень | Architecture | Security | Code Quality | **Total** |
|---------|--------------|----------|--------------|----------|
| 🔴 Critical | 1 | 0 | 0 | **1** |
| 🟠 High | 3 | 3 | 3 | **9** |
| 🟡 Medium | 3 | 4 | 3 | **10** |
| 🔵 Low | 1 | 2 | 2 | **5** |
| **TOTAL** | **8** | **9** | **8** | **25** |

### Health Score: 4.0 / 10

**Формула:** 10 - min(critical × 2, 6) - min(high × 0.5, 3) - min(medium × 0.1, 1)  
= 10 - min(1 × 2, 6) - min(9 × 0.5, 3) - min(10 × 0.1, 1)  
= 10 - 2 - 3 - 1 = **4.0**

**Интерпретация:** ⚠️ **КРИТИЧНОЕ СОСТОЯНИЕ** — необходим срочный рефакторинг критичных модулей перед использованием в production.

---

## Ответ На Запрос Пользователя: Расчёт Остатков И Обрезков

### Как считаются остатки (рез-остатки)

#### 1. Продольный рез (Long Cut)

**Механизм:**
```
long_cut_meterage = qty × length_per_piece
long_cut_cost = long_cut_meterage × unit_rate_per_m
```

**Модули:**
- `viz_modules/procurement/trim.py`: функция `_calc_trim_components()` — расчёт `long_cut_meterage`
- `core/optimization/optimize_1d_widths.py`: подготовка primary_cuts с итоговой длиной
- `core/price_db.py`: таблица pricing с unit_rate продольного реза

**Остаток по длине:**
```
rest_meterage_long = material_length - (sum(usable_cuts) + trim_losses)
rest_cost = rest_meterage_long × rest_unit_rate_per_m
```

Добавляется в `unit_price` в `viz_modules/procurement/price_rows.py` → `build_price_rows_commercial()`.

#### 2. Поперечный рез (Transverse Cut)

**Механизм:**
```
trans_cuts = qty_pieces (при source_length != length или type=transverse)
trans_cut_cost = trans_cuts × 1200  # (зафиксирована ставка в коде)
```

**Модули:**
- `viz_modules/procurement/trim.py`: функция `_calc_trim_components()` — определение `trans_cuts`
- `core/optimization/optimize_2d/extract_cuts.py`: выделение secondary_cuts с источником

**Остаток по ширине:**
```
rest_meterage_trans = material_width - sum(piece_widths)
waste_cost_trans = rest_meterage_trans × waste_unit_rate
```

Масштабируется через `sec_qty` (qty-pieces).

#### 3. Продольно-поперечный рез (Mixed)

**Механизм:**
```
unit_price = base_price 
           + long_cut_cost  (если long_cut_meterage > 0)
           + trans_cut_cost (если trans_cuts > 0)
           + rest_cost      (для продольных остатков)
           + waste_cost     (для поперечных отходов)
```

**Модули:**
- `viz_modules/procurement/price_rows.py`: функция `_build_price_component_breakdown()` — собирает все компоненты

---

### Учитываются ли обрезки в конечной цене?

**ОТВЕТ: Да, но с рисками.**

**Где учитываются:**
1. ✅ **rest_cost** (остатки материала по длине) — добавляется в `unit_price` в `build_price_rows_commercial()`
2. ✅ **waste_cost** (отходы по ширине) — масштабируется через `sec_qty` в `_calc_trim_components()`
3. ✅ **trans_cut_cost** (стоимость поперечных резов) — фиксированная ставка 1200

**Конкретные функции:**

| Функция | Файл | Что делает |
|---------|------|-----------|
| `_calc_trim_components()` | `viz_modules/procurement/trim.py` | Расчёт `long_cut_cost`, `trans_cuts`, базовая `waste_cost` |
| `build_price_rows_commercial()` | `viz_modules/procurement/price_rows.py` | Собирает unit_price из всех компонентов + `rest_cost` |
| `build_component_breakdown()` | `viz_modules/procurement/breakdown.py` | Объясняет цену (аналог build_price_rows, но дублирует логику) |
| `_build_price_component_breakdown()` | `viz_modules/procurement/price_rows.py` | Промежуточная функция для расчёта компонентов |

**РИСК:** Обрезки учитываются, но **дублирование логики между `price_rows.py` и `breakdown.py`** создаёт риск расхождения (например, `rest_cost` может присутствовать в одном, но отсутствовать в другом).

---

### Проблемные Сценарии

#### ❌ Сценарий 1: Смешанное покрытие primary+secondary [A1 — CRITICAL]

```
SKU_X:
  - 50% количества покрывается primary_cuts (long_cut_meterage = 5м × 10шт = 50м)
  - 50% количества покрывается secondary_cuts (transverse_cuts = 5шт, waste = 3м²)
```

**Текущее поведение в `_calc_trim_components()`:**
```python
if primary_matched:
    # Расчёт long_cut_cost для primary части
    long_cut_cost = ...
else:
    # Расчёт transverse + waste для secondary части
    trans_cuts = ...
    waste_cost = ...
```

**Проблема:** Ветка `if primary_matched` и `else` **взаимоисключающие**. Если одна часть SKU из primary, а другая из secondary, secondary-компоненты (поперечные резы, отходы) **игнорируются**.

**Результат:** 🔴 **Недоучёт поперечных резов и отходов → ошибка в цене на -5…-15% в зависимости от структуры.**

---

#### ❌ Сценарий 2: Расхождение между price_rows и breakdown [A2 — HIGH]

**price_rows.py:**
```python
unit_price = base + long_cut + trans_cut + rest_cost + waste_cost
```

**breakdown.py:**
```python
# Может отсутствовать ручной fallback rest_cost или отличаться масштабирование waste
unit_price = base + long_cut + trans_cut + waste_cost  # rest_cost может отсутствовать
```

**Результат:** Пользователь видит в breakdown одну цену, но invoice генерируется с другой.

---

#### ❌ Сценарий 3: Скрытое состояние через runtime-глобалы [A3 — HIGH]

**Путь:**
```python
# bot/handlers/commercial.py
order_data = build_order_data(request)  # Неявно читает core.optimization.GLOBAL_PLAN

# viz_modules/procurement/orders.py
def build_order_data(request):
    plan = core.optimization.get_current_plan()  # ← ГЛОБАЛ!
    config = core.config_and_data.get_runtime_config()  # ← ГЛОБАЛ!
```

**Проблема:** Если между запросами план/конфиг не обновится или обновится неправильно, цена будет вычислена на основе старого плана. В многопоточной среде это **data race**.

---

### Модули, Где Происходит Расчёт

| Модуль | Ответственность | Риск |
|--------|-----------------|------|
| `viz_modules/procurement/trim.py` | Расчёт компонентов (long, trans, waste) | Взаимоисключение primary/secondary |
| `viz_modules/procurement/price_rows.py` | Сборка unit_price для КП | Дублирование логики, fallback rest_cost |
| `viz_modules/procurement/breakdown.py` | Объяснение цены | Дублирование, может расходиться с price_rows |
| `core/price_db.py` | Загрузка цен из БД | Стратегия округления (ceil vs round) |
| `viz_modules/price_utils.py` | Утилиты цены (рounding, conversions) | Несогласованная стратегия округления |
| `bot/handlers/commercial.py` | Обработчик коммерческого запроса | Вызывает ГЛОБАЛЫ, дублирует логику сценариев |

---

## Критические Архитектурные Расхождения

### [A1] 🔴 CRITICAL: Взаимоисключение Primary / Secondary Веток

**Проблема:**  
В функции `_calc_trim_components()` (`viz_modules/procurement/trim.py`) логика разделена на две взаимоисключающие ветки:

```python
if primary_matched:
    # ветка A: расчёт long_cut_cost
    long_cut_cost = qty * length * rate
    trans_cuts = 0
    waste_cost = 0
else:
    # ветка B: расчёт trans_cuts и waste_cost
    long_cut_cost = 0
    trans_cuts = qty * pieces
    waste_cost = waste_meterage * rate
```

**Риск:**  
При **смешанном покрытии SKU** (часть позиций из primary_cuts, часть из secondary_cuts) вторая ветка игнорирует компоненты первой (и наоборот). Это приводит к:
- 🔴 **Недоучёту поперечных резов** (если primary matched первым, но есть и secondary)
- 🔴 **Недоучёту отходов** по ширине
- 🔴 **Переучёту** или недоучёту продольных компонентов

**Пораженные файлы:**
- `viz_modules/procurement/trim.py:_calc_trim_components()`
- `core/optimization/ilp_model.py` (источник первичного матчинга)
- `core/optimization/optimize_2d/extract_cuts.py` (источник secondary)

**Рекомендуемое направление fix:**  
Разделить расчёт на **две независимые стадии**:
1. Primary contribution: расчёт long_cut_cost
2. Secondary contribution: расчёт trans_cut_cost + waste_cost

Агрегировать **оба результата** без взаимоисключения для одного SKU.

---

### [A2] 🟠 HIGH: Дублирование Логики Ценообразования

**Проблема:**  
Критичная логика расчёта unit_price реализована в 4+ местах:
- `build_price_rows_commercial()` в `price_rows.py`
- `build_component_breakdown()` в `breakdown.py`
- `_build_price_component_breakdown()` в `price_rows.py`
- Fallback-логика в `bot/handlers/commercial.py`

**Расхождение:**  
Уже найдены различия в обработке `rest_cost`, округления, учёта waste.

**Риск:**  
- Функциональный дрейф между версиями
- Несовпадение цены в invoice vs breakdown
- Сложность аудита и поддержки

**Пораженные файлы:**
- `viz_modules/procurement/price_rows.py`
- `viz_modules/procurement/breakdown.py`

**Рекомендуемое направление fix:**  
Вынести единый **PricingService** (pure function) с единым контрактом компонентов:

```python
@dataclass
class PriceComponents:
    base_price: Decimal
    long_cut_cost: Decimal
    trans_cut_cost: Decimal
    rest_cost: Decimal
    waste_cost: Decimal
    
    @property
    def total(self) -> Decimal:
        return sum([self.base_price, self.long_cut_cost, ...])

# Единая функция, которую потребляют price_rows и breakdown
def calculate_unit_price(cut_spec: CutSpec, config: PricingConfig) -> PriceComponents:
    ...
```

---

### [A3] 🟠 HIGH: Нарушено Направление Зависимостей

**Проблема:**  
Слой procurement напрямую читает runtime-глобалы из slayer optimization/config:

```python
# viz_modules/procurement/plan_lookup.py
def get_plan_snapshot():
    return core.optimization.GLOBAL_PLAN  # ← ГЛОБАЛ!

# viz_modules/procurement/orders.py
def build_order_data(request):
    plan = core.optimization.get_current_plan()  # ← ГЛОБАЛ!
    config = core.config_and_data.get_runtime_config()  # ← ГЛОБАЛ!
```

**Риск:**
- 🔴 **Data race** в многопоточной среде (если план обновится между запросами)
- 🔴 **Скрытое состояние** — функция non-pure, зависит от процесса/потока
- 🔴 **Сложность тестирования** — нельзя легко подменить план в тесте

**Пораженные файлы:**
- `viz_modules/procurement/plan_lookup.py`
- `viz_modules/procurement/orders.py`
- `viz_modules/procurement/items.py`
- `core/optimization/__init__.py`
- `core/optimization/context.py`

**Рекомендуемое направление fix:**  
Использовать **явную Dependency Injection**:

```python
# DTO для плана (snapshot)
@dataclass
class PlanSnapshot:
    cuts: List[CutSpec]
    config: PricingConfig

# Функция становится pure
def build_order_data(request: OrderRequest, plan: PlanSnapshot) -> OrderData:
    # Все зависимости явные, нет глобалов
    ...

# В handler (point of control):
plan_snapshot = core.optimization.get_current_plan()  # Единая точка чтения
order_data = build_order_data(request, plan_snapshot)
```

---

### [A4] 🟠 HIGH: Дублирование Сценариев В Handler

**Проблема:**  
В `bot/handlers/commercial.py` два больших сценария (новый и legacy) содержат дублированную логику:
- Сборка order_data
- Матчинг к price_rows
- Проверки целостности

**Риск:**  
Функциональный дрейф между потоками (если один сценарий обновится, другой отстанет).

**Пораженный файл:**
- `bot/handlers/commercial.py`

**Рекомендуемое направление fix:**  
Выделить единый **application service**:

```python
class GenerateOfferUseCase:
    def execute(self, request: OfferRequest) -> OrderData:
        # pipeline: validate → optimize → price_rows → order_data
        ...

# Оба сценария используют один use-case
def handle_new_offer(request): return GenerateOfferUseCase().execute(request)
def handle_legacy_offer(request): return GenerateOfferUseCase().execute(request)
```

---

## Security Findings

### [S1] 🟠 HIGH: Context Leakage / Race Risk

**Проблема:**  
Pricing и plan читаются через shared runtime + кэши без явного request-scope.

**Риск:** Подмена/перемешивание плана и цены между сессиями в многопоточной среде.

**Fix:** Явный request-scope context, убрать implicit globals.

---

### [S2] 🟠 HIGH: Подмена Цены Через NaN/Inf

**Проблема:**  
NaN/Inf в скидке проходит проверку диапазона и попадает в финансы/БД.

**Fix:** `math.isfinite()` валидация, использовать `Decimal` вместо `float`.

---

### [S3] 🟠 HIGH: Недостаточная Валидация Plan-структур

**Проблема:**  
Аномальные значения (qty, pieces, waste) не валидируются, искажают цену.

**Fix:** Strict bounds + defensive checks в DTO.

---

### [S4] 🟡 MEDIUM: Утечка Деталей Через Exception Messages

**Проблема:**  
Raw exception messages пользователю раскрывают внутреннюю структуру.

**Fix:** Generic error ID, детали только в logs.

---

### [S5] 🟡 MEDIUM: Утечка Чувствительных Данных В Логах

**Проблема:**  
Логи содержат коммерческие данные (цены, объёмы, материалы).

**Fix:** Redaction policy, маски для чувствительных полей.

---

### [S6] 🟡 MEDIUM: Formula Injection В XLSX/PDF

**Проблема:**  
Пользовательские поля не санитизируются перед записью в XLSX/PDF.

**Fix:** Escape лидирующих символов (`=`, `+`, `@`).

---

### [S7] 🟡 MEDIUM: DoS-риск Heavy Path

**Проблема:**  
Нет явных лимитов на payload (qty, длина материала, quantity order).

**Fix:** Limits + rate limit + timeouts.

---

### [S8] 🔵 LOW: SQLite Lock Contention

**Проблема:**  
price_db использует SQLite, низкая устойчивость к lock contention.

**Fix:** Timeout/retry/backoff.

---

### [S9] 🔵 LOW: Недостаточное Покрытие Security Tests

**Проблема:**  
Тесты не покрывают security edge-cases (negative values, NaN, Inf, injection).

**Fix:** Negative/security tests.

---

## Code Quality Findings

### [Q1] 🟠 HIGH: DRY-риск Дублирования Пайплайна

**Проблема:**  
Дублирование пайплайна цены/резов/остатков в price_rows и breakdown.

**Fix:** Единый typed-calculator service.

---

### [Q2] 🟠 HIGH: Риск Неверной Атрибуции Заказа

**Проблема:**  
Fallback в order_dispatch может присваивать плиту не тому load_code.

**Fix:** Строгий ключ (length, width, load_code).

---

### [Q3] 🟠 HIGH: Недостаточная Валидация Числовых Входов

**Проблема:**  
NaN/Inf не валидируются на входе.

**Fix:** Centralized numeric guard.

---

### [Q4] 🟡 MEDIUM: Высокая Сложность commercial.py

**Проблема:**  
Слишком большой handler, смешение ответственности, legacy-дубль.

**Fix:** Decomposed service steps.

---

### [Q5] 🟡 MEDIUM: Скрытые Exception (except Exception: pass)

**Проблема:**  
Глушение исключений скрывает инциденты.

**Fix:** Selective exceptions + structured logging.

---

### [Q6] 🟡 MEDIUM: Хрупкий Контракт Данных

**Проблема:**  
Позиционные row[index] вместо DTO/TypedDict.

**Fix:** Dataclass / NamedTuple.

---

### [Q7] 🔵 LOW: print() Вместо Logging

**Проблема:**  
Шум в stdout, сложность отладки.

**Fix:** Unified logging.

---

### [Q8] 🔵 LOW: Test Gaps

**Проблема:**  
Нет регрессий на NaN/Inf, fallback-атрибуции, parity price_rows vs breakdown.

**Fix:** Integration tests.

---

## Приоритетный План Действий

### 🔴 P0 — CRITICAL (Немедленно, блокирует релиз)

| ID | Название | Модули | Трудозатраты | Результат |
|----|---------|----|-------|----------|
| P0-1 | Fix [A1]: Разделить primary/secondary ветки | trim.py, ilp_model.py, extract_cuts.py | 1-2 дня | Полный учёт long + trans + waste для смешанного покрытия |
| P0-2 | Fix [A2]: Единый PricingService | price_rows.py, breakdown.py | 1-2 дня | Единая истина для unit_price |
| P0-3 | Fix [A3]: Убрать глобалы, добавить DI | plan_lookup.py, orders.py, items.py | 1-2 дня | Pure функции, безопасность в многопоточности |
| P0-4 | Fix [S2]: math.isfinite валидация | commercial.py, price_rows.py | 0.5 дня | Защита от NaN/Inf в финансах |

**Итого P0:** 4-6.5 дней

### 🟠 P1 — HIGH (Срочно, до конца спринта)

| ID | Название | Модули | Трудозатраты | Результат |
|----|---------|----|-------|----------|
| P1-1 | Fix [A4]: Единый GenerateOfferUseCase | commercial.py | 1 день | Один path для обоих сценариев |
| P1-2 | Fix [A5]: Единая стратегия rounding | price_db.py, price_utils.py, price_rows.py | 0.5 дня | Консистентность ceil/round |
| P1-3 | Fix [A6]: Формализовать контракт secondary | trim.py, optimize_1d_widths.py, geometry.py | 1 день | Чёткий invariant qty/pieces/waste |
| P1-4 | Fix [Q1-Q3]: Centralized numeric guard | validation.py (новый) | 1 день | Валидация всех числовых входов |
| P1-5 | Fix [S1]: Request-scope context | context.py, commercial.py | 0.5 дня | Безопасность в многопоточности |
| P1-6 | Fix [S3]: Strict bounds в DTO | plan_snapshot.py, trim.py | 0.5 дня | Валидация план-структур |

**Итого P1:** 4.5 дней

### 🟡 P2 — MEDIUM (Текущий квартал)

| ID | Название | Модули | Трудозатраты | Результат |
|----|---------|----|-------|----------|
| P2-1 | Fix [S4-S7]: Error handling, logging, sanitization | commercial.py, price_db.py, price_utils.py | 1.5 дней | Безопасность и observability |
| P2-2 | Fix [Q4-Q6]: Decompose commercial.py, DTO контракты | commercial.py | 1.5 дней | Читаемость, тестируемость |
| P2-3 | Fix [A7]: Integration tests (mixed primary+secondary, parity) | test_procurement_trim_cuts.py | 1.5 дней | Регрессионная защита |
| P2-4 | Fix [A8]: Синхронизировать документацию | prise_rules.md | 0.5 дня | Актуальная docs |
| P2-5 | Fix [S8-S9]: SQLite resilience, security tests | core/price_db.py, tests/ | 1 день | Устойчивость к load, безопасность |

**Итого P2:** 6 дней

---

## Затронутые Файлы (Scope)

```
viz_modules/procurement/
├── trim.py               [A1, A6, S3, Q1] ← CRITICAL
├── price_rows.py         [A2, A5, Q1, Q6] ← CRITICAL
├── breakdown.py          [A2, Q1]
├── plan_lookup.py        [A3, S1]
├── orders.py             [A3, A4, S1]
├── items.py              [A3]

core/optimization/
├── ilp_model.py          [A1]
├── context.py            [S1]
├── order_dispatch.py     [Q2, Q5]
├── optimize_1d_widths.py [A6]
├── optimize_2d/
│   ├── extract_cuts.py   [A1]
│   ├── prep_solve.py     [Q5]
│   ├── finalize.py       [Q5]
├── geometry.py           [A6]
├── validation.py         [Q3] (новый файл для centralized guard)

core/
├── price_db.py           [A5, S5, S8]
├── config/
│   └── constants.py      [A3] (review глобалов)

viz_modules/
├── price_utils.py        [A5, S5]

bot/
├── handlers/commercial.py [A4, S2, S4, S7, Q4, Q6]

tests/
├── test_procurement_trim_cuts.py [A7, S9, Q8]

ai_docs/develop/
├── prise_rules.md        [A8]
```

---

## Рекомендации Для Немедленного Действия

1. **Остановить ввод** расчётов смешанного primary+secondary в production до fix [A1].
2. **Добавить сквозной тест** на parity price_rows vs breakdown.
3. **Добавить numeric guard** (math.isfinite) в point of control (commercial handler).
4. **Провести code review** всех изменений в `price_rows.py` и `trim.py` за последние 2 месяца для выявления функционального дрейфа.
5. **Спланировать рефакторинг** в следующем спринте (P0 — 4-6 дней).

---

## Контрольные Вопросы Для Validation

- [ ] Вся логика ценообразования централизована в одном PricingService?
- [ ] Нет глобальных читаний в слое procurement (только DI)?
- [ ] Все числовые входы валидируются (isfinite, bounds)?
- [ ] Есть integration test на primary+secondary смешанное покрытие?
- [ ] Price breakdown совпадает с price_rows?
- [ ] Нет формулной injection в XLSX?
- [ ] Логирование не содержит коммерческих тайн?

---

## Заключение

Система расчёта цены имеет **серьёзные архитектурные проблемы**, которые могут привести к:
- 🔴 **Недоучёту обрезков и остатков** (особенно в смешанных сценариях)
- 🔴 **Несовпадению итоговой цены** между различными вычислениями
- 🔴 **Уязвимостям безопасности** (context leakage, formula injection, NaN propagation)

**Обрезки (остатки) учитываются в unit_price, но НЕПОЛНО в критичных сценариях.**

Требуется немедленный рефакторинг P0-задач перед использованием в production.

---

**Дата отчета:** 2026-06-01 | **Версия:** 1.0 | **Статус:** 🔴 Требует вмешательства
