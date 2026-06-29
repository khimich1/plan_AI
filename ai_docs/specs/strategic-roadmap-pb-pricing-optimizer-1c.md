# Spec: Стратегическая дорожная карта — цены ПБ, целостность оптимизатора, 1С, мульти-завод

> **Тип:** product vision + phased roadmap  
> **Фаза SDD:** SPECIFY  
> **Дата:** 2026-06-18  
> **Статус:** черновик на ревью  
> **Связанные документы:** [`project-baseline.md`](./project-baseline.md)

---

## ASSUMPTIONS I'M MAKING

1. **Сейчас в фокусе только изделия ПБ** (плиты); другие ЖБИ — будущий этап без детальной спеки.
2. **Эталон цены ПБ** — прайс в `pb.db` (таблица `prices`) + правила резов из `core/config/constants.py`; сверка с Excel-источником в `банк знаний/`.
3. **«Верная цена»** = совпадение с утверждённым прайсом завода **и** корректный учёт продольных/поперечных резов, нагрузки 12,5п, округления длины в дм (ceil).
4. **«Оптимизатор ничего не теряет»** = для каждого заказа `verify_coverage.ok == True` и нет `lost_plates` при `plan_commit` без явного rescue/unmapped.
5. **Интеграция с 1С** — внешняя зависимость; спека API/обмена **блокирована** до ответа 1С-специалистов.
6. **Мульти-завод** — целевое состояние через 12+ месяцев; сейчас single-tenant (один завод, один `plita.db` + `pb.db`).
7. **Не ломаем текущий production** — изменения через тесты, golden fixtures, feature flags в settings.

→ Поправь допущения сейчас или подтверди.

---

## Objective

Обеспечить **доверие к системе** на заводе ЖБИ:

1. Менеджер и производство видят **правильную цену ПБ** в КП, архиве и производственной смете.
2. Оптимизатор раскладки **не теряет плиты** на всём пути: заказ → ILP → план → коммит в БД → завершение смены.
3. Подготовить **контур обмена с 1С** (номенклатура, заказы, отгрузки) — после получения требований от 1Сников.
4. Заложить архитектуру для **новых типов изделий** (не только ПБ) и **других заводов**.

### Пользователи и ценность

| Роль | Боль сейчас | Ценность после |
|------|-------------|----------------|
| Менеджер | Сомнения в сумме КП | КП = прайс завода ± скидка, прозрачная разбивка |
| Производство | Риск «потерянных» плит в плане | План = заказ, аудит при расхождении |
| Руководство | Ручной ввод в 1С | Автообмен, меньше ошибок |
| IT / 1С | Нет контракта | Согласованный API/формат обмена |

---

## Текущее состояние (as-is)

### Расчёт цены ПБ — несколько путей (риск расхождений)

| Путь | Модуль | Когда используется | Особенности |
|------|--------|-------------------|-------------|
| Базовая цена из БД | `core/price_db.py` → `get_price()` | Fallback в КП | `length_m_to_price_length_dm` — **ceil** до дм |
| КП (PDF/XLSX) | `core/commercial_offer.py`, `commercial_offer_xlsx.py` | Итоги КП | Приоритет `unit_price` в позиции; иначе `get_plate_price()` |
| Производственная смета | `viz_modules/procurement/price_rows.py`, `price_utils.py` | После оптимизации | Учёт резов, армирования по дорожкам; fallback ±1 дм |
| Себестоимость | `factory_cost/cost_engine.py` | Отдельный контур | Таблица `factory_plate_costs`, не путать с прайсом продажи |
| Константы резов | `core/config/constants.py` | Надбавки за резы | `LONG_CUT_PRICE_PER_M=460`, `TRANSVERSE_CUT_PRICE=1200` |

**Известные риски ценообразования:**

- Дублирование `get_plate_price` в `commercial_offer.py` и логики в `price_db.py` / `price_utils.py`.
- Fallback `площадь × 4000 ₽` при отсутствии цены в БД — **скрытая ошибка**, не должна попадать в production КП без warning.
- Разные правила lookup: точное совпадение vs `ABS(length_dm-?)<=1` в `price_utils.py`.
- Нагрузка **12,5п** — особая логика (`floor` → колонка 12) — должна быть единообразной везде.
- PDF и XLSX дублируют `calculate_total_cost` — уже есть тесты на логистику (`test_commercial_logistics_cost.py`), но нет единого **golden suite** по прайсу.

### Оптимизатор — механизмы контроля целостности

| Механизм | Модуль | Поведение сейчас |
|----------|--------|------------------|
| `verify_coverage()` | `core/optimization/coverage_verify.py` | Сверка demand vs primary+secondary; `ok: bool` |
| `PlateAudit` | `core/plate_audit.py` | Checkpoints между стадиями; `has_losses()` |
| Finalize audit | `core/optimization/optimize_2d/finalize.py` | Лог ERROR при потерях post-correction |
| `plan_commit` | `core/plan_commit.py` | `lost_plates` → warning; `optimizer_unmapped` → **PlanCommitError** |
| Result contract | `core/optimization/result_contract.py` | `_opt_status`: ok / error / partial |
| Тесты | `tests/test_optimization_baseline.py` | Контракт: `verify_coverage.ok == True` на baseline-заказах |

**Известные ограничения:**

- `verify_coverage` в finalize — «наблюдатель»; комментарий в коде: «в этап 2 модель должна выдавать ok=True всегда».
- `lost_plates` при commit — **warning**, не всегда блокирует (rescue-плиты отдельно).
- Нет единого **регрессионного каталога** реальных заказов с завода с автоматической проверкой coverage + цены.

### 1С

- В кодовой базе **нет** упоминаний 1С, OData, CommerceML, обмена XML/JSON.
- Интеграция = greenfield.

### Мульти-изделия / мульти-завод

- Домен завязан на парсер ПБ (`plate_line_parser`, марка «ПБ 78-12-8п»).
- `Settings` — один набор путей к БД и прайсам (`plita_db_path`, `pb_db_path`, `price_xlsx_path`).
- Нет tenant_id / factory_id в схеме БД.

---

## Tech Stack (без изменений на Phase 1–2)

См. [`project-baseline.md`](./project-baseline.md). Для Phase 3–5 возможны дополнения (см. ниже) — **только после согласования**.

---

## Roadmap (фазы)

```
Phase 1 ──► Phase 2 ──► Phase 3 ──► Phase 4 ──► Phase 5
 Цены ПБ    Оптимизатор   1С         Новые ЖБИ    Мульти-завод
 (4–6 нед)  (4–6 нед)    (blocked)   (TBD)        (TBD)
```

---

## Phase 1: Верификация и унификация расчёта цены ПБ

### Objective

Менеджер получает в КП **ту же цену**, что в утверждённом прайсе завода, с корректными надбавками за резы и логистику. Любое отклонение — явное предупреждение, не silent fallback.

### Scope

**In scope:**
- Единый сервис ценообразования продажи ПБ в `core/` (без импорта `app/`).
- Golden tests: фикстуры «плита → ожидаемая цена» из `pb.db` / Excel.
- Сверка PDF vs XLSX vs API preview breakdown.
- Документирование правил: ceil длины, 12,5п, резы, НДС 22%, скидка, логистика.
- UI: индикация позиций с fallback-ценой (если нет в прайсе).

**Out of scope:**
- Себестоимость (`factory_cost/`) как источник цены продажи.
- Интеграция с 1С.

### Предлагаемая архитектура

```
core/pricing/
  pb_sales_price.py      # единая точка: get_unit_price(plate_spec) → PriceResult
  price_rules.py         # ceil dm, load_code 12.5, cut surcharges
  types.py               # PlatePriceSpec, PriceResult(warnings)
```

Потребители (рефакторинг поэтапно):
- `commercial_offer.py` / `commercial_offer_xlsx.py` → вызывают `core/pricing`
- `viz_modules/procurement/price_rows.py` → тот же API + production-надбавки
- `app/services/commercial_calculation_service.py` → делегирует в core

### Success Criteria (измеримые)

| # | Критерий | Проверка |
|---|----------|----------|
| 1 | 100% позиций из golden-набора (≥50 кейсов) совпадают с эталоном ±0.01 ₽ | `pytest tests/test_pb_pricing_golden.py` |
| 2 | `calculate_total_cost` PDF == XLSX для одного `order_data` | существующие + новые тесты |
| 3 | Нет silent fallback 4000×площадь в production без `warning` в ответе API | integration test |
| 4 | `length_m_to_price_length_dm(2.73)==28` и lookup по load 6/8/10/12 | `test_price_length_dm.py` + расширение |
| 5 | Breakdown в wizard показывает компоненты цены (база + резы) | manual + snapshot test |

### Commands (verify)

```powershell
pytest tests/test_price_length_dm.py tests/test_commercial_logistics_cost.py -q
pytest tests/test_pb_pricing_golden.py -q          # после реализации
pytest tests/test_commercial_web_flow.py -q -k breakdown
```

### Risks

| Риск | Митигация |
|------|-----------|
| Прайс Excel ≠ pb.db | Скрипт сверки + CI check при обновлении прайса |
| Регрессия в procurement | Общий `core/pricing`, не копировать логику |

### Open Questions (Phase 1)

1. Кто на заводе — **источник истины** по цене: Excel, 1С или `pb.db`?
2. Нужна ли сверка с **историческими КП** (архив XLSX в БД)?
3. Fallback при отсутствии цены: **запрет сохранения** или ручное подтверждение менеджером?

---

## Phase 2: Гарантия целостности оптимизатора («ничего не теряет»)

### Objective

Для любого заказа, прошедшего оптимизацию и попавшего в план: **каждая заказанная плита** учтена в раскладке и привязана к заказу/КП. Потери = блокирующая ошибка, не warning.

### Scope

**In scope:**
- Жёсткий gate: `verify_coverage.ok == False` → `_opt_status: error`, план не строится.
- `plan_commit`: `lost_plates` и `optimizer_unmapped` → **всегда** `PlanCommitError` (кроме явно документированного rescue).
- Расширение `PlateAudit`: обязательные checkpoints в orchestrator.
- Регрессионный каталог заказов (`tests/fixtures/orders/`) с auto-coverage.
- Мониторинг: метрика `optimizer_coverage_failures` в логах.

**Out of scope:**
- Изменение математики ILP (только если тесты выявят баг).
- Оптимизация скорости solver.

### Success Criteria

| # | Критерий | Проверка |
|---|----------|----------|
| 1 | Baseline-набор (≥20 заказов из production) — все `verify_coverage.ok` | `test_optimization_baseline.py` расширен |
| 2 | API `POST /production/plans/build` возвращает 422 при coverage fail | integration test |
| 3 | `plan_commit` не помечает плиты в БД при любом unmapped optimizer plate | `test_plan_commit.py` |
| 4 | Wizard КП: после calculate в metadata есть `_coverage_summary.ok` | commercial flow test |
| 5 | Документ «что считается потерей» согласован с производством | sign-off |

### Текущие точки усиления (файлы)

- `core/optimization/optimize_2d/finalize.py` — поднять coverage fail до error
- `core/plan_commit.py` — ужесточить политику `lost_plates`
- `app/services/production_planning_service.py` — проброс `_opt_status` в API
- `app/services/optimization_service.py` — единая валидация результата

### Commands (verify)

```powershell
pytest tests/test_optimization_baseline.py tests/test_plan_commit.py -q
pytest tests/test_optimization_semantics_and_tracks.py -q
pytest tests/test_layout_identity_integrity.py -q
```

### Open Questions (Phase 2)

1. Допустим ли статус `_opt_status: partial` в production или только `ok`/`error`?
2. Rescue-треки — отдельный бизнес-процесс или часть основного плана?
3. Нужен ли **отчёт аудита** для мастера смены в UI?

---

## Phase 3: Интеграция с 1С (заблокировано)

### Objective

Двусторонний обмен: номенклатура, контрагенты, заказы/КП, статусы производства, отгрузки — без двойного ввода.

### Blocker

**Ждём от 1С-специалистов:**

| # | Вопрос | Зачем |
|---|--------|-------|
| 1 | Версия 1С (УТ, ERP, КА, своя доработка)? | Выбор протокола |
| 2 | Способ обмена: REST/OData, HTTP-сервис, файлы (XML/JSON), Kafka/Rabbit? | Архитектура |
| 3 | Мастер-данные: что ведётся в 1С vs в «Шишов»? | Source of truth |
| 4 | Формат номенклатуры ПБ (код, GUID, характеристики)? | Маппинг |
| 5 | События: КП создано, план утверждён, смена закрыта — что уходит в 1С? | Event model |
| 6 | Требования к идемпотентности, retry, очереди ошибок? | Надёжность |

### Предварительная архитектура (до ответа 1С)

```
app/integrations/onec/
  client.py           # HTTP/файловый клиент
  schemas/            # DTO обмена (Pydantic)
  mappers/            # KP ↔ 1С документ
  sync_service.py     # оркестрация
  outbox_table        # в plita.db или отдельная очередь — TBD
```

**Принципы:**
- `core/` не знает про 1С — только доменные события (`KpSaved`, `PlanActivated`, `DayCompleted`).
- Integration layer в `app/integrations/` подписывается на события сервисов.
- Feature flag: `ONEC_SYNC_ENABLED=false` по умолчанию.

### Placeholder Success Criteria (уточнить после ответа 1С)

- КП, сохранённое в «Шишов», создаёт документ в 1С за < N минут.
- Номенклатура из 1С обновляет `pb.db` / справочники без ручного импорта Excel.
- Повторная отправка не дублирует документ (idempotency key).

### Out of Scope до разблокировки

- Любая реализация кода интеграции.
- Изменение схемы `plita.db` под 1С GUID — только после согласования маппинга.

---

## Phase 4: Новые типы изделий (не только ПБ)

### Objective

Добавлять ЖБИ (балки, фундаментные блоки, лестничные марши и т.д.) без переписывания всей системы.

### Scope (high-level, детализация после Phase 1–2)

**Архитектурные направления:**

1. **Product type registry** в `core/domain/products/`:
   - `ProductKind.PB`, `ProductKind.BEAM`, …
   - Парсер, валидатор, pricing strategy, optimizer strategy — по типу.

2. **Расширяемый парсер** — сейчас `plate_line_parser`; новые — отдельные модули + единый `OrderLine` DTO.

3. **Оптимизация** — не все изделия требуют 2D ILP; routing: `optimizer_for(product_kind)`.

4. **Схема БД** — `kp_plates.plate_name` + `product_kind` column (migration в `kp_db_schema.py`).

5. **UI wizard** — шаг выбора типа изделия или автоопределение из текста.

### Blocker

Недостаточно информации: номенклатура, правила ценообразования, раскладка/планирование для не-ПБ.

### Open Questions

1. Какие **3 следующих типа** изделий после ПБ по приоритету?
2. Общий прайс или отдельные БД/таблицы?
3. Нужна ли оптимизация раскроя для не-ПБ или только КП/учёт?

---

## Phase 5: Внедрение на другие заводы (multi-factory)

### Objective

Один codebase — несколько заводов с изолированными данными и настраиваемыми прайсами/правилами.

### Предварительная модель

| Аспект | Single-tenant (сейчас) | Multi-factory (цель) |
|--------|------------------------|----------------------|
| БД | `plita.db`, `pb.db` | `data/{factory_id}/plita.db` или PostgreSQL schemas |
| Прайс | один Excel | per-factory config |
| Пользователи | общий allowlist | `factory_id` в users / bot allowlist |
| Настройки резов | constants.py | `FactoryProfile` в settings/БД |
| Deploy | один docker compose | compose per factory или k8s + tenant |

### Success Criteria (draft)

- Новый завод подключается через конфиг + seed БД без форка кода.
- Данные заводов не пересекаются (IDOR-тесты).
- Прайс и календарь — per factory.

### Out of Scope до Phase 4

- SaaS billing, self-service onboarding.

---

## Commands (общие)

```powershell
Set-Location "c:\Users\Роман\Desktop\Шишов"
.\.venv\Scripts\Activate.ps1

# Регрессия ценообразования + оптимизации (текущий набор)
pytest tests/test_price_length_dm.py tests/test_commercial_logistics_cost.py -q
pytest tests/test_optimization_baseline.py tests/test_plan_commit.py -q

# Полный прогон перед релизом фазы
pytest tests/ -q

# Frontend (если меняется breakdown UI)
Set-Location frontend
npm run build
npm run test
```

---

## Project Structure (целевые добавления)

```
core/
  pricing/                    # Phase 1 — NEW
    pb_sales_price.py
    price_rules.py
  domain/
    products/                 # Phase 4 — NEW
      registry.py
      pb.py
tests/
  fixtures/
    pricing_golden.json       # Phase 1
    orders_regression/        # Phase 2
app/
  integrations/
    onec/                     # Phase 3 — после разблокировки
ai_docs/specs/
  onec-integration.md         # дочерняя spec после ответа 1С
  product-pb-pricing.md       # детальная spec Phase 1
  optimizer-integrity.md      # детальная spec Phase 2
```

---

## Code Style (эталон для Phase 1 — единый pricing)

```python
# core/pricing/pb_sales_price.py
from core.pricing.types import PlatePriceSpec, PriceResult

def resolve_pb_unit_price(spec: PlatePriceSpec, *, db_path: str) -> PriceResult:
  """
  Единая цена продажи ПБ с НДС. Без side effects.
  warnings: ["price_not_in_db_used_fallback"] — если применён fallback.
  """
  ...
```

```python
# Phase 2 — gate в optimization
from core.optimization.coverage_verify import verify_coverage
from core.optimization.result_contract import opt_error, ERROR_COVERAGE_INCOMPLETE

cov = verify_coverage(demand_2d, primary_cuts, secondary_cuts)
if not cov["ok"]:
    return opt_error(ERROR_COVERAGE_INCOMPLETE, detail=cov)
```

---

## Testing Strategy

| Фаза | Тип теста | Файлы / подход |
|------|-----------|----------------|
| 1 | Golden pricing | `test_pb_pricing_golden.py` + JSON fixtures |
| 1 | Parity PDF/XLSX/API | `test_commercial_logistics_cost.py`, web flow |
| 2 | Coverage regression | `test_optimization_baseline.py` + `fixtures/orders/` |
| 2 | Commit integrity | `test_plan_commit.py`, `test_production_planning_service.py` |
| 3 | Contract tests | mock 1С server, после spec onec-integration |
| 4–5 | Product/factory isolation | boundary tests |

**Покрытие:** не гнаться за %; критично — golden paths для денег и плит.

---

## Boundaries

### Always

- Phase 1–2 **не ждут** 1С.
- Любое изменение цены или coverage — с тестом.
- `core/` не импортирует `app/`.
- Fallback-цены — только с явным warning в API.

### Ask first

- Ужесточение `plan_commit` (может сломать существующие планы с rescue).
- Новые таблицы в `plita.db` / `pb.db`.
- Зависимости для 1С (httpx, celery, …).
- Изменение `_opt_status` контракта (frontend/bot).

### Never

- Silent потеря плит «для удобства».
- Цена из себестоимости в КП без явного бизнес-решения.
- Интеграция 1С без письменного контракта обмена.
- Hardcode данных одного завода в `core/` (готовить к Phase 5).

---

## Success Criteria (дорожная карта «готово»)

### Milestone A — Доверие к цене (Phase 1 complete)

- [ ] Golden suite ≥50 кейсов зелёный
- [ ] Нет undocumented fallback в production path
- [ ] Производство и менеджеры подписали sample КП vs Excel

### Milestone B — Доверие к оптимизатору (Phase 2 complete)

- [ ] Regression orders — 100% coverage ok
- [ ] API блокирует план при потерях
- [ ] Документирован audit trail

### Milestone C — 1С (Phase 3)

- [ ] Spec `onec-integration.md` утверждена 1С + IT
- [ ] Pilot обмен на staging
- [ ] Production sync с мониторингом ошибок

### Milestone D — Расширение (Phase 4–5)

- [ ] Второй тип изделия в production
- [ ] Второй завод на том же codebase

---

## Risks & Mitigations (сквозные)

| Риск | Вероятность | Митигация |
|------|-------------|-----------|
| Прайс и оптимизатор — разные команды правды | Высокая | Phase 1–2 параллельно, общий regression CI |
| 1С затягивается | Высокая | Не блокировать A/B; ручной экспорт CSV как временный мост |
| Scope creep «новые изделия» | Средняя | Phase 4 только после Milestone A+B |
| Multi-factory преждевременно | Средняя | FactoryProfile в config, не микросервисы |

---

## Open Questions (для владельца продукта)

1. **Приоритет:** Phase 1 (цены) и Phase 2 (оптимизатор) — параллельно или сначала цены?
2. Есть ли **эталонные КП** (Excel/PDF) для сверки цен от бухгалтерии?
3. Есть ли **реальные проблемные заказы**, где оптимизатор «терял» плиты?
4. Когда ожидается ответ от **1Сников** (хотя бы по протоколу)?
5. Какой **следующий тип изделия** после ПБ и есть ли по нему прайс?
6. **Второй завод** — тот же владелец/процессы или другая модель?

---

## Out of Scope (весь документ)

- Редизайн UI wizard
- Миграция SQLite → PostgreSQL (кроме Phase 5 если потребуется)
- Мобильное приложение
- Закупки PDF из `закупки/` (отдельная фича)

---

## Следующие шаги SDD

1. **Ревью этой spec** — подтвердить ASSUMPTIONS и приоритет фаз.
2. Детализировать **дочерние spec** (по одной на фазу):
   - `product-pb-pricing.md` (Phase 1)
   - `optimizer-integrity.md` (Phase 2)
3. После ответа 1С — `onec-integration.md` (Phase 3).
4. Фаза **PLAN** → **TASKS** → **IMPLEMENT** — только для выбранной Phase 1 или 2.
