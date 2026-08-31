# Промпт: ГСМ hardening этап 1 (мультиагент)

Скопировать всё от строки «НАЧАЛО ПРОМПТА» до «КОНЕЦ ПРОМПТА» в новое окно / оркестратор.

---

НАЧАЛО ПРОМПТА

Ты в репозитории «Шишов» (FastAPI + React). Выполни **только этап 1** спеки hardening закрытия месяца ГСМ. Спека и план уже утверждены владельцем. Не переоткрывай продуктовую дискуссию.

## Прочитай сначала (обязательно)

1. `.cursor/skills/project-shishov/SKILL.md`
2. `.cursor/skills/plan-web-context/SKILL.md`
3. `ai_docs/specs/gsm-month-close-hardening.md`
4. `ai_docs/develop/plans/2026-08-28-gsm-month-close-hardening.md`
5. `ai_docs/specs/gsm-month-close-kit.md` — инвариант комплекта (не ломать UX)
6. Существующий гейт UI: `frontend/src/features/gsm/lib/exportGate.ts` (`planKit`, `planBulkGenerate`)
7. `app/services/gsm_overview_service.py` (`_chain_broken`, `_status_of`)
8. `app/services/gsm_report_service.py`, `gsm_export_service.py`, `gsm_generation_service.py`

Скиллы на код: `test-driven-development`, `incremental-implementation`. Перед merge-качеством — `code-review-and-quality`. Оркестрация: `/orchestrate` по задачам T1→T4 плана, **не** изобретать этапы 2–4.

## Проблема

Комплект месяца (`POST /api/v1/gsm/report/usage`) и `POST /gsm/waybills/export` ставят ПЛ в `exported`. Фронт (`planKit`) не кладёт в запрос машины с хвостом июля / разрывом цепи / красными днями. **Бэкенд это не проверяет** (кроме красных на usage). Прямой API или кнопка «Экспорт» в строке может закрыть не тот месяц.

## Сделать (этап 1, 4 задачи по плану)

**T1.** `app/services/gsm_kit_gate.py` + `tests/test_gsm_kit_gate.py`.  
Один helper eligibility машины за `period_from`…`period_to`. Данные — те же поля, что overview (`open_before`, `open_before_month`, `chain_broken` / те же пороги 0.01 л, red = `manual_intervention` в периоде). Не копировать расходящийся SQL.

Правила (комплект = usage + прямой export):

| Условие | Комплект | Generate |
|---------|----------|----------|
| red в периоде | запрещён | разрешён |
| `open_before > 0` и месяц периода ≠ `open_before_month` | запрещён | запрещён |
| месяц периода = `open_before_month` (закрываем хвост) | хвост не блокирует | разрешён |
| `chain_broken` | запрещён | **разрешён** |
| иначе | разрешён | разрешён |

Коды: `gsm_kit_tail` | `gsm_kit_chain` | `gsm_kit_red` (или маппинг в `gsm_kit_gate`).

**T2.** Подключить гейт в `build_usage_zip` **до** soffice и **до** flip `exported`; в `export_zip` — то же (прямой API).  
`vehicle_ids: null` = все активные, затем фильтр.  
Смесь чистых и плохих → 200, zip только чистых, больные **не** `exported`. Список skip в HTTP **не** добавлять (только UI).  
Одна запрошенная плохая / ноль прошедших → 4xx (`gsm_report_no_data` / `gsm_export_empty` / kit-код), статусы не менять.  
Регрессия мая 848 не должна разъехаться. Soffice в тестах мокать как сейчас.

**T3.** `generate` / `generate_bulk`: хвост чужого месяца → 4xx или per-id `ok: false` (bulk не откатывает соседей). `chain_broken` generate **не** стопит.

**T4.** `_status_of`: если `red_days > 0` → `has_red_days` **раньше**, чем `needs_generation`.  
`FleetOverviewView.handleExportKit`: для текущего периода не звать `runKit` в обход `planKit`. Прыжок на месяц хвоста оставить; 4xx сервера → ошибка на экране, не «скачан zip».

TDD: красные тесты → код. После каждой задачи — команды verify из плана.

## Verify

```bash
.venv/bin/python -m pytest tests/test_gsm_kit_gate.py tests/test_gsm_usage_report.py tests/test_gsm_export.py tests/test_gsm_generation_api.py tests/test_gsm_overview_api.py tests/test_gsm_generator.py -q

cd frontend && npx vitest run src/features/gsm/components/FleetOverviewView.test.tsx src/features/gsm/lib/exportGate.test.ts
```

## Мультиагент (DAG)

- T1 первым (один worker).
- T2 и T3 после T1, можно параллельно (разные сервисы).
- T4 после T2 (контракт 4xx).
- Не параллелить запись в одни и те же файлы.
- test-runner после каждого task; reviewer в конце этапа 1.
- documenter: короткий отчёт в `ai_docs/develop/reports/2026-08-28-gsm-month-close-hardening-stage1.md` (что сделано, какие тесты).

## NEVER

- Этапы 2–4: `norm_l_per_100`, banker's round, импорт xls, пакетный soffice, фоновый job, Celery.
- `core/gsm/generator.py`, шаблоны бланков, схема БД, live `plita.db`.
- Коммит и push без явной просьбы пользователя в **этом** чате.
- Рефакторинг god-сервисов «заодно», полный аудит 2026-08-28 (КП/IDOR).
- Файл `исключения.txt` в zip; заголовок skip в API.

## ALWAYS

- Слои: роутер тонкий, логика в services.
- Минимальный diff.
- Существующие паттерны ошибок GSM (`GsmReportError` / `GsmExportError` / `GsmGenerationError` + HTTP mapper).
- Русский UI, стабильные machine codes.

## Готово когда

Чеклист Checkpoint этапа 1 в плане выполнен (тесты; live UI не гонять, пока пользователь не скажет). Кратко отчитайся: файлы, команды, что не делал.

КОНЕЦ ПРОМПТА
