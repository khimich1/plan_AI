# Brief для оркестратора: GSM lookahead-генератор с географией

> Вставь это сообщение **целиком** в новое окно агента (Agent mode).
> Идея, спека и план уже утверждены человеком (2026-08-15). **Не** начинать SDD/interview/idea-refine заново.

---

```
/orchestrate execute без подтверждения плана

Задача: реализовать lookahead-генератор ГСМ с географией (round-trip + направление к следующей заправке + частичная генерация), устранив массовые падения 422 на плотных участках заправок.

## Источники правды (читать первым делом)

1. Спека: ai_docs/specs/gsm-geo-lookahead-generator.md
2. План задач: ai_docs/develop/plans/2026-08-15-gsm-geo-lookahead.md
3. Идея (контекст, не scope): ai_docs/ideas/gsm-geo-lookahead-generator.md
4. Базовая спека модуля: ai_docs/specs/gsm-module-putevye-listy.md
5. Скиллы: project-shishov, orchestration, incremental-implementation, test-driven-development
6. Backend: pytest tests/ из корня с venv (команда: venv/bin/pytest); frontend: npm test / tsc из frontend/

План УЖЕ содержит Task 1–13 с acceptance, files, dependencies, gates. Не вызывать planner заново. Не переписывать спеку/план, кроме статусов задач (pending → completed) и короткого implementation report в конце.

Workspace: создай .cursor/workspace/active/orch-2026-08-15-gsm-geo-lookahead/ с progress.json, tasks.json, links.json, указывающими на план выше. Resume с первой невыполненной задачи.

## Принятые решения — не переоткрывать

- R1: Ночёвки — ФАЗА 2, НЕ в этом срезе. Проверено на истории: при max_daily_km=700 жёстких случаев, требующих ночёвки для баланса, нет ни у одной машины. MVP = геокодинг + привязка станций + lookahead + round-trip (2 плеча) + частичная генерация. НЕ реализовывать overnight_trip / return_leg / разбиение на 2 дня.
- R2: max_daily_km = 700 (дефолт в gsm_setting, переопределяется). Round-trip: день = 2 плеча, дневной km = 2×km плеча, сжигается burn(2×km).
- R3: Приоритет — сходимость бака > география. Сначала отбор маршрутов по km (баланс), потом мягкая сортировка по направлению к следующей АЗС (3 уровня приоритета, angle_diff ≤ 90°).
- R4: Минимальный достаточный km — не жечь лишнего.
- R5: Частичная генерация вместо 422. Нерешаемый якорь → draft manual_intervention, период собирается целиком. 422 ТОЛЬКО для конфигурационных ошибок (нет машины/маршрутов/водителя), НЕ для баланса. Контракт POST /gsm/waybills/generate меняется: 200 с problematic_days.
- R6: Базы машин — две: Кострома (Кузнецкая 18Б) и Ярославль (Домостроителей). «Вернуться домой» = в одну из них.
- R7: Геокодинг — только из scripts/ (через существующий кэш ГСМ/geo_cache/addresses.json, Nominatim, rate-limit 1 req/s). НЕ геокодить из backend в runtime.
- R8: Дата-миграции не меняют схему: только UPDATE NULL/пустых полей gsm_station.lat/lon и gsm_route.typical_station_ids. Существующие данные не затирать.
- R9: core/gsm/geo.py и generator.py — чистые функции без I/O и без импортов app.*. Солвер детерминирован (seed при tie-break).
- R10: Коммиты/git push/gh pr — только если пользователь явно попросит; по умолчанию НЕ коммитить.

## Порядок (DAG)

Phase 0 (Data):
T1 geocode stations (∥ T2) → T2 geo.py+тесты → T3 link_route_stations (после T1+T2)
Gate: все станции с координатами; typical_station_ids заполнены; Palisade имеет маршруты через свои АЗС.

Phase 1 (Core):
T4 round-trip (2 плеча) → T5 lookahead → T6 география/направление (после T2+T3+T5) → T7 частичная генерация (после T5+T6)
Gate: pytest tests/test_gsm_geo.py tests/test_gsm_generator.py -q зелёный; май Palisade собирается без 422 в unit-симуляции.

Phase 2 (Service/API):
T8 max_daily_km setting → T9 API problematic_days (после T7+T8)
Gate: POST /gsm/waybills/generate на май Palisade → 200 с днями + problematic_days.

Phase 3 (Frontend):
T10 warning-коды/бейджи → T11 GsmPeriodView частичная генерация (после T9+T10)

Phase 4 (Acceptance):
T12 acceptance май Palisade → T13 docs

После каждой задачи — verification из плана. Красные тесты писать ДО реализации (TDD).

## Never

- Реализовывать ночёвки / дальний рейс на 2 дня (overnight_trip, return_leg) — это фаза 2
- Составные дни из 3+ плеч (полный солвер)
- Геокодить из backend в runtime
- Менять схему SQLite (новых таблиц нет; только UPDATE NULL-полей)
- Затирать существующие координаты/данные при миграциях
- Ломать v1-поведение на простых периодах (обязательны регрессионные тесты)
- Коммитить без явной просьбы
- Трогать bot_archived
- Удалять падающие тесты без замены

## Команды верификации (минимум)

Backend core:
venv/bin/pytest tests/test_gsm_geo.py tests/test_gsm_generator.py -q

Service/API:
venv/bin/pytest tests/test_gsm_generation_service.py tests/test_gsm_api_integration.py -q

Frontend:
cd frontend && npm test -- --run src/features/gsm/

Регрессия (обязательно):
venv/bin/pytest tests/ -q

Data-миграции:
venv/bin/python scripts/geocode_gsm_stations.py --db plita.db
venv/bin/python scripts/link_route_stations.py --db plita.db

## Definition of done

Критерии спеки SC-G1…SC-G6. Palisade май 2026 (04.05–31.05, старт 28 л / 128327 км) собирается без 422; ≤3 дней manual_intervention; маршруты правдоподобны. Все существующие тесты зелёные (регрессия v1). Краткий report: ai_docs/develop/reports/2026-08-15-gsm-geo-lookahead-implementation.md

Если задача упирается в противоречие с R1–R10 — остановись и напиши человеку. Не угадывай продукт заново.
```

---

## Как запустить в другом окне

1. Новое окно Cursor Chat, **Agent mode** (не Ask).
2. Прикрепи `@ai_docs/develop/plans/2026-08-15-gsm-geo-lookahead-brief.md`, `@ai_docs/develop/plans/2026-08-15-gsm-geo-lookahead.md` и `@ai_docs/specs/gsm-geo-lookahead-generator.md`.
3. Вставь блок между тройными кавычками выше (от `/orchestrate execute` до конца).
4. Не отвечай «ок» в том окне повторно на решения R1–R10 — они уже закрыты.

Это окно можно не держать открытым для реализации. Сюда имеет смысл вернуться, только если оркестратор остановится на противоречии или красном gate.

## Примечание про сеть

Task 1 (геокодинг) требует доступа к nominatim. Если sandbox блокирует сеть, оркестратор должен запросить `required_permissions=["network"]` для команды `geocode_gsm_stations.py`. Остальные задачи сети не требуют.
