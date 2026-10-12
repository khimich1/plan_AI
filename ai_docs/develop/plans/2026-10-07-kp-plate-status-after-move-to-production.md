# Implementation Plan: КП — «в производстве» у плиты только после кнопки

**Спека**: [../../specs/kp-plate-status-after-move-to-production.md](../../specs/kp-plate-status-after-move-to-production.md)  
**Идея**: [../../ideas/kp-plate-status-after-move-to-production.md](../../ideas/kp-plate-status-after-move-to-production.md)  
**Дата**: 2026-10-07  
**Статус**: PLAN ✅ · IMPLEMENT ✅

## Overview

Коммерческая вставка плиты пишет пустой статус. Кнопка «В производство» в той же транзакции ставит «в производстве» пустым строкам без плана. Уже сохранённые плиты архива и согласования без плана один раз лишаются ложной подписи. Запрет удаления, конструктор и обмен с 1С не меняются.

## Architecture Decisions

- Пустой статус — `NULL` в `INSERT`, не новое слово и не перестройка колонки. `DEFAULT 'в производстве'` остаётся для вставок цеха и тестов, которые колонку не указывают.
- Запрет в `_delete_line_row` не трогаем. Пустой статус в него не входит, поэтому замена марки в архиве начинает проходить.
- Подпись ставится в `commit_move_to_production` на том же соединении, до `commit`. Отказ гейта выше по стеку эту функцию не вызывает.
- Чистка — идемпотентный `UPDATE` в `_init_schema_impl`, после блока, который выполняется только когда колонки `status` ещё нет. Старый `NULL → в производстве` не расширять и на рабочую базу не гонять.
- Тест `test_update_rejects_delete_of_plate_in_production` больше не ждёт подпись от `save_kp_to_db`. Статус «в производстве» он выставляет сам, затем проверяет отказ. Защиту не ослабляем.

## Dependency graph

```
Пустой статус при вставке
    ├── замена марки в архиве проходит
    └── кнопка ставит «в производстве»
            └── чистка старых архивных строк
                    └── отбор цеха: архив пуст, после кнопки плита видна
```

## Risks

| Риск | Что делаем |
|---|---|
| Тесты ждут «в производстве» сразу после сохранения КП | Ищем такие проверки и переводим их на явный `UPDATE`, не на умолчание вставки |
| Чистка заденет КП «в работе» | `UPDATE` только через `kp_meta.status IN ('в архиве', 'на согласовании')` и пустой `plan_id` |
| Повтор старой миграции затрёт `NULL` | Её условие `status not in columns` не меняем |
| Кнопка запишет статус, а перевод КП откатится | Оба шага на одном `conn` до единого `commit` |

## Task 1: Пустой статус при сохранении КП

**Description:** Вставка плиты из конструктора пишет `status = NULL`. Замена `line_id` такой строки сохраняется. Строка, которой статус «в производстве» выставлен отдельно, по-прежнему не удаляется.

**Acceptance criteria:**

- [ ] `save_kp_to_db` создаёт плиту с `status IS NULL` и пустым `plan_id`
- [ ] Смена `line_id` этой плиты в `update_kp_from_order_data` коммитит новую марку и соседнюю новую строку
- [ ] После явной установки «в производстве» или «в плане» удаление строки по-прежнему даёт `ValueError`, транзакция откатывается

**Verification:**

- [ ] `pytest tests/test_kp_persistence_mixed.py -q`

**Dependencies:** нет

**Files likely touched:**

- `core/kp_persistence_service.py`
- `tests/test_kp_persistence_mixed.py`

**Estimated scope:** S

## Task 2: Кнопка ставит «в производстве»

**Description:** `commit_move_to_production` до `commit` помечает пустые строки этого КП без `plan_id`. Условия входа в `move_to_production` не меняются.

**Acceptance criteria:**

- [ ] Успешный перевод: КП «в работе», у плит с пустым статусом и пустым `plan_id` стоит «в производстве»
- [ ] Строка «в плане» или с `plan_id` этой командой не переписывается
- [ ] Исключение до `commit` не оставляет КП «в работе» с уже подписанными плитами

**Verification:**

- [ ] `pytest tests/test_move_to_production_atomicity.py -q`

**Dependencies:** Task 1

**Files likely touched:**

- `core/kp/offers_write.py`
- `tests/test_move_to_production_atomicity.py`

**Estimated scope:** S

## Task 3: Чистка ложной подписи и отбор цеха

**Description:** При старте схемы обнулить «в производстве» только у плит без плана, чей КП «в архиве» или «на согласовании». Проверить, что цех не видит архивную плиту с пустым статусом и видит её после перевода КП в работу.

**Acceptance criteria:**

- [ ] Повторный запуск чистки не меняет уже пустые строки и не трогает КП «в работе»
- [ ] Строка «в плане» и строка с `plan_id` остаются
- [ ] `get_all_plates_in_production` не возвращает плиту архивного КП с `status IS NULL`
- [ ] После `commit_move_to_production` эта плита в списке есть

**Verification:**

- [ ] `pytest tests/test_kp_persistence_mixed.py tests/test_move_to_production_atomicity.py tests/test_kp_candidates_scope.py -q -k "plate_status or in_production or move_to_production"`
- [ ] Если фильтр `-k` ничего не поймает, прогнать три файла целиком

**Dependencies:** Task 1, Task 2

**Files likely touched:**

- `core/kp_db_schema.py`
- `tests/test_kp_persistence_mixed.py` или соседний тест схемы
- при необходимости точечная правка теста отбора цеха

**Estimated scope:** S

## Checkpoint

- [ ] Три задачи зелёные
- [ ] `pytest tests/test_production_planning_service.py tests/test_kp_candidates_scope.py -q` без регрессии отбора
- [ ] Фронтенд, `sgp_service`, `insert_kp_plate_remainder_row` не менялись
- [ ] Ревью плана с человеком до кода — выполнено, если этот файл принят

## Not in this plan

- Смена `DEFAULT` колонки и перестройка `kp_plates`
- Новая подпись статуса, кнопки, тексты карточки
- Автозапись `delta_` в 1С
- Правка состава КП «в работе»
