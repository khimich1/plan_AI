# Отчёт: Recover stuck plan plates — CLI guard and tests

**Дата:** 2026-05-07  
**Статус:** завершено

## Кратко

Добавлен модуль pytest `tests/test_recover_stuck_plan_plates_cli.py`: изолированные subprocess-тесты для `scripts/recover_stuck_plan_plates.py` с временной SQLite и `kp_db.init_schema`.

## Реализованное покрытие

1. **`--apply` без `--plan-id`** — код выхода **2**, срабатывает защита **argparse** до логики восстановления; в выводе фигурирует `--plan-id`.
2. **`--help`** — код выхода **0**, в stdout есть `usage:`.
3. **Dry-run, изолированная БД** — после `kp_db.init_schema` в temp-файле, без зависших планов; ожидается успешный выход и текст про отсутствие зависших плит.
4. **Dry-run с явным `plan_id`** — застрявших строк 0, в выводе dry-run / что изменения не применялись.
5. **`--apply` + `--plan-id`** — в temp DB вставлена тестовая строка `kp_plates` (`в плане`, `plan_id` задан); после CLI: статус **`в производстве`**, **`plan_id` = NULL**; в выводе — итог возврата в производство.

## Windows / кодировка

Чтобы дочерний процесс при выводе из `kp_db.init_schema` (эмодзи/русский) не ломал декодирование на **cp1251**, в тестах:

- в окружение subprocess добавляются **`PYTHONUTF8=1`** и **`PYTHONIOENCODING=utf-8`** (`setdefault`);
- **`subprocess.run`**: `encoding="utf-8"`, `errors="replace"`.

## Как запускать

```bash
python -m pytest tests/test_recover_stuck_plan_plates_cli.py -v
```

## Связанные файлы

- `tests/test_recover_stuck_plan_plates_cli.py`
- `scripts/recover_stuck_plan_plates.py`
- `core/kp_db.py` (`init_schema`)
