# Report: GSM — один месяц, один комплект

**Date:** 2026-08-26  
**Orchestration:** `orch-2026-08-26-gsm-month-close-kit`  
**Status:** ✅ Completed (implementation). Коммита нет.  
**Spec:** [`../../specs/gsm-month-close-kit.md`](../../specs/gsm-month-close-kit.md)  
**Plan:** [`../plans/2026-08-26-gsm-month-close-kit.md`](../plans/2026-08-26-gsm-month-close-kit.md)

## Summary

Обзор ГСМ закрывает календарный месяц одним комплектом (сводка + бланки, `draft` → `exported`). Хвост и разрыв цепи режут комплект и generate **по машине**, не по флоту. Live `plita.db` не писалась.

## Tasks

| Task | Содержание | Статус |
|------|------------|--------|
| T1 | Overview: `open_before_month`, `chain_broken` | ✅ |
| T2 | Отчёт: `KIT_STATUSES`, skip красных, draft → exported | ✅ |
| T3 | Типы + подпись месяца + `planKit` | ✅ |
| T4 | Один календарный месяц (журнал = фильтры) | ✅ |
| T5 | Обзор: комплект, баннер, строка, пересчёт | ✅ |

## Spec SC 1–7

| SC | Критерий | Tasks | Вердикт |
|----|----------|-------|---------|
| 1 | Август + хвост июля Monjaro / Palisade чистая: баннер «Июль не выгружен»; журнал = август; bulk 0 галок = noop; галка Palisade → generate; галка Monjaro → skip; отчёт без галок — Palisade в zip, Monjaro в alert | T1, T3, T4, T5 | ✅ (vitest + pytest; live UI не гоняли) |
| 2 | Стрелка «предыдущий месяц» → 01.07–31.07 в фильтрах и сетке | T4 | ✅ |
| 3 | Комплект июля без красных: N сводки = N бланков = N draft → `exported` | T2, T5 | ✅ |
| 4 | Красный день: нет в zip, alert с датой, статус остаётся draft | T2, T3, T5 | ✅ |
| 5 | `chain_broken` на августе: комплект disabled, «Пересчитать август» → generate; после generate стык → `false` | T1, T3, T5 | ✅ (стык после generate — контракт API + invalidate overview; live generate не гоняли) |
| 6 | Май 848 usage-report зелёный | T2 | ✅ |
| 7 | Нет «Экспорт zip выбранных»; «Экспорт» в строке — `button` | T5 | ✅ |

## T5 review

Один цикл, verdict APPROVE. Правки:

- Recalc gated on tail (не предлагать «Пересчитать», пока открыт хвост).
- Kit / generate invalidate overview (чтобы бейдж хвоста и `chain_broken` обновлялись).

## Metrics

- pytest (overview + usage-report + repository): **56 passed**
- vitest `src/features/gsm/`: **218 passed**
- May 848 usage-report: **green**

## Leftovers (FYI, не блокеры)

- Visual UI checkpoint не делали в этом прогоне (нет live generate, нет browser pass). Чекбокс плана оставлен открытым.
- Yellow confirm skipped: на `FleetOverviewRow` нет yellow-поля.
- Nit: неиспользуемый `CONFIRMED_STATUSES`; leftover `UsageReportDialog`.

## Constraints honored

- Live `plita.db` не писали
- `core/gsm/generator.py` / `_resolve_start` не трогали
- `POST /gsm/waybills/export` оставлен в API
- Манифеста исключений в zip нет (UI alert + backend skip)
- `_rechain` не вызывали
- Git commit не делали
