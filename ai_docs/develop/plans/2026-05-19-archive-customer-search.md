# Plan: Поиск заказчика в архиве КП (web)

**Created:** 2026-05-19  
**Orchestration:** `orch-2026-05-19-archive-customer-search`  
**Goal:** Добавить поиск по названию заказчика в архив коммерческих предложений (веб) в дополнение к поиску по номеру КП.  
**Total Tasks:** 7  
**Priority:** High  

> Полная версия плана с деталями реализации: `.cursor/workspace/active/orch-2026-05-19-archive-customer-search/plan.md`

## Tasks

- [ ] **ARCH-001** — `core/kp_db.py`: `_escape_sql_like`, `search_kp_by_customer_name`
- [ ] **ARCH-002** — `KpArchiveRepository` + `ArchiveService.search`
- [ ] **ARCH-003** — `ArchiveSearchResponse` + `GET /search` endpoint
- [ ] **ARCH-004** — Backend tests (endpoints + service)
- [ ] **ARCH-005** — Frontend types, `archiveApi.search`, `useArchiveSearchQuery`
- [ ] **ARCH-006** — `ArchiveSearchBar` (два поля, валидация)
- [ ] **ARCH-007** — `CommercialOfferArchivePage` + `ArchiveOfferList` для результатов

## Work order

Backend: ARCH-001 → 002 → 003 → 004, then Frontend: 005 → 006 → 007.

**Constraint:** не менять `bot/handlers/archive.py`.
