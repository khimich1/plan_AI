---
name: audit-workflow
description: Module-scoped audit for Шишов — registry + delta report in Russian (architecture, security, quality). No Health Score, no auto-remediation. Use with /audit gsm|kp|layout|auth|--since|--full.
---

# Audit Workflow (Шишов)

**Purpose**: After work on a module, produce a **Russian delta** — what closed, what is new, which P0/P1 are still open — by updating a stable findings registry. Not a new 500-line novel. Not a 0–10 score. Not auto-fix.

**Idea / spec**: `ai_docs/ideas/audit-shishov.md`

---

## Model: Composer 2 for every `Task` (`/audit`)

When the user runs **`/audit`**, every subagent call **must** use:

`model="composer-2-fast"`

Do not omit `model`. Subagents must not spawn `Task` with another model slug.

Sequential shape:

1. `Task(subagent_type="senior-reviewer", model="composer-2-fast", prompt="…")`
2. `Task(subagent_type="security-auditor", model="composer-2-fast", prompt="…")`
3. `Task(subagent_type="reviewer", model="composer-2-fast", prompt="…")`
4. Coordinator builds a delta table from subagent outputs **plus** the registry (Step 6). Do not review application source yourself.
5. `Task(subagent_type="documenter", model="composer-2-fast", prompt="…")` — writes the report **and** updates `FINDINGS.md`

There is **no Phase 5**. Do not spawn `refactor`, `planner`, `worker`, or `debugger` from `/audit`. Fixes: user runs `/orchestrate`, `/implement`, or `/refactor` explicitly.

---

## Coordinator: allowed vs forbidden

**Allowed (Step 0 — you must do this before any Task):**

- Read this skill, `.cursor/skills/project-shishov/SKILL.md`, `ai_docs/develop/audits/FINDINGS.md`, the last report for this scope
- `git diff --name-only` only when the user passed `--since`
- Ask the user (AskQuestion) if scope is missing — **do not default to full repo**

**Forbidden:**

- Read or analyze application source (`app/`, `core/`, `frontend/src/`, `tests/`, `viz_modules/`)
- Write the report or `FINDINGS.md` yourself (documenter does)
- Auto-remediation, Health Score, auditing `bot/`, `bot_archived/`, `tests/archived/` as live systems

---

## Step 0: Scope

| Input | Scope |
|-------|--------|
| `/audit gsm` | `app/services/gsm_*`, `app/api/v1/endpoints/gsm.py`, `app/repositories/gsm_repository.py`, `app/schemas/gsm.py`, `core/gsm/`, `frontend/src/features/gsm/` |
| `/audit kp` | commercial / offers / archive: `app/api/v1/endpoints/commercial.py`, `offers*`, `archive*`, `app/services/commercial*`, `app/security/offer_access.py`, `core/kp/`, related frontend commercial/archive |
| `/audit layout` | layout / ILP / plate runtime: `core/optimization/`, `core/plate_runtime_state.py`, `core/config_and_data.py`, `core/domain/plate_order.py`, `core/production/planning.py`, `core/ports/visualization.py`, `viz_modules/` |
| `/audit auth` | `app/security/`, session/CSRF/roles, login rate limit (not bot) |
| `/audit --since <ref>` | files from `git diff --name-only <ref>` (then map to the nearest module checklist) |
| `/audit --full` | registry sweep: critical 20–30% of live code (not bot, not archived tests). **Only if the user typed `--full` or explicitly «весь проект / --full».** |
| `/audit <path>` | that file or directory |
| `/audit` with no args | **Ask** — options: gsm, kp, layout, auth, `--since HEAD~20`, `--full`. Wait. Do not invent a default. |

Last report: newest `ai_docs/develop/audits/20*-{gsm|kp|layout|auth|full}*-audit.md` matching the scope. If none, say so in the prompt.

---

## Mandatory prompt preamble (every Task, including documenter)

Copy this block and fill the braces. Domain checklist: paste the matching section from this skill (gsm / kp / layout; auth uses the short auth list + generic second pass).

```
## Mandatory context (read FIRST)
1. Read `.cursor/skills/project-shishov/SKILL.md`
2. Read `ai_docs/develop/audits/FINDINGS.md`
   - REUSE existing IDs. One ID = one problem forever.
   - New ID only if nothing in the registry matches. Next free A/S/Q number (never recycle).
   - Do not drop an open finding because you did not look at it. If you cannot reproduce: status unreproduced, not resolved, not omitted.
3. Last report for this scope: {path or "none"}
4. Scope: {module} files: {paths}
5. Domain checklist (first pass): {paste}
6. Language: findings and the final report in Russian.
7. Evidence: Critical/High MUST cite file:line and/or test name and/or a grep that you actually ran. No evidence = do not emit Critical/High.
8. Skip bot/, bot_archived/, tests/archived/ as live systems.
9. Report only. Do not patch code. Do not propose launching remediation in this chat.
10. [S1] offer access is by-design shared archive (ADR offer-access-policy.md) — do not re-open as IDOR.
```

Then the role-specific checklist (architecture / security / quality) as a **second** pass.

---

## Step 2: Architecture (senior-reviewer)

First pass: domain checklist for the scope.  
Second pass (only inside scope): module boundaries, god modules, coupling, circular deps, over-engineering.

Output:

```markdown
## Architecture Findings

### Critical
- [A1] … (reuse ID) — evidence: `path:line` — status suggestion: open|resolved|unreproduced

### High
…

### Medium
… (only if there is a concrete action this week; else omit)

Delta vs registry: closed / still open / new / unreproduced
```

---

## Step 3: Security (security-auditor)

Pass architecture summary. First pass: domain + auth notes. Second: validation, secrets, cookies/CSRF, rate limit, PII, dependency vulns, error leakage — **inside scope**.

Do not flag shared-archive KP access as IDOR ([S1] by-design).

Same evidence and ID rules. Prefix `S`.

---

## Step 4: Code quality (reviewer)

Pass architecture + security summaries so you do not duplicate them. First pass: domain. Second: DRY, complexity, naming, error handling, dead code, tests, types — **inside scope**. Prefix `Q`.

---

## Domain checklists (≤15 each; invariants, not SOLID)

### gsm

1. Month-close: cannot generate the next month while the previous is open or the tank/odometer chain is broken.
2. Liter and odometer continuity across the month boundary (no silent `_rechain` of old km).
3. Same liter rounding rule on frontend and backend (no `Math.round` vs banker's round drift).
4. Report zip days == waybill days for the same period (including former drafts).
5. Generation starts from last `confirmed`/`exported`, not from a stale draft chain.
6. LibreOffice/`soffice` must not block the HTTP worker.
7. Import: file/row limits + unit of work (no partial commit without rollback).
8. Seasonal logic: single source of truth; frontend must not drift from the API contract.
9. Waybill CRUD vs generation: not one god service.
10. `manual_intervention` vehicles excluded from zip with a visible reason.
11. Re-export of an already exported month requires confirm.
12. `GsmGenerationError` (and user-facing GSM errors) in Russian.
13. Tests for chain break, rounding, and month-close gate.

### kp

1. Offer access is **shared archive by design** — do not report IDOR; `owner_user_id` is reserved, unused in policy.
2. Logistics must not receive KP financial fields if that restriction still exists in code/tests.
3. Thin HTTP handlers; orchestration in services, not in `commercial.py` / `production.py`.
4. Persistence through repositories — no raw SQL in services that already have a repo layer.
5. OCR/LLM: documents leave the factory — treat as a policy item, not a surprise.
6. Drafts: path traversal and ownership guards stay intact.
7. Destructive DB ops go through `destructive_db_guard` / admin guards.
8. Product-type pipeline: no sixth copy of the same draft/OCR flow.
9. Archive/move-to-production: errors visible, not `except → None`.
10. CSRF validated before parsing multipart bodies.
11. HTTP errors must not leak internals.
12. Wizard/archive god-hooks: note only if still huge **and** you have a split proposal with evidence.

### layout

1. `PlateOrderContext` passed explicitly — no business logic via `config_and_data` / mutable globals.
2. Isolation for HTTP **and** BackgroundTasks, CPU pool, CLI (no ContextVar leak).
3. Planning must not import matplotlib / `core.visualization` at module load.
4. Viz goes through `core/ports/visualization.py` (ports already exist — do not regress).
5. ILP/sequence objective matches the spec (waste, tracks, due dates) — evidence from tests or solver inputs, not vibes.
6. Sequence builder: no large dead branches; golden/hash tests still mean something.
7. Wide / unpriced plate resolve: one implementation.
8. Layout jobs must not depend on GSM.
9. In-process caches/rate limits: single-worker is a documented constraint, not silent multi-worker.

### auth (short; use with `/audit auth` or as add-on to kp)

- HttpOnly session cookie; CSRF double-submit (non-HttpOnly CSRF cookie is expected).
- Roles vs shared-archive policy (see kp item 1).
- Login/OCR rate limit: in-memory means single worker only ([A2]/[S3]).
- Do not treat `bot_archived/` as a live auth surface.
- Session TTL / refresh as written in `app/security/`.

---

## Evidence rule

| Severity | Required |
|----------|----------|
| Critical / High | File:line **or** failing/passing test name **or** grep you ran. Else drop or demote to Medium without the label Critical/High. |
| Medium | Only if there is an action this week. Otherwise omit from the report (registry row can stay). |
| Low | Do not add new Lows in module audits. `--full` may list Lows in an appendix only. |

`unreproduced` ≠ `resolved`. Omitting an open ID ≠ closed.

---

## Step 6: Delta (coordinator) — no Health Score

Do **not** compute a 0–10 score.

From registry ∩ this scope, plus new IDs from subagents:

```
closed:        was open, subagent proved resolved (evidence)
still_open:    registry open, still valid
new:           new ID assigned
unreproduced:  registry open, not verified this run
by-design / wontfix: do not "fix" and do not re-litigate
```

**Metric:** count of **open P0** and **open P1** in this scope.

- P0 = open + Critical
- P1 = open + High

---

## Step 7: Report + registry (documenter)

**Report path:** `{auditsPath}/YYYY-MM-DD-{scope-slug}-audit.md`  
`auditsPath` from `.cursor/config.json` → `documentation.paths.audits` (default `ai_docs/develop/audits`).

Also update `ai_docs/develop/audits/FINDINGS.md`: status, last_seen, evidence, notes. Add new rows for new IDs. Never delete rows. Never reuse an ID.

### Report format (Russian)

```markdown
# Аудит: {scope}

**Дата**: YYYY-MM-DD
**Скоуп**: …
**Реестр**: FINDINGS.md
**Прогон 0**: {previous report or —}

## Дельта

| ID | Было | Стало | Суть |
|----|------|-------|------|
| A3 | open | resolved | … |

- Закрыто: N
- Открыто (подтверждено): N
- Новое: N
- Не воспроизвелось: N

**Открытые P0 / P1 в скоупе**: {n} / {m}

## Действия (максимум 8)

1. …

## Открытые P0 / P1

### [A1] …
**Статус**: open
**Улика**: `file:line` / тест
**Зачем**: …

## Приложение

Подробности Medium и контекст. Без Health Score. Без «Start remediation?».
```

After the report: tell the user the path, the P0/P1 counts, and that fixes are a separate command. **Do not ask to auto-fix.**

---

## What /audit does not do

- Auto-fix Critical/High/anything (no Phase 5)
- Health Score 0–10
- Default to full-repo
- Invent new IDs for old problems
- Audit the Telegram bot as a live product
- Treat `[S1]` shared archive as a vulnerability

For remaining work the user can run `/refactor`, `/implement`, or `/orchestrate`.
