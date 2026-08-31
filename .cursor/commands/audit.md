---
name: audit
description: Module audit for Шишов — registry + Russian delta (gsm|kp|layout|auth|--since|--full). No Health Score, no auto-fix.
---

# Audit Command

## Coordinator only — with a narrow Step 0

You are the coordinator. Subagents do architecture, security, quality, and the written report.

### Step 0 (you, before any Task)

You **must**:

1. Read `.cursor/skills/audit-workflow/SKILL.md` — right now, then follow it exactly
2. Read `.cursor/skills/project-shishov/SKILL.md`
3. Read `ai_docs/develop/audits/FINDINGS.md`
4. Read the last audit report for this scope (if any)
5. Resolve scope per the skill. If the user gave **no** module, `--since`, path, or `--full` — **ask** (AskQuestion). Do **not** default to the whole repo.
6. For `--since` only: `git diff --name-only <ref>`

### Forbidden

- Analyze application source (`app/`, `core/`, `frontend/src/`, `tests/`, `viz_modules/`)
- Write the report or `FINDINGS.md` yourself — `documenter` does that
- Spawn remediation (`refactor`, `planner`, `worker`, `debugger`) — **no Phase 5, no auto-fix**
- Compute or print a Health Score 0–10
- Skip, summarize, or shortcut skill steps

Every analysis/report step is a `Task`. If you are about to review product code or patch it — STOP.

## Model: Composer 2 only

**Every** `Task` **must** set `model="composer-2-fast"`. Do not omit `model`. Do not use another slug. Applies to senior-reviewer, security-auditor, reviewer, and documenter.

## Scope cheat sheet

`gsm` | `kp` | `layout` | `auth` | `--since <ref>` | `--full` | a concrete path

`--full` = explicit whole-project registry sweep only. Bare `/audit` → ask the user.

## Prompts

Every Task prompt **must** include the mandatory preamble from the skill (project-shishov, FINDINGS.md, last report, domain checklist, Russian, evidence rule, no bot, no IDOR-on-[S1], report-only).
