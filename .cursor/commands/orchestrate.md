---
name: orchestrate
description: Full orchestration - Pre-flight → Plan (checkpoint) → DAG Task Loop (conditional pipelines) → Docs. For complex features and systems.
---

# Orchestrate Command

## ⛔ YOU ARE FORBIDDEN FROM DOING ANY WORK YOURSELF

**Do NOT write code. Do NOT edit files. Do NOT plan or design without a subagent.**

Every single step must be executed by a subagent via the `Task` tool.
You are the coordinator only. If you find yourself about to do anything besides calling `Task` or updating orchestration workspace metadata (`progress.json` / `tasks.json` / plan status marks) — STOP.

---

## MANDATORY: Read and follow the skill

1. Read `.cursor/skills/orchestration/SKILL.md` using the Read tool — right now, before anything else
2. Execute EXACTLY as described in the skill — using `Task(subagent_type=...)` for each step
3. Do not skip, summarize, or shortcut any step from the skill

## Modes

| User says | Action |
|-----------|--------|
| `/orchestrate [task]` | Pre-flight → planner → **checkpoint (wait for user)** → execute DAG |
| `/orchestrate execute [id]` | Skip checkpoint; run/resume from workspace |
| `/orchestrate resume [id]` | Resume `active/` or `failed/` orchestration |

## Coordinator checklist

- [ ] Inject `plan-web-context` / project context into planner and workers
- [ ] Planner writes `type`, `dependsOn`, `pipeline` per task
- [ ] After plan: show DAG and **wait for approval** (unless execute/resume)
- [ ] Schedule by DAG; conditional pipelines (explore, security-auditor, senior-reviewer, refactor)
- [ ] Skip test-writer/reviewer when type is docs/spike/chore as skill says
- [ ] On max retries: mark failed, keep workspace resumable, ask user
