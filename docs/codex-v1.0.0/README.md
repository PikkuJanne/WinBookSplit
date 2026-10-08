# Implementation control room

This directory is the durable Codex handoff. It is not application documentation for end users and must not be shipped in the runtime ZIP.

Read `STATUS.md` and `NEXT_SESSION.md` at every thread start. `TASKS.json` is the canonical task graph; `tasks/*.md` are implementation briefs. `ROADMAP.md` groups tasks, `IMPROVEMENT_MAP.md` maps the approved review, and `SCOPE_AND_DECISIONS.md` prevents scope drift. `ACCEPTANCE_CASES.json` is the canonical acceptance catalog; none of its entries is evidence of a pass. `PLAN_ORACLES.json` provides deterministic expected split ranges.

Use `GITHUB_WORKFLOW.md` for checkpoint/PR handling, `TESTING.md` for validation, and `RELEASE_RUNBOOK.md` for the actual finish. Templates belong under `templates/`; fill copies under `evidence/`, never mark a template as completed evidence. Keep `AGENTS.md` small and read detail on demand.

All tasks start `todo`, all test evidence starts unexecuted, and release status starts `not_started`. If the repository has moved since baseline, M0 reconciles the current state without discarding work.

## Helper commands after import

```powershell
python .\tools\codex-handoff\validate_plan.py --plan-root .\docs\codex-v1.0.0
python .\tools\codex-handoff\check_sync.py --repo .
```

The sync checker reads Git and performs a fresh `ls-remote`; it never commits, fetches, checks out, pushes, or publishes. Exit 0 means clean/equal at the observed instant, 1 means unsynchronized, and 2 means unknown/configuration/error. GitHub API/PR/CI/release checks remain separate. A green sync check is not a release gate by itself.
