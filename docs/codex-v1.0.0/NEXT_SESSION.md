# Next session

Current task: M0-T01, pending implementation push and live verification. Repair/finish that checkpoint before starting M0-T02. Do not trust this file over current GitHub reality.

Workspace: `D:\projects\WinBookSplit-main`.
Branch: `codex/winbooksplit-v1-m0`; origin fetch/push `https://github.com/PikkuJanne/WinBookSplit.git`.
Confirmed live baseline: `6edbed7c1a0c94968999882c5a46d90492d3c327`; existing application files remain unchanged.
The directory originally lacked `.git`; only freshly fetched live history/index metadata was attached after an exact seven-file blob comparison. No previous local Git metadata was available to audit.

Read AGENTS.md, STATUS.md, TASKS.json, SCOPE_AND_DECISIONS.md, GITHUB_WORKFLOW.md, M0-T01 evidence and the selected task brief. Start with:

```powershell
git status --short --branch
git remote -v
python .\tools\codex-handoff\check_sync.py --repo .
python .\tools\codex-handoff\validate_plan.py --plan-root .\docs\codex-v1.0.0
```

After M0-T01 is synchronized, the exact next task is M0-T02: generate original page-marked PDF fixtures and characterize the original embedded engine/launcher defects. Read tasks/M0-T02.md, BASELINE_AUDIT.md, PLAN_ORACLES.json and TESTING.md. pypdf is absent from the current Python 3.14.7; Calibre was not located. Resolve required fixture dependencies locally in isolation; no supported-version claim exists yet. No application or Windows UI/Calibre test has passed.

## Continuation prompt for a new thread

Continue PikkuJanne/WinBookSplit toward the sole public v1.0.0 release in this existing local checkout. Read applicable AGENTS.md and docs/codex-v1.0.0/{STATUS.md,NEXT_SESSION.md,TASKS.json,SCOPE_AND_DECISIONS.md} plus the next task brief. Inspect actual origins, branch, dirt, HEAD and live GitHub; repair any incomplete synchronization checkpoint first. Never reset/discard unrelated work or force push.

Complete one dependency-ready task, starting with M0-T02 after the recorded M0-T01 checkpoint is freshly verified. Preserve the batch launcher, PowerShell console/drag-and-drop workflow, Python/pypdf splitter and optional local Calibre conversion. All document processing stays local; no website or intermediate public release.

Normal scoped implementation, tests, commits, feature pushes, PRs and reviewed merges, plus final v1.0.0 publication after all recorded gates, are authorized. Respect tool restrictions; no destructive history, work/tag/release deletion, protection/visibility/settings changes or paid services. Record actual tests/skips/evidence, update durable task/status/continuation records, commit intended files, push and verify clean HEAD equals a fresh live branch SHA. UNKNOWN/UNSYNCED is not completion.

End with task outcome, actual tests, gaps, code/checkpoint commits, branch/PR, live synchronization and exact next task. Only RELEASE_RUNBOOK.md's public/download-verified v1.0.0 and synchronized final main completes the project. Use a fresh thread for the next focused task.
