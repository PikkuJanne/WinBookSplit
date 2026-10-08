# Next session

Next task: **M0-T03 — Freeze support and dependency decisions**. M0-T02 implementation C `8f1f556f318be488839e839830264fa2e5f4fbc8` was tested on clean C, normally pushed and freshly SYNCED at `2026-10-08T17:00:16.304703+00:00` (19:00:16 Europe/Berlin). This evidence/status checkpoint receives its own final receipt in [draft PR #2](https://github.com/PikkuJanne/WinBookSplit/pull/2)/the thread. Verify actual clean/live equality before starting and repair an incomplete evidence push first. Historical receipts do not prove current synchronization.

Workspace: `D:\projects\WinBookSplit-main`.
Branch/upstream: `codex/winbooksplit-v1-m0` / `origin/codex/winbooksplit-v1-m0`.
Origin fetch/push: `https://github.com/PikkuJanne/WinBookSplit.git`.
PR #1 was already merged when M0-T02 started; main then was `1c8c681cae0019dc2f413803e150951a962287ff`. A normal fetch and fast-forward brought that actual history into the feature branch. Continue draft PR #2 for the remaining M0 work; M0-T04 still owns cumulative review/test scaffolding/merge. Do not reuse closed PR #1 or infer the milestone gate is done from its early merge.

Read AGENTS.md, STATUS.md, TASKS.json, SCOPE_AND_DECISIONS.md, GITHUB_WORKFLOW.md, M0-T02 evidence, tasks/M0-T03.md and its specs. Start with:

```powershell
git status --short --branch
git remote -v
git worktree list
git rev-parse --abbrev-ref --symbolic-full-name '@{upstream}'
git rev-parse '@{upstream}'
python .\tools\codex-handoff\check_sync.py --repo .
python .\tools\codex-handoff\validate_plan.py --plan-root .\docs\codex-v1.0.0
```

Inspect live main, feature branch, PR state, tags/releases and authentication as well. Do not trust these records over actual local/remote state.

## M0-T02 results to preserve

`tests/fixtures/generate_pdf_fixtures.py` authors original MIT 10/12-page PDFs with visible IDs, deterministic anonymous metadata and known outlines. `tests/baseline/characterize_original.py` executes the unchanged embedded body in owned temporary storage and reopens every slice to verify page IDs. `expected_original.json` intentionally contains known-bad results; corrected targets remain in `PLAN_ORACLES.json`. Do not call a successful characterization a fixed-engine acceptance pass.

[M0-T02-baseline.md](evidence/M0-T02-baseline.md) and [M0-T02-characterization.json](evidence/M0-T02-characterization.json) cover AC-004/005/006. Actual findings: `1` writes no files with exit 0; `1,4,7` loses pages 1–3; invalid manual tokens are filtered; Level 1 omits front matter; nested Level 2 A2 includes B-opening pages 9–10. `4,7` is a complete-coverage control. Both launcher early error paths also return 0. All seven original application blobs remain unchanged.

Actual development environment: Windows 11 build 26300.9457; orchestration PowerShell 7.6.5; bounded launcher probe Windows PowerShell 5.1.26100.9444; explicit bundled Python 3.12.14 with pypdf 6.10.0 and ReportLab 4.4.9; Poppler 26.07.0. The default PATH Python 3.14.7 is separate and had no pypdf. Calibre was not found in M0-T01's inspected locations and no conversion was executed. Probe current reality for M0-T03; no supported-version decision exists yet.

Optional baseline rerun, using a new explicit external report path:

```powershell
$FixturePython = 'C:/Users/jtvuo/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe'
$ReportPath = Join-Path $env:TEMP ('WinBookSplit-baseline-' + [guid]::NewGuid().ToString('N') + '.json')
& $FixturePython -I .\tests\baseline\characterize_original.py --launcher-probes --report $ReportPath
if ($LASTEXITCODE -ne 0) { throw 'Baseline characterization failed' }
```

This uses a bundled development runtime, not the selected supported application interpreter. No global dependencies were installed. Real Windows success-path/Explorer/Calibre/release-package checks and fixed-engine tests remain NOT RUN.

## M0-T03 scope

Probe actual Windows/shell/Python/pypdf/Calibre tools and consult current official requirements/advisories. Choose exact supported versions and tested Python range; separate runtime/developer dependencies and provide explicit isolated-venv/path setup. Record Windows 11 PowerShell 5.1/7 targets, required EPUB/AZW3 conversion support, encrypted-feature rejection, optional unsigned release, and unclaimed Windows 10/UNC status. Do not modify runtime splitting or publish a release in this task. See tasks/M0-T03.md, specs/PROCESS_AND_PATHS.md, specs/PDF_SUPPORT.md and SOURCES.md.

## Continuation prompt for a new thread

Continue PikkuJanne/WinBookSplit toward its sole public v1.0.0 release in this existing checkout. Read applicable AGENTS.md and docs/codex-v1.0.0/{STATUS.md,NEXT_SESSION.md,TASKS.json,SCOPE_AND_DECISIONS.md} plus the next task/specs. Inspect actual origins, branch/upstream/worktree, dirt, HEAD and live GitHub; repair incomplete synchronization first. Preserve unrelated work and never reset, discard, auto-stash or force push.

Complete one dependency-ready task, starting with M0-T03 after freshly verifying M0-T02's evidence checkpoint. Preserve batch/PowerShell console and drag/drop, Python/pypdf splitting and optional local Calibre tooling. All document processing stays local; no website or interim public release. Baseline known-bad checks demonstrate reproduction, not correct behavior.

Scoped implementation/tests/commits/normal pushes/PRs/reviewed merges, plus final v1.0.0 publication after all recorded gates, are delegated. Respect protections and tool approval constraints; no destructive history, tag replacement, deletion, visibility/settings change, scope expansion, paid service or new public version is authorized. Record actual outcomes/skips/source hashes, update evidence/tasks/status/continuation, stage intended files only, push and verify clean HEAD equals a fresh live branch SHA. UNKNOWN/UNSYNCED is not a completed checkpoint.

End with task outcome, actual tests/gaps, C/checkpoint commits, branch/PR, live synchronization and exact next task. Only RELEASE_RUNBOOK.md's public/download-verified v1.0.0 and synchronized final main completes the project. Use a fresh thread for the next focused task.
