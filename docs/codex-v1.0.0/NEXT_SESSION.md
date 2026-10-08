# Next session

Next task: **M0-T04 — Add local test and evidence scaffolding**. M0-T03 implementation C `2e63f259a59438a689e75ff0ff9c1a9fcd600c1f` passed the recorded isolated setup/dependency checks, was normally pushed and freshly SYNCED at `2026-10-08T17:33:12.039489+00:00` (19:33:12 Europe/Berlin). Clean-C structural plan validation and recorded source-path digest checks passed. This evidence/status checkpoint receives its own final receipt in [draft PR #3](https://github.com/PikkuJanne/WinBookSplit/pull/3)/the thread. Verify actual clean/live equality before starting and repair an incomplete evidence push first. Historical receipts do not prove current synchronization.

Workspace: `D:\projects\WinBookSplit-main`.
Branch/upstream: `codex/winbooksplit-v1-m0` / `origin/codex/winbooksplit-v1-m0`.
Origin fetch/push: `https://github.com/PikkuJanne/WinBookSplit.git`.
PR #2 was already merged when M0-T03 started; main then was `a66c8f95f922c36c58b47b6cbfba82be399a552a`. A normal fetch and fast-forward brought that actual history into the feature branch. Continue draft PR #3 for the remaining M0 work; M0-T04 owns cumulative review/test scaffolding/merge. Do not reuse closed PR #2 or infer the milestone gate is done from the earlier merges.

Read AGENTS.md, STATUS.md, TASKS.json, SCOPE_AND_DECISIONS.md, GITHUB_WORKFLOW.md, TESTING.md, SUPPORT_AND_SETUP.md, M0-T03 evidence, tasks/M0-T04.md and its specs. Start with:

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

Historical M0-T02 development environment: Windows 11 build 26300.9457; orchestration PowerShell 7.6.5; bounded launcher probe Windows PowerShell 5.1.26100.9444; explicit bundled Python 3.12.14 with pypdf 6.10.0 and ReportLab 4.4.9; Poppler 26.07.0. These dependency versions remain historical evidence, not the current supported pins. The default PATH Python 3.14.7 remains separate with no pypdf. M0-T03 freshly probed the workstation/hosts and Calibre absence, then selected and tested the updated isolated dependencies below.

Optional baseline rerun, after following SUPPORT_AND_SETUP.md to create the selected fresh dev venv; set its absolute interpreter path and a new external report path:

```powershell
$FixturePython = 'C:\Tools\WinBookSplit\.venv-dev\Scripts\python.exe'
$ReportPath = Join-Path $env:TEMP ('WinBookSplit-baseline-' + [guid]::NewGuid().ToString('N') + '.json')
& $FixturePython -I .\tests\baseline\characterize_original.py --launcher-probes --report $ReportPath
if ($LASTEXITCODE -ne 0) { throw 'Baseline characterization failed' }
```

Use regular x64 CPython 3.14.8 with the hashed developer requirements; do not rely on an old TEMP venv or confuse PATH Python with the fixture interpreter. Handoff helper scripts can still use PATH Python without selecting it as application support. No global dependencies were installed. Real Windows success-path/Explorer/Calibre/release-package checks and fixed-engine tests remain NOT RUN.

## M0-T03 frozen decisions and evidence

Application minimum = selected = **regular GIL CPython 3.14.8 x64 only**, plain **pypdf 6.19.0**. Separate root runtime/dev hashed requirements; developer packages are ReportLab 5.0.1/Pillow 12.3.0/charset-normalizer 3.5.2. Fresh explicit runtime/dev installs, pip checks, import-origin/absence assertions and the complete documented runtime setup under both actual shell hosts passed. Initial PS5.1 native quoting failed and was independently reproduced; the corrected PowerShell double-quoted arguments with Python single-quoted literals passed in both hosts. A temporary PS5.1 -File probe was blocked by local execution policy; direct interactive setup commands passed without overriding/changing policy. Preserve the tested quoting and no-global-install boundary.

The original harness reproduced all 14 engine cases and two early launcher error paths on the final pins, with exact page IDs and unchanged sources/neighbors. New fixture bytes/hashes are recorded without overwriting historical M0-T02 receipts; this is not repaired-engine acceptance or a new visual-render pass. All seven original application files remain byte-identical; raw Git blob comparisons use `--no-filters` because `core.autocrlf=true` changes filtered hashes.

Calibre **9.15.0** is the selected mandatory EPUB/AZW3 conversion test target, absent and NOT RUN; optional only for PDF-only users. Current primary vendor requirements/advisories prompted newer Python/pypdf/ReportLab pins; no automated scanner claim. Windows10/ARM/UNC/other Python remain unclaimed. Unsigned release remains allowed with disclosure/checksums; encryption/forms/active-feature rejection remains a later implementation requirement.

## M0-T04 scope

Add one small local test/evidence runner, targeted/full invocations, owned temporary fixtures, acceptance/source-digest reports, meaningful native exit-code propagation and ignore/privacy boundaries. Preserve runtime behavior; engine extraction/fixes remain later tasks. Use **stdlib unittest**, not an unnecessary pytest dependency. Prepare **Pester 6.2.0** and **PSScriptAnalyzer 1.25.0** in an isolated tool directory with exact-version save/absolute-manifest import, then verify under both hosts. These modules are not installed/tested yet; inbox Pester 3.4.0 must not silently substitute. Poppler 26.07.0 is historical; 26.10.0 is the later fidelity target, with Windows binary provenance still to review before M4. See the frozen setup plan and M0-T04 AC-009/010.

Review the cumulative M0 diff/evidence, run appropriate complete layers, inspect actual required CI/review state, merge passing reviewed work through the available process, then synchronize clean local main. Preserve feature history. Full splitting/Explorer/Calibre/release-package checks remain open; do not weaken later gates.

## Continuation prompt for a new thread

Continue PikkuJanne/WinBookSplit toward its sole public v1.0.0 release in this existing checkout. Read applicable AGENTS.md and docs/codex-v1.0.0/{STATUS.md,NEXT_SESSION.md,TASKS.json,SCOPE_AND_DECISIONS.md} plus the next task/specs. Inspect actual origins, branch/upstream/worktree, dirt, HEAD and live GitHub; repair incomplete synchronization first. Preserve unrelated work and never reset, discard, auto-stash or force push.

Complete one dependency-ready task, starting with M0-T04 after freshly verifying M0-T03's evidence checkpoint. Preserve batch/PowerShell console and drag/drop, Python/pypdf splitting and optional local Calibre tooling. All document processing stays local; no website or interim public release. Baseline known-bad checks demonstrate reproduction, not correct behavior.

Scoped implementation/tests/commits/normal pushes/PRs/reviewed merges, plus final v1.0.0 publication after all recorded gates, are delegated. Respect protections and tool approval constraints; no destructive history, tag replacement, deletion, visibility/settings change, scope expansion, paid service or new public version is authorized. Record actual outcomes/skips/source hashes, update evidence/tasks/status/continuation, stage intended files only, push and verify clean HEAD equals a fresh live branch SHA. UNKNOWN/UNSYNCED is not a completed checkpoint.

End with task outcome, actual tests/gaps, C/checkpoint commits, branch/PR, live synchronization and exact next task. Only RELEASE_RUNBOOK.md's public/download-verified v1.0.0 and synchronized final main completes the project. Use a fresh thread for the next focused task.
