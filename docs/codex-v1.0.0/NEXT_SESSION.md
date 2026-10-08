# Next session

Next task: **M1-T01 — Extract the existing Python engine narrowly**.
M0-T01 through M0-T04 are done. M0's reviewed scaffold implementation C
`01ab01bccbbe16ea2bc03eeca0153b18b7c5bfee` was normally pushed and freshly
SYNCED at `2026-10-08T17:59:55.797912+00:00` (19:59:55 Europe/Berlin).
[PR #4](https://github.com/PikkuJanne/WinBookSplit/pull/4) merged at
`2026-10-08T18:01:01Z` into main `ae9ddd0444e61578ed59625d6d3df4038e6fb09e`.
Clean local main matched live main at `2026-10-08T18:01:05.818622+00:00`;
tested raw source bytes were rechecked after switching branches.
This following documentation-only closure receives its own final normal
push/live receipt in the thread. Recheck current equality before starting;
historical receipts and recorded done status do not prove current sync.

Workspace: `D:\projects\WinBookSplit-main`. Active branch/upstream:
`main` / `origin/main`. Origin fetch/push:
`https://github.com/PikkuJanne/WinBookSplit.git`.
Retained milestone branch `codex/winbooksplit-v1-m0` points to C; do not delete
or reuse closed PRs. Create a normal `codex/` M1 milestone branch from the
freshly verified current main. Never reset to recorded historical commits.

Read AGENTS.md, STATUS.md, TASKS.json, SCOPE_AND_DECISIONS.md,
GITHUB_WORKFLOW.md, TESTING.md, SUPPORT_AND_SETUP.md,
evidence/M0-T04-scaffolding.md, evidence/M0-T04-harness.json,
tasks/M1-T01.md and specs/PROCESS_AND_PATHS.md. Start with actual inspection:

```powershell
git status --short --branch
git remote -v
git worktree list
git rev-parse --abbrev-ref --symbolic-full-name '@{upstream}'
git rev-parse '@{upstream}'
python -B .\tools\codex-handoff\check_sync.py --repo .
python -B .\tools\codex-handoff\validate_plan.py --plan-root .\docs\codex-v1.0.0
```

Inspect authentication, live main/feature refs, PR state/protection/required
checks and tags/releases separately. Repair unknown/unsynced checkpoints
first. Preserve unrelated dirt/history; no auto-stash, force, reset or deletion.

## M0 evidence to preserve

AC-009 passed actual native exit 23, Python unittest failure and one failed
Pester test in each host through the central runner; later commands/reporting
succeeded but overall exit stayed 1. AC-010 passed two full unrelated-directory
runs: 11 Python tests, 9 Pester tests per actual Windows PowerShell
5.1.26100.9444/PowerShell 7.6.5, zero syntax/new-scaffold static findings,
14 original known-bad engine cases plus two early launcher boundaries.
Input/neighbor/all-checkout bytes were preserved and owned temps removed.
Actual tested-path digest:
`fc8e5fd5ebb1001beba6d1a4f072b9172fc444410336c029e84841834c1bd536`.
Use raw bytes and `--no-filters` when comparing original blobs under autocrlf.

Independent cumulative M0 review found no remaining blocker; all seven
original application files remain byte-identical. External handoff helper
suite: 56 passed/one real symlink-creation skip. All 66 isolated shell-tool
file hashes/sizes were independently checked. The unchanged application has
68 legacy analyzer observations per host; these are not a runtime static pass.
Main had no protection/rulesets/required checks or workflows/runs. No GitHub
review submission or CI execution is claimed; CI remains M5-T01.

Regular GIL CPython **3.14.8 x64 only**, plain **pypdf 6.19.0**;
dev-only ReportLab **5.0.1**, Pillow **12.3.0**, charset-normalizer **3.5.2**.
Fresh hash-required install/pip check/exact import-origin/version assertions
passed. Follow SUPPORT_AND_SETUP.md for a new explicit dev venv, rather than
assuming an old TEMP directory still exists or selecting PATH Python as
application support. Handoff helpers need only stdlib.

Use `tests/run_tests.py` with explicit interpreter `-I -B`, a new absolute
external report, external shell tools and explicit host executables.
See tests/README.md and tests/powershell/README.md for full/targeted commands
and exact Save-Module setup. Pester **6.2.0** and PSScriptAnalyzer **1.25.0**
are imported by absolute versioned manifests, never inbox substitution.
Test children request process-only RemoteSigned with their own host built-in
module path; stored user/machine policy is unchanged and managed policy wins.
The harness is developer tooling, not ordinary application-launch acceptance.

Original characterization verifies visible page identities/order and its
known-bad results separately from corrected PLAN_ORACLES.json targets.
Reproduced defects remain: manual `1` creates zero files/exit 0; `1,4,7`
loses pages 1-3; invalid tokens are filtered; Level 1 omits front matter;
nested Level 2 A2 crosses into B pages 9-10. `4,7` is the complete-coverage
control. Both original early launcher rejection paths return 0.
Fixtures are original MIT 10/12-page generated PDFs; no user documents.
Do not weaken expectations or call reproduced defects repaired acceptance.

## M1-T01 scope

Move only the existing embedded body into shipped
`engine/winbooksplit_engine.py`, with import-safe functions and guarded entry.
Resolve the shipped engine from the PowerShell script root, preserving the
existing batch/PowerShell workflow and Python/pypdf. Do not fix parsers,
bookmarks, coverage or output behavior during mechanical extraction.

AC-011: compare understood pre/post-extraction characterization and prove
import has no processing side effects. The original harness deliberately
refuses changed launcher hashes; add an explicit extraction-equivalence route
with separate source identities instead of silently treating modified code
as the original. Preserve historical receipts/expected_original.json.
AC-012: exercise both entry points from unrelated directories and parallel
launches, locating the shipped engine without a shared generated engine file.
Record actual supported host/interpreter paths and limits. Dependency discovery
redesign remains M2-T04; engine bug fixes start M1-T02.

Run targeted plus affected regressions, obtain independent source/evidence
review, update evidence/tasks/status/next-session, stage intended files only,
commit/push normally and verify clean HEAD equals a fresh live branch SHA.
One focused task per thread; leave M1-T02 for its own task checkpoint.

Calibre **9.15.0** remains absent; actual EPUB/AZW3 conversion is mandatory
later. Full repaired application splitting, Explorer, renderer **26.10.0**,
CI, release-package and public-download verification remain NOT RUN.
Windows10/ARM/UNC/other Python are unclaimed. No tag/release exists.
Only RELEASE_RUNBOOK.md's verified public non-draft/non-prerelease v1.0.0,
matching anonymous downloads, fixed tag and synchronized final main finish
the project. Unsigned release is allowed with disclosure/checksums.
