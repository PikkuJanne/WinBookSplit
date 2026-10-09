# Next session

Next task: **M1-T06 — Distinguish no-outline and invalid-document results**.
M0-T01 through M0-T04 and M1-T01 through M1-T05 are done.
Implementation C `3b3e97ec0551e9ed1ac72db5bcf98fe67ca1daaf` was normally pushed; clean HEAD matched
live feature at `2026-10-09T05:45:17.796145+00:00` (07:45:17 Europe/Berlin).
Following evidence/status E receives its own receipt in the thread/PR; recheck.

Workspace: `D:\projects\WinBookSplit-main`; branch/upstream
`codex/winbooksplit-v1-m1` / `origin/codex/winbooksplit-v1-m1`.
Origin: `https://github.com/PikkuJanne/WinBookSplit.git`.
[Draft PR #9](https://github.com/PikkuJanne/WinBookSplit/pull/9) continues M1.
PR #8 was already merged at `2026-10-09T05:22:09Z` to main
`828b232b6a25b2c67c65fc3e31d2f3f7cb30a38c`; normal fast-forwards reconciled
feature/local main. Preserve unrelated history/edits and reconcile fresh changes
normally. Earlier merges do not waive M1-T06 cumulative review/merge requirements.

Read AGENTS.md, STATUS.md, TASKS.json, SCOPE_AND_DECISIONS.md,
GITHUB_WORKFLOW.md, TESTING.md, SUPPORT_AND_SETUP.md, evidence/M1-T05-plan.md/json,
tasks/M1-T06.md, specs/SPLIT_CONTRACT.md, specs/CLI_AND_UX.md and PLAN_ORACLES.json.
Inspect actual origin/upstream/branch/worktree/authentication/live main/feature/
PR/protections/checks/tags/releases before changing files:

```powershell
git status --short --branch
git remote -v
git worktree list
git rev-parse --abbrev-ref --symbolic-full-name '@{upstream}'
python -B .\tools\codex-handoff\check_sync.py --repo .
python -B .\tools\codex-handoff\validate_plan.py --plan-root .\docs\codex-v1.0.0
```

Sibling-import handoff helpers use -B, without -I. Application tests use a fresh
supported absolute developer Python -I -B. Last live audit found no protections/
rulesets/workflows/runs/tags/releases; historical receipts are not current proof.
No reset/force push/auto-stash/deletion/settings changes.

## Shared contract and verified behavior to preserve

Existing plan_manual_starts/normalize_outline/normalize_bookmarks/plan_level1/
plan_level2 logic is unchanged through M1-T05. Preserve complete manual starts
(ASCII tokens, whole document for 1, sorting/dedup notices/page 1), Level 1
source order/lineage/aliases/front matter and bounded raw/normalized traversal
(depth 64, 10,000 nodes/definitions/entries including containers), and Level 2
own retained direct children, parent openings/fallbacks, no boundary crossing,
alias/invalid-parent diagnostics and global usable-child gate. Deeper entries
remain metadata. Sentinel/55 applies only the established no-plan categories;
malformed/zero-page fails, no silent noninteractive fallback.

M1-T05 adds PlanError, validate_plan, frozen PreparedSplit, prepare_split,
preview_plan and execute_split. Every production mode finalizes one shared
deeply immutable plan and derived complete coverage. Preview writes no chapters;
execution uses its exact entries/filenames on the captured private pypdf reader,
without reopening or replanning. The result retains source/coverage/entries and
page counts. Source path changes/deletion safely use the original reader snapshot;
new PDF/conversion requires a new job. Public model substitution/changed captured
stream rejects. Private cache/memory manipulation is outside this API's contract.
Original/resolved source aliases and all existing outputs reject before writer;
xb refuses a subsequent race. Full output staging/rollback remain later tasks.

Final seven-stage full: 105 Python tests; nine Pester per actual
PS5.1.26100.9444/PS7.6.5; no skips/syntax/scaffold findings. Plan layer has nine
structural/two preview/three real modes/two source aggregates plus deletion and
300 seeded plans, exact IDs/content; stable-source plan units: 18. Manual/L1/L2
keep 22/4/4 targets, 8/5/6 aggregates, 250/150/150 plans and 25/10/10 writers.
Six actual owned launchers include three concurrent, corrected PS7 BM-03 bytes/
IDs and preserved sources/neighbors/decoys/shared TEMP. Remaining launchers cover
manual; corrected Level 2 PS5.1/BAT and Level 1 console remain unclaimed.
Seven expected NullObject warnings, PS5.1 CLIXML progress, PS7 launcher RawUI
stderr and 60 legacy analyzer observations each remain recorded limitations.
Tested 41-path digest `73f7c3848cf35dcd3444cdd5af861231e81efe775d9afff298eac22abc2defae`; exact bytes matched clean C,
without a clean-C full rerun. Earlier setup/unit-consistency observations and a
passing pre-final-report-guard full are preserved separately in task evidence.

## M1-T06 boundary and test route

Implement AC-030/031: distinct no-outline/no-selected-level/no-usable/read/
malformed/zero-page categories, no zero-output success, correct applicable manual
or Level 1 fallback decision without silently selecting it. Structured diagnostic
handler/decision tests belong here; complete interactive fallback UX belongs M3.
Run affected and full M1 regressions, inspect cumulative source/evidence/CI and
review/merge through the normal available process; synchronize actual main.
Do not claim submitted review/CI/human UI where they did not run. Only M1-T06
is selected; no launch into M2 or release tasks in the same thread.

Use tests/run_tests.py --layer plan, level2, bookmarks or manual for focused
checks; full has seven stages and requires explicit actual hosts/tool root.
Reports must be new absolute external files from unrelated CWD. Plan/bookmark/
level2 routes have no shell args. Manual historical bookmark comparisons are
empty; its corrected six-section BM-03 actual CLI reference supplies unchanged
owned launchers. Preserve immutable original/extraction guards/oracles/generator
and all older evidence. Historical AC-011 reads M1-T01 C `88c2149`; explicit old
diagnostics intentionally refuse changed source. Never weaken those guards.

## Runtime and remaining gates

Regular GIL CPython 3.14.8 x64 only/pypdf 6.19.0; developer ReportLab 5.0.1,
Pillow 12.3.0, charset-normalizer 3.5.2. Fresh hash-required isolated venv/exact
import origins; old TEMP may be absent. External Pester 6.2.0/PSScriptAnalyzer
1.25.0 pins: reverify all 66 file hashes/sizes and actual absolute hosts. Children
use process-only RemoteSigned/host-owned modules; stored policy stays unchanged.
Official portable Calibre 9.15.0 remains at `$UserProfile\Apps\Calibre915\Calibre Portable`;
converter `Calibre\ebook-convert.exe` beneath it. M1-T01 hash/signature/version
receipt is historical, not a conversion pass. Trusted-path discovery is M2-T04;
actual EPUB/AZW3 conversion is M4-T04/package/release work. No bundled binaries.
Ordinary launchers/discovery/error paths, interactive preview/Explorer, output
transactions, feature/fidelity, clean OS/extracted package, CI (M4-T05), public
release/download remain open. Windows10/ARM/UNC/other Python stay unclaimed.
No release/tag; only final public v1.0.0/package/download/fixed-tag/synchronized-
main runbook closes the project. Unsigned publication is allowed with disclosure.
