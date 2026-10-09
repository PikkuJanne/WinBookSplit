# Next session

Next task: **M1-T03 — Normalize bookmarks and fix Level 1 coverage**.
M0-T01 through M0-T04 and M1-T01/M1-T02 are done. M1-T02 implementation C
`ced0d96f0540d96adc2fdacc4e1ff31fa25be7b8` was normally pushed; clean local
HEAD matched the fresh live feature branch at `2026-10-09T03:39:03.607417+00:00`
(05:39:03 Europe/Berlin). The following evidence/status checkpoint receives its
own final normal-push/live receipt in the thread/PR; recheck current equality.

Workspace: `D:\projects\WinBookSplit-main`; active branch/upstream:
`codex/winbooksplit-v1-m1` / `origin/codex/winbooksplit-v1-m1`.
Origin fetch/push: `https://github.com/PikkuJanne/WinBookSplit.git`.
[Draft PR #6](https://github.com/PikkuJanne/WinBookSplit/pull/6) continues M1;
keep it draft until cumulative M1-T06 review/merge. PR #5 was already merged
at `2026-10-09T03:26:06Z` to main `8c7c22fa8eda3c61c23b63d429d0b6abbd40432a`.
A normal fetch/fast-forward incorporated actual main before M1-T02. That merge
does not waive remaining M1 tasks. Retain existing branches/history, preserve
unrelated edits and reconcile fresh remote work normally.

Read AGENTS.md, STATUS.md, TASKS.json, SCOPE_AND_DECISIONS.md,
GITHUB_WORKFLOW.md, TESTING.md, SUPPORT_AND_SETUP.md,
evidence/M1-T02-manual.md/json, tasks/M1-T03.md, specs/SPLIT_CONTRACT.md
and PLAN_ORACLES.json. Start with actual inspection:

```powershell
git status --short --branch
git remote -v
git worktree list
git rev-parse --abbrev-ref --symbolic-full-name '@{upstream}'
git rev-parse '@{upstream}'
python -B .\tools\codex-handoff\check_sync.py --repo .
python -B .\tools\codex-handoff\validate_plan.py --plan-root .\docs\codex-v1.0.0
```

Do not add `-I` to the sibling-import handoff helpers. Application tests use the
explicit supported fresh developer Python `-I -B`. Inspect authentication,
live feature/main refs, current PR/protections/required checks and tags/releases
separately. The last live audit found main/feature unprotected with no rulesets,
workflows/runs/tags/releases; this is historical evidence, not current proof.
No auto-stash/reset/force push/deletion or setting changes.

## Verified behavior to preserve

`engine/winbooksplit_engine.py` exposes import-safe `plan_manual_starts`,
`split_pdf` and the existing writer. Manual starts validate every ASCII decimal
token and bound before normalization/writes; always include physical page 1,
sort/dedupe with notices and produce one whole-document PDF for `1`. Invalid
lists exit 1 with `invalid_start_pages`; zero-page manual input reports
`invalid_document`. Huge invalid and long-leading-zero tokens avoid integer
digit-limit problems. PowerShell still resolves the shipped engine from
`$PSScriptRoot`; both launchers/README/license/artwork retain M1-T01 raw bytes.

Final full gate: 36 Python tests, nine Pester tests per actual PS5.1.26100.9444/
PS7.6.5, zero skips/syntax/scaffold findings; all 22 manual CLI oracles, eight
extra CLI edge cases, seed 20261009's 250 valid plans and 25 real writer samples.
Each slice/flattened output preserves exact physical page identities/order.
Six existing GUID/marker-owned actual launcher probes include three concurrent
launches and preserve inputs/synthetic neighbors/CWD decoys/shared TEMP. Only
owned Documents outputs were inspected/removed; private entries were not read.
Two PS7 probes retain legacy RawUI cursor stderr; 60 legacy analyzer observations
per host remain unclaimed as a runtime static pass. Final tested-path digest:
`27cf0208484d831d5e818e64b1f2478a5066b96f436e2fb9da53f8029a3fe1ad`.
Clean C matched all tested raw bytes; no clean-C full rerun was claimed.

Regular GIL CPython **3.14.8 x64 only**, **pypdf 6.19.0**; dev **ReportLab
5.0.1**, **Pillow 12.3.0**, **charset-normalizer 3.5.2**. Use the support setup
for a fresh hash-required dev venv and assert import origins; do not assume old
TEMP environments remain. Exact shell tools: **Pester 6.2.0** and
**PSScriptAnalyzer 1.25.0**, all 66 prior saved file hashes/sizes reverified.
Discover actual absolute host locations; the docs' PS7 example directory may
be absent. Children use their own built-in module paths/process-only RemoteSigned;
stored policies remain unchanged and managed policy wins.

## Current test route and M1-T03 boundary

Use the central `tests/run_tests.py --layer manual` for targeted acceptance and
`--layer full` with both actual host paths/tool root for the current full route;
all reports use new external absolute paths. Full deliberately runs corrected
manual behavior, safe import and three unchanged known-bad bookmark observations.
Original/extraction harnesses/raw guards/oracles/fixture generator and historical
evidence remain unchanged. Historical AC-011 AST checks read immutable M1-T01
C `88c2149`. Explicit `--layer baseline` refuses the changed launcher;
`--layer extraction` refuses changed processing ASTs as expected. Do not claim
current code is mechanically identical to the original or weaken its guards.
The new manual receipt checker requires all promised cases/preservation/cleanup;
native Windows path equality accepts actual host casing/slash variants while
still rejecting unrequested or duplicate hosts.

Level 1 still omits front matter; Level 2 still crosses parent boundaries.
M1-T03 covers AC-019 through AC-022: destination/depth/lineage/source-order
normalization, finite integer page validation with diagnostics, duplicate/out-of-
order determinism, front matter and bounded malformed/deep outline traversal.
Reproduce each affected defect before editing. Deliberately transition only
affected current bookmark observations to corrected oracles; retain manual
regressions and immutable historical evidence. Level 2 repair belongs M1-T04;
shared validation/output/metadata tasks follow. One task per thread.

## Remaining gates

Official portable **Calibre 9.15.0** remains per-user at
`$UserProfile\Apps\Calibre915\Calibre Portable`; converter:
`$UserProfile\Apps\Calibre915\Calibre Portable\Calibre\ebook-convert.exe`.
Prior hash/Valid-signature/install/version evidence is in M1-T01. No GUI or
actual EPUB/AZW3 conversion is claimed. This trusted nonstandard path remains
outside unchanged discovery; integrate in M2-T04 and run real conversion in
M4-T04/package/release gates. Do not commit/bundle dependency binaries.

All launcher/error/discovery/output transaction paths, Explorer, renderer/fidelity,
clean OS/extracted release package, CI (M4-T05), public release/downloads remain
NOT RUN/open. Windows10/ARM/UNC/other Python stay unclaimed. No tag/release was
created. Only RELEASE_RUNBOOK.md's verified public non-draft/non-prerelease
v1.0.0, matching anonymous downloads, fixed tag and synchronized final main
complete the project. Unsigned publication remains allowed with disclosure.
