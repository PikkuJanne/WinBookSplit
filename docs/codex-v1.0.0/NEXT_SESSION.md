# Next session

Next task: **M1-T05 — Unify planning, preview data and execution**.
M0-T01 through M0-T04 and M1-T01 through M1-T04 are done. Implementation C
`8c193ae2aac5df4819cad216635d59df57d12cbc` was normally pushed; clean local
HEAD matched live feature at `2026-10-09T05:03:45.266791+00:00`
(07:03:45 Europe/Berlin). Following evidence/status checkpoint receives its own
normal-push/live receipt in the thread/PR; recheck current equality.

Workspace: `D:\projects\WinBookSplit-main`; branch/upstream:
`codex/winbooksplit-v1-m1` / `origin/codex/winbooksplit-v1-m1`.
Origin fetch/push: `https://github.com/PikkuJanne/WinBookSplit.git`.
[Draft PR #8](https://github.com/PikkuJanne/WinBookSplit/pull/8) continues M1;
keep draft through cumulative M1-T06. PR #7 was already merged at
`2026-10-09T04:52:11Z` to main `90297a80256315de20730e44403d820d40d9d4e9`.
Normal fast-forwards reconciled feature/local main. Earlier merges do not waive
later gates. Preserve unrelated edits/history and reconcile fresh work normally.

Read AGENTS.md, STATUS.md, TASKS.json, SCOPE_AND_DECISIONS.md,
GITHUB_WORKFLOW.md, TESTING.md, SUPPORT_AND_SETUP.md,
evidence/M1-T04-level2.md/json, tasks/M1-T05.md, specs/SPLIT_CONTRACT.md,
specs/OUTPUT_AND_CONVERSION.md and PLAN_ORACLES.json. Start with actual inspection:

```powershell
git status --short --branch
git remote -v
git worktree list
git rev-parse --abbrev-ref --symbolic-full-name '@{upstream}'
git rev-parse '@{upstream}'
python -B .\tools\codex-handoff\check_sync.py --repo .
python -B .\tools\codex-handoff\validate_plan.py --plan-root .\docs\codex-v1.0.0
```

Do not add `-I` to sibling-import handoff helpers. Application tests use explicit
fresh supported developer Python `-I -B`. Inspect authentication/live feature/
main/PR/protections/checks/tags/releases separately. Last audit found no
protections/rulesets/workflows/runs/tags/releases; historical evidence is not
current proof. No reset/force push/auto-stash/deletion or settings changes.

## Verified behavior to preserve

The import-safe shipped engine exposes `plan_manual_starts`, `normalize_outline`,
`normalize_bookmarks`, `plan_level1`, `plan_level2`, `split_pdf` and the existing
writer. Manual starts validate every ASCII token/bound before normalization or
writes, preserve page 1, sort/dedupe with notices, handle huge/leading-zero tokens,
produce one whole-document file for `1` and reject invalid/zero-page requests.

Level 1 normalization retains source order/depth/parent IDs/lineage, including
unusable records. Invalid/external destinations warn/skip; raw internal page
references/fits are validated before trusting pypdf pages. Valid named/indirect
destinations retain correct pages. First-source duplicate parents win; physical
sorting warns; front matter plus parent intervals cover every page. Raw outline/
name trees and iterative normalization reject unreadable/cyclic/reused/over-limit
structures before writing. Limits: depth 64; 10,000 raw nodes per tree/named
definitions; 10,000 normalized entries plus child-list containers.

Level 2 calls Level 1 once and groups only depth 2 children by retained parent ID.
Filter to parent.start <= page < parent.end before dedupe/sort; first valid child
alias wins. Parent aliases retain first subtree; invalid-parent descendants are
warned/ignored, never adopted. Depth 3+ is metadata. Add front matter and parent
opening/whole-parent fallback entries; each child ends at next sibling/parent end.
At least one retained usable direct child is required globally; otherwise
no_bookmarks_at_level/sentinel/exit 55 before writer, with no noninteractive fallback.
Empty/no usable parents retain distinct errors; malformed/zero-page fails exit 1.
Interactive fallback choice/UI remains M3. Entry parent_id/bookmark_id/reason
and sequential title-derived filenames are verified by actual identities.

Final full gate: 85 Python tests; nine Pester tests per actual PS5.1.26100.9444/
PS7.6.5; zero skips/syntax/scaffold findings. Manual: 22 targets/8 extras/250 plans/25 writers;
Level 1: 4 targets/5 aggregates/150 plans/10 writers; Level 2: 4 targets/6 aggregates/
150 plans/10 writers. Every output/flattened sequence preserves physical IDs and
selected-parent ownership. Additional real-PDF unit lineage cases pass; seven
pypdf NullObject warnings remain captured. Six actual owned launchers include
three concurrent, preserve sources/neighbors/decoys/shared TEMP and remove only
owned Documents outputs without reading private entries. Both PS7 Level 2 BM-03
probes match corrected six-section PDF bytes/IDs. Other probes cover manual;
actual corrected Level 2 PS5.1/batch and Level 1 console remain unclaimed. Both
PS7 probes retain RawUI stderr; 60 legacy analyzer observations per host remain.
Tested-path digest:
`fca9a3b30142da14f0403a971b9d5ee02ee7e81d0cf3899ed0cc28f7609c1003`.
All 38 actual tested paths matched clean C; no clean-C full rerun was claimed.

## Current test route and M1-T05 boundary

Use `tests/run_tests.py --layer level2`, `--layer bookmarks` or `--layer manual`
for focused checks. `--layer full` now has six stages and requires both actual
hosts and isolated tool root. Always use a fresh absolute external report and
unrelated CWD. Bookmark/level2 routes take no shell args. Manual historical
bookmark comparisons are empty; its corrected actual BM-03 CLI reference
supplies existing six owned launcher probes. Immutable original/extraction
harnesses/guards/oracles/generator/older evidence are unchanged. Historical
AC-011 reads M1-T01 C `88c2149`; current safe import remains checked. Explicit
baseline/extraction diagnostics deliberately refuse changed source. Do not
weaken historical guards or claim current code equals known-bad original logic.

M1-T05 covers AC-027 through AC-029: one small shared validated plan/result,
nonempty/in-range/contiguous/complete ordered coverage, read-only preview data,
execution of that same plan, and a reader-bound or checked source identity.
Test gaps/overlaps/reversal/overflow/empty and seeded invariants, exact preview-
filename/range/page identity parity and controlled source changes/repointing.
Keep the existing app/page writer/Python; no new framework. Source binding must
not silently execute a stale plan on changed data. UI binding remains M3 and
output transactions have later task ownership. Preserve all manual/Level 1/Level 2
regressions and historical receipts. One thread-sized task at a time.

## Runtime and remaining gates

Regular GIL CPython **3.14.8 x64 only**, pypdf **6.19.0**; dev ReportLab 5.0.1/
Pillow 12.3.0/charset-normalizer 3.5.2. Fresh hash-required dev venv/exact origins;
do not assume old TEMP persists. Pester 6.2.0/PSScriptAnalyzer 1.25.0 are external
pins; reverify all 66 prior file hashes/sizes. Discover actual absolute shell
locations; PS7 example directory may be absent. Children request process-only
RemoteSigned and host-owned modules; stored policies stay unchanged and managed
policy wins.

Official portable Calibre 9.15.0 remains per-user at
`$UserProfile\Apps\Calibre915\Calibre Portable`; converter:
`$UserProfile\Apps\Calibre915\Calibre Portable\Calibre\ebook-convert.exe`.
Hash/Valid-signature/install/version evidence is M1-T01, not a current conversion
pass. Integrate the trusted nonstandard path in M2-T04; real EPUB/AZW3 conversion
is M4-T04/package/release work. No dependency binaries may be bundled.
Shared plan/source invariants, full launcher/discovery/error/output transactions,
interactive fallback/Explorer, renderer/fidelity, clean OS/extracted package,
CI (M4-T05), public release/download gates remain open. Windows10/ARM/UNC/other
Python stay unclaimed. No tag/release. Only the verified public final v1.0.0/
package/download/fixed-tag/synchronized-main runbook closes the project.
Unsigned publication remains allowed with disclosure.
