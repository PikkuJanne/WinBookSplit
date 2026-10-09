# Next session

Next task: **M1-T04 — Make Level 2 parent-aware**.
M0-T01 through M0-T04 and M1-T01 through M1-T03 are done. Implementation C
`a0a742ef165e751dbbbd4667962e88d7df6c4827` was normally pushed; clean local
HEAD matched live feature at `2026-10-09T04:12:06.716013+00:00`
(06:12:06 Europe/Berlin). The following evidence/status checkpoint receives its
own normal-push/live receipt in the thread/PR; recheck current equality.

Workspace: `D:\projects\WinBookSplit-main`; branch/upstream:
`codex/winbooksplit-v1-m1` / `origin/codex/winbooksplit-v1-m1`.
Origin fetch/push: `https://github.com/PikkuJanne/WinBookSplit.git`.
[Draft PR #7](https://github.com/PikkuJanne/WinBookSplit/pull/7) continues M1;
keep draft through cumulative M1-T06 review/merge. PR #6 was already merged
at `2026-10-09T03:52:26Z` to main `9bdc03b47fd38213784ff55b5d76a0a0120815e3`.
Normal fast-forwards reconciled feature and local main. Earlier merges do not
waive later tasks. Preserve unrelated edits/history and reconcile fresh work.

Read AGENTS.md, STATUS.md, TASKS.json, SCOPE_AND_DECISIONS.md,
GITHUB_WORKFLOW.md, TESTING.md, SUPPORT_AND_SETUP.md,
evidence/M1-T03-level1.md/json, tasks/M1-T04.md, specs/SPLIT_CONTRACT.md
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

Do not add `-I` to sibling-import handoff helpers. Application tests use the
explicit supported fresh developer Python `-I -B`. Inspect authentication/live
feature/main refs/PR/protections/checks/tags/releases independently. The last
audit found no protections/rulesets/workflows/runs/tags/releases; that is
historical evidence. No reset/force push/auto-stash/deletion/settings changes.

## Preserve the verified behavior

The import-safe shipped engine exposes `plan_manual_starts`, `normalize_outline`,
`normalize_bookmarks`, `plan_level1`, `split_pdf` and the existing writer.
Manual starts validate all ASCII tokens/bounds before writing, preserve page 1,
sort/dedupe with notices, produce one whole-document file for `1`, and reject
invalid/zero-page requests. Huge decimal/leading-zero tokens stay bounded.

Level 1 normalization retains source order/depth/parent IDs/lineage and unusable
records. Recoverable invalid/external destinations warn and skip. Raw page
references and supported fits are checked before trusting pypdf's normalized
pages; named/indirect valid destinations retain their pages. First-source
duplicate title wins, physical sorting warns, front matter and parent intervals
cover every page. Invalid parents remain in metadata so future Level 2 cannot
orphan their descendants. Depth 3+ is retained, never a Level 1 boundary.

Raw outline/name-tree preflight runs before recursive pypdf retrieval; unreadable,
cyclic/reused/over-limit trees reject before writing. Limits: depth 64; 10,000
raw nodes per tree/named definitions; 10,000 normalized entries PLUS child-list
containers. Empty/no usable Level 1 emits structured error plus compatibility
sentinel/exit 55; malformed/zero-page requests exit 1. These are bounded handling
semantics, not arbitrary-depth support. Both launchers remain unchanged.

Final full gate: 63 Python tests; nine Pester tests per actual PS5.1.26100.9444/
PS7.6.5; zero skips/syntax/scaffold findings. Four corrected Level 1 CLI targets,
five normalization aggregates, 150 seeded plans/ten writer samples check exact
slice/flattened physical IDs. Manual: 22 targets/eight extras/250 plans/25 writer
samples. Six existing GUID-owned actual launchers include three concurrent
invocations, preserve inputs/neighbors/decoys/shared TEMP and remove owned
Documents outputs without reading private entries. These launchers cover
manual MAN-03 and historical Level 2 BM-03, not corrected Level 1 console paths.
Both PS7 probes retain RawUI stderr; 60 legacy analyzer findings per host remain.
Tested-path digest:
`b937df04cc8035a30d01dbae9737114fe3646fc937033970c416b32e769cc468`.
Clean C matched every tested actual byte; no clean-C full rerun claimed.

## Test route and M1-T04 boundary

Use `tests/run_tests.py --layer bookmarks` for targeted Level 1 regression and
`--layer full` with both actual hosts/tool root for the five-stage full route.
Use a fresh external report path and unrelated CWD. Bookmark route takes no
shell arguments; full/manual uses six controlled owned launcher probes. Current
manual route retains only unchanged historical BM-03. Original/extraction
harnesses/raw guards/oracles/generator/evidence are unchanged. Historical
AC-011 reads immutable M1-T01 C `88c2149`; current safe import stays checked.
Explicit baseline/extraction diagnostics deliberately refuse changed source.
Do not weaken historical guards or call current code mechanically identical.

Level 2 still has its old recursive flattening/bare exception and parent crossing.
M1-T04 covers AC-023 through AC-026: use normalized parent intervals and each
parent's own direct usable children; add opening sections and whole-parent
fallback, warn/ignore outside children, retain front matter, avoid alias subtree
orphaning and define structured `no_bookmarks_at_level`. Reproduce affected
defects first. Transition BM-03 and other Level 2 targets deliberately while
retaining all manual/Level 1 regressions and immutable historical evidence.
Shared validation/output/metadata tasks follow. One thread-sized task at a time.

## Runtime and remaining gates

Regular GIL CPython **3.14.8 x64 only**, pypdf **6.19.0**; dev ReportLab 5.0.1/
Pillow 12.3.0/charset-normalizer 3.5.2. Create a fresh hash-required dev venv and
assert import origins; do not assume old TEMP environments persist. Pester 6.2.0/
PSScriptAnalyzer 1.25.0 are isolated external pins; reverify all 66 prior files.
Discover actual absolute host paths; PS7 example directory may be absent.
Children request process-only RemoteSigned with host-owned module paths;
stored policies remain unchanged and managed policy wins.

Official portable Calibre 9.15.0 remains per-user at
`$UserProfile\Apps\Calibre915\Calibre Portable`; converter:
`$UserProfile\Apps\Calibre915\Calibre Portable\Calibre\ebook-convert.exe`.
Prior hash/Valid-signature/install/version evidence is M1-T01. No actual ebook
conversion/GUI is claimed. Integrate the trusted nonstandard path in M2-T04;
real EPUB/AZW3 conversion remains M4-T04/package/release work. Do not bundle
dependency binaries. Full launcher/discovery/error/output transactions,
corrected Level 1 console, Explorer, rendering, clean OS/extracted package,
CI (M4-T05), release/download gates remain open; Windows10/ARM/UNC/other Python
remain unclaimed. No tag/release created. Only the verified public final
v1.0.0/package/download/fixed-tag/synchronized-main runbook closes the project.
Unsigned publication remains allowed with disclosure.
