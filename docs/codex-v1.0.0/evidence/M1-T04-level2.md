# M1-T04 — Parent-aware Level 2 coverage

Observed: 9 October 2026, Europe/Berlin. Final full run began at
`2026-10-09T07:01:23.965488+02:00` (05:01:23 UTC).
Author: Codex primary; independent read-only source/harness/raw-evidence review;
separate unit-test and test-route agents.
Base: `90297a80256315de20730e44403d820d40d9d4e9`, tree
`d8472411b73f952ccf7e50dfae6d4516f2babe8d`.
Implementation C: `8c193ae2aac5df4819cad216635d59df57d12cbc`, tree
`6c45267978a41af14881115de02521ea6fb30d99`.
Tested actual-path digest:
`fca9a3b30142da14f0403a971b9d5ee02ee7e81d0cf3899ed0cc28f7609c1003`.
The [machine record](M1-T04-level2.json) holds the 38-path map, complete Level 2
observations/page identities/output hashes, regression summaries and raw receipt
hashes. Full tests preceded C; all tested actual bytes matched clean C. No
clean-C full rerun is claimed.

## Behavior and scope

Before runtime edits, four generated PDF CLI cases reproduced Level 2 defects.
BM-03 omitted front matter and parent openings, allowing A2 to absorb pages 9–10
from B. BM-04 let A2 absorb B rather than create B's intact fallback. BM-07 used
an outside child as a boundary without diagnostics. BM-06 returned the inherited
sentinel/55 without the required structured selected-level error. Eleven separate
callable mock cases captured aliases, invalid ancestors/destinations, deeper
nodes, malformed recursion and no-plan behavior. Actual PDF and mock observations
remain distinctly identified; the unchanged pre-fix source and script are hashed.

`plan_level2` reuses Level 1 normalization and parent intervals, then selects only
each retained parent's own direct depth-2 children. Children must satisfy
`parent.start <= page < parent.end` before deduplication/sorting. A child at the
parent start creates no empty opening; a child at the parent end is outside.
First-source child aliases win within their own parent and ordering/duplicates
warn. Ignored duplicate-parent subtrees and invalid-parent descendants retain
lineage and receive diagnostics; their children are never adopted by another
parent. Depth 3+ remains metadata and cannot provide Level 2 boundaries.

Front matter precedes parent sections. When the first child starts later, the
planner emits `<Parent> - Opening pages`; a parent without retained children is
one intact fallback. Every child ends at its next sibling or its own parent's
end. At least one retained usable direct child somewhere is required for the
whole Level 2 plan. Otherwise valid parents yield `no_bookmarks_at_level`, the
compatibility `[NO_BOOKMARKS_FOUND]` sentinel and exit 55 before any writer call.
No silent noninteractive fallback occurs. Empty/no usable parents retain their
distinct errors; malformed/zero-page requests fail with exit 1. Console fallback
choice/UI certification remains M3.

BM-03's exact output is `[0,2),[2,3),[3,6),[6,8),[8,10),[10,12)` with titles
Front matter, A - Opening pages, A1, A2, B - Opening pages and B1. The same
planned entries supply execution ranges/filenames. Manual, Level 1, writer and
launcher functions retain their behavior; shared plan validation/source binding
is M1-T05. No generalized framework or engine replacement was introduced.

The current manual route now has no historical defective bookmark comparisons.
Its new actual corrected BM-03 CLI reference checks all six ranges/titles/IDs
and supplies the unchanged GUID-owned launcher helper. Original/extraction
harnesses and raw guards, fixture generator/oracles and earlier evidence remain
unchanged. Both launchers, original README/license and artwork match M1-T01 blobs.

## Actual environment, commands and outcomes

Existing Windows 11 Pro 26H2 x64 workstation, build 26300.9457; no clean-OS
claim. Fresh external regular GIL CPython 3.14.8 x64 dev venv; pypdf 6.19.0,
ReportLab 5.0.1, Pillow 12.3.0 and charset-normalizer 3.5.2. Hash-required
binary-only install, exact versions/venv import origins and pip check passed.
All 66 isolated shell-tool hashes/sizes matched the prior audited manifest.
Actual PS5.1.26100.9444/PS7.6.5 hosts imported Pester 6.2.0 and
PSScriptAnalyzer 1.25.0 from absolute pins. Children used host-owned modules and
process-only RemoteSigned; stored policies remained unchanged. No elevation or
global dependency/module/policy change.

Variables below are symbolic redactions of actual trusted absolute paths.
Raw reports and synthetic PDFs remain external; no private document was read,
committed or uploaded.

| Executed command/check | Actual result |
| --- | --- |
| Fresh external venv; hash-required binary-only dev install; `$DevPython -I -B $M1T04ExternalRoot\verify_environment.py` | Exit 0; supported exact versions/import origins/pip check |
| `$DevPython -I -B $M1T04ExternalRoot\before_level2.py` before runtime edits | Exit 0; 4 real PDF CLI cases and 11 separately identified mock cases reproduced defects |
| Targeted `unittest discover ... -p test_level2.py`; external lineage-unit receipt | Final 21 tests pass; includes 5 real-PDF unusable/alias/descendant lineage cases; expected pypdf NullObject warnings retained in stderr |
| `$DevPython -I -B $Repo\tests\run_tests.py --layer level2 --report <new-external-path>` from unrelated CWD | First run exit 0: 4 CLI oracles, 6 hierarchy aggregates, 150 seeded plans, 10 real CLI writer samples |
| `$DevPython -I -B $Repo\tests\run_tests.py --layer full --tool-root $ToolRoot --shell-path $PS51 --shell-path $PS7 --report $M1T04ExternalRoot\full-final.json` from unrelated CWD | Exit 0: six stages, 85 Python tests, 9 Pester tests per actual host, zero skips |
| Both-host syntax/new-scaffold checks | Zero errors/findings; 60 inherited application observations per host remain separately limited |
| Manual and Level 1 regressions in full | Manual 22 targets/8 extras/250 plans/25 writers; Level 1 4 targets/5 aggregates/150 plans/10 writers; all exact page identities preserved |
| Existing owned actual launcher probes | 6 probes/3 concurrent; both PS7 Level 2 probes match corrected BM-03 six-section PDF bytes/IDs; other probes retain corrected manual MAN-03 |
| Plan validator and intended/staged whitespace checks | VALID: 36 tasks/107 acceptance cases/30 oracles; exit 0 |

Full raw SHA-256:
`9b581eddd16685c220cfdedc0d5be296967e69ce79660ee425ba5a11b545f2e1`.
Level 2 child SHA-256:
`84451197e3e5f3bcd3876c3375982a059407aad99b45f7efcdbef5b7b402c44d`.
All output slices and flattened sequences were checked by physical page identity
and order; seeded plans independently checked parent ownership/title selection.
Source inputs, synthetic neighbors, CWD engine decoys and the shared TEMP
sentinel remained unchanged. Cleanup removed only observed GUID/marker-owned
Documents outputs and task-owned temporary fixtures, without reading private
Documents entries. No failed T04 harness/test receipt occurred in these runs.

| Acceptance | Executed evidence |
| --- | --- |
| AC-023 | BM-03 actual CLI exact six ranges/IDs/titles; real writer and both controlled PS7 launcher references |
| AC-024 | BM-04 intact no-child parent fallback; leading/middle/final fallback unit cases; seeded ownership checks |
| AC-025 | BM-07 actual outside-child warning; parent-start equality, end exclusion, duplicate/out-of-order first-source child/parent subtree tests |
| AC-026 | BM-06 actual CLI structured no_bookmarks_at_level/sentinel/55/no files; alias-only, invalid-parent-only, outside-only and grandchild-only cases cannot satisfy selected-level gate; no writer on rejection |

## Review, Git and remaining limits

Independent read-only review verified source/harness, real-PDF lineage stderr,
raw full receipts, every source hash, recomputed digest, exact page IDs and
corrected launcher reference bytes; no remaining M1-T04 blocker. Public
evidence/privacy review accompanies E. No submitted GitHub review or CI run is
claimed. The Python unit stage includes seven expected pypdf NullObject warnings;
both PS7 piped launcher probes retain inherited RawUI invalid-handle stderr.
Those are recorded observations, not empty-stderr or full console acceptance.

Branch: `codex/winbooksplit-v1-m1`; [draft PR #8](https://github.com/PikkuJanne/WinBookSplit/pull/8).
C was normally pushed; clean local HEAD equaled the live feature at
`2026-10-09T05:03:45.266791+00:00` (07:03:45 Europe/Berlin).
PR #7 was already merged at `2026-10-09T04:52:11Z` to main `90297a8`; normal
fast-forwards reconciled feature and local main. The fresh audit found no
protections/rulesets/workflows/runs/tags/releases; no settings changed. PR #8's
later read is recorded separately from the pre-code audit timestamp.
Following documentation-only E receives its own normal-push/live receipt in the
thread/PR; neither its own SHA nor a future sync claim is embedded here.

Actual corrected Level 2 splitting through PS5.1/batch, corrected Level 1 console
paths, ordinary discovery/all launcher/error paths, full interactive fallback,
Explorer, output transactions, Calibre conversion, rendering, clean OS/extracted
package, CI and release/download gates remain open. No current Calibre probe or
conversion ran. Windows10/ARM/UNC/other Python remain unclaimed. No merge/tag/
release was created in this task. Keep PR draft through cumulative M1-T06.

Next: **M1-T05 — Unify planning, preview data and execution**. Reverify the
final checkpoint/live state, read its specs, then implement the shared plan,
coverage validator, preview/execution parity and reader/source binding.
