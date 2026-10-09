# M1-T03 — Bookmark normalization and complete Level 1 coverage

Observed: 9 October 2026, Europe/Berlin. Final full run began at
`2026-10-09T06:10:38.792098+02:00` (04:10:38 UTC).
Author: Codex primary; independent read-only source/harness/evidence review;
separate unit-test and test-route agents.
Base: `9bdc03b47fd38213784ff55b5d76a0a0120815e3`, tree
`e3c5f91a6d50d815efb5be70fa7100a0c3861bea`.
Implementation C: `a0a742ef165e751dbbbd4667962e88d7df6c4827`, tree
`b9cc8882883ef2585fb390a96725896b67c8a1f6`.
Tested actual-path digest:
`b937df04cc8035a30d01dbae9737114fe3646fc937033970c416b32e769cc468`.
The [machine record](M1-T03-level1.json) contains the 36-path map, actual
page identities/output hashes, commands, pre-fix observations and raw receipt
hashes. Full tests preceded C; every tested byte matched clean C. No clean-C
full rerun is claimed.

## Behavior and scope

Before editing, three generated PDF writer cases reproduced lost opening pages
and absent duplicate/reordering notices. Sixteen separate callable mock cases
reproduced unsafe invalid/external destinations, silent resolution failures,
zero-output success and recursion failures on cyclic/deep outlines. The raw
pre-fix receipt is hashed in the machine record; mocks are not PDF/Windows tests.

The existing shipped Python engine now normalizes source order, depth, parent
IDs and lineage with iterative traversal. Recoverable invalid/external
destinations retain metadata and receive warnings before being skipped. Valid
Level 1 starts sort physically, keep the first source title at duplicates and
report duplicate/reordering warnings. Nonempty front matter precedes complete
parent intervals. The same planned entries supply writer boundaries/filenames.
For ten pages and starts 4,7, exact output is `[0,3) Front matter`, `[3,6) A`,
`[6,10) B`, preserving every physical page exactly once.

Raw outline/name-tree preflight bounds pypdf's otherwise recursive retrieval.
Depth is capped at 64; each raw tree and named definitions at 10,000. Normalized
destination entries plus child-list containers share a 10,000-item budget.
Unreadable, cyclic, reused/contradictory and over-limit trees reject before
writes. This is bounded outline handling, not arbitrary-depth PDF support.
Empty/no usable Level 1 outlines report structured errors plus the inherited
`[NO_BOOKMARKS_FOUND]` sentinel/exit 55; malformed/zero-page documents exit 1.

Review reproduced two pypdf pitfalls: unsupported fits can synthesize page 0,
and bare numbers can be mistaken for page object IDs. Raw internal destinations
must reference page objects and use supported fit names. Real named/indirect
destinations retain the requested pages; invalid fits/object IDs warn and skip.
The review receipt explicitly marks the earlier engine probe's uncaptured hash
and timestamp as unknown, while current library/fixed-engine bytes are hashed.

Manual parsing/writing, the launchers and the Level 2 branch retain their scope.
Level 2 still has its historical parent-crossing defect; repair is M1-T04.
The original/extraction guards, baseline expectations, generator/oracles and
historical evidence are unchanged. Only current affected Level 1 observations
transition to corrected targets; current manual characterization retains BM-03.
The two launchers, original README/license and artwork match M1-T01 raw blobs.

## Actual commands, environment and outcomes

Existing Windows 11 Pro 26H2 x64 workstation, build 26300.9457; no clean-OS
claim. Fresh external regular GIL CPython 3.14.8 x64 dev venv; pypdf 6.19.0,
ReportLab 5.0.1, Pillow 12.3.0, charset-normalizer 3.5.2. Hash-required binary-only
install, pip check, exact versions and venv import origins passed. All 66
external shell-tool hashes/sizes matched the prior manifest. Actual hosts:
PS5.1.26100.9444 and PS7.6.5; Pester 6.2.0/PSScriptAnalyzer 1.25.0 imported
from absolute pins. Children used their own host modules and process-only
RemoteSigned. Stored policies remained unchanged; no elevation/global install.

Variables below redact actual trusted absolute paths. Raw receipts and all
generated PDFs remain external; no private document was processed or uploaded.

| Executed command/check | Actual result |
| --- | --- |
| `$SetupPython -I -m venv <new-external-venv>`; hash-required binary-only dev install; `$DevPython -I -B $M1T03ExternalRoot\verify_environment.py` | Exit 0; exact versions/import origins and pip check |
| `$DevPython -I -B $M1T03ExternalRoot\before_level1.py` before runtime edits | Exit 0; 3 actual generated PDF cases and 16 separately identified mocks reproduce defects |
| Targeted `unittest discover ... -p test_level1.py` | Final 26 tests pass; raw integer collision/real named/fit, bounds and no-writer rejection covered |
| `$DevPython -I -B $Repo\tests\run_tests.py --layer bookmarks --report <new-external-path>` | Final exit 0: 4 CLI oracles, 5 normalization aggregates, 150 seeded plans and 10 real writer samples |
| `$DevPython -I -B $Repo\tests\run_tests.py --layer full --tool-root $ToolRoot --shell-path $PS51 --shell-path $PS7 --report $M1T03ExternalRoot\full-final.json` from unrelated external CWD | Exit 0: five stages, 63 Python tests, 9 Pester tests per host, zero skips |
| Both-host syntax/new-scaffold analysis | Zero errors/findings; 60 inherited application findings per host remain separately recorded limitations |
| Manual regression in full run | All 22 targets, 8 extra CLI cases, 250 seeded plans, 25 real writer samples and unchanged historical BM-03 pass |
| Existing owned actual launcher probes | Six probes, including three concurrent; exact references/page IDs preserved, inputs/neighbors/CWD decoys/shared TEMP unchanged, owned Documents outputs removed |
| Plan validator; intended/staged whitespace checks | VALID: 36 tasks/107 acceptance cases/30 oracles; exit 0 |

Full raw SHA-256:
`707bf661912988db51aba408ae197c2eea2ee19f269639d4358df8bf754b43c2`.
Level 1 child SHA-256:
`7fa28ed6ac735d680e561505592dc34a3ef45bd3ebf29db01b9cf57863cf7e00`.

| Acceptance | Executed evidence |
| --- | --- |
| AC-019 | BM-01/02 actual CLI exact ranges/page IDs, front-matter title; every slice and flattened sequence checked |
| AC-020 | BM-08 actual CLI retains first A over Alias, physical order and duplicate/reordering warnings; seeded first-source selection |
| AC-021 | Invalid type/bound/resolver mocks and actual raw/named/external PDF CLI cases; invalid fits/page-object IDs excluded with warnings and full valid coverage |
| AC-022 | Depth 3/64 metadata/lineage, 65 rejection, 10,000/10,001 budgets, raw named bounds, cycles/reuse/malformed mocks; actual serialized cyclic PDF CLI fails before outputs |

Three early bookmark-harness runs failed with a bracket syntax error, missing
synthetic fixture `children`, and missing mock-reader `pages`. They were fixed
before final targeted/full runs; their separate native exits/digests/raw hashes
remain in the machine record. Earlier passes are distinct from final source.
The public-summary privacy check also refused its first external builder result
before writing because absolute-path dictionary keys needed redaction. Recursive
key redaction fixed it; the reviewed public files contain symbolic paths.

## Review, synchronization and limits

Independent read-only review verified frozen source/harness, pre-fix evidence,
raw full receipts, recomputed source digest and exact page identities; no
remaining M1-T03 blocker. Public evidence/privacy review accompanies this E.
No submitted GitHub review or CI execution is claimed.

Branch: `codex/winbooksplit-v1-m1`; [draft PR #7](https://github.com/PikkuJanne/WinBookSplit/pull/7).
C was normally pushed; clean HEAD equaled live feature C at
`2026-10-09T04:12:06.716013+00:00` (06:12:06 Europe/Berlin).
PR #6 was already merged at `2026-10-09T03:52:26Z` to base main `9bdc03b`;
normal fast-forwards reconciled feature and local main. Live audit found no
protections/rulesets/workflows/runs/tags/releases; no settings changed.
Following documentation-only E receives its own normal-push/live receipt in
the thread/PR. This record contains neither its own SHA nor a future sync claim.

The six launcher probes exercise manual MAN-03 and historical Level 2 BM-03;
they do not certify corrected Level 1 console/Explorer. Both PS7 probes retain
inherited RawUI invalid-handle stderr. All launcher/discovery/error/output
transactions, Explorer, Calibre conversion, rendering/fidelity, extracted
package/clean OS, CI and release/download gates remain open. No current Calibre
conversion/version probe ran. Unsupported Windows10/ARM/UNC/other Python remain
unclaimed. No tag/release/merge was created in this task.

Next: **M1-T04 — Make Level 2 parent-aware**. Reverify final checkpoint/live
state and read its task/spec/oracles before reproducing and fixing parent cuts.
