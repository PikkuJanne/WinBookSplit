# M1-T05 — Shared plan, preview and execution

Observed: 9 October 2026, Europe/Berlin. Final full run began at
`2026-10-09T07:42:45.913161+02:00` (05:42:45 UTC).
Author: Codex primary; independent read-only source/harness/raw/public-evidence
review, with separate unit-test and test-route agents.
Base: `828b232b6a25b2c67c65fc3e31d2f3f7cb30a38c`, tree
`5691bd09234822fd96ff944f4810671fb6b3172b`.
Implementation C: `3b3e97ec0551e9ed1ac72db5bcf98fe67ca1daaf`, tree `450ee7db1cd2330643cf2fd431cb7855f9c489a9`.
Tested actual-path digest: `73f7c3848cf35dcd3444cdd5af861231e81efe775d9afff298eac22abc2defae`.
The [machine record](M1-T05-plan.json) includes the complete 41-path map,
actual preview/output/page identities/content hashes, source-change observations,
runtime/setup/policy records, regressions and raw receipt hashes. Full tests
preceded C; every tested actual byte matched clean C. No clean-C full rerun is
claimed. The following documentation checkpoint receives its separate final
normal-push/live receipt in the thread/PR; no future/self SHA is embedded here.

## Changes and reproduction

Before engine changes, six injected logical-planner cases used real writers to
show that manual, Level 1 and Level 2 each accepted a gap (physical page 4 lost)
and an empty plan (normal return with zero outputs). A separate old callable
Level 1 planning observation followed by source replacement produced different
filenames/ranges at execution. That observation was not a shipped console UI.
Two further actual owned-file probes showed source/output aliasing changed the
source bytes and an existing planned chapter sentinel was overwritten. These
mutable synthetic copies are distinguished from preserved original fixtures.

The existing mode planners and normalization stay unchanged. A common
`validate_plan` rejects empty/noninteger/out-of-range/reversed ranges, gaps,
overlaps, missing first/tail pages, inconsistent declared ranges, incomplete
entry metadata, nonconsecutive sequences and unsafe/duplicate basenames. It
derives complete coverage/ranges from entries and freezes detached nested data.
`prepare_split` probes once and binds that plan to pypdf's in-memory reader;
`preview_plan` exposes the same read-only data without chapter writes;
`execute_split` consumes the exact entries without reopening or replanning.
The immutable result reports the same entries/source/coverage and page counts.

Source identity records requested/resolved PDF path, SHA-256 and captured size.
Changed, repointed or deleted filesystem paths safely retain the original reader
snapshot. Replacing the plan/reader through ordinary dataclass substitution or
altering the captured stream rejects before writing. Private Python memory/page
cache manipulation is outside the supported API; this is not a sandbox.
All planned outputs are checked for source aliases and existing files before
any slice; the existing writer now uses exclusive creation to refuse a later
collision. This is narrow immutable-input/plan safety, not complete output
staging/rollback/publication. Console preview/fallback integration remains M3;
diagnostic categories/M1 cumulative review remain M1-T06.

## Environment and executed checks

Existing Windows 11 Pro 26H2 x64 workstation, build 26300.9457; not a clean OS.
Fresh external regular GIL CPython 3.14.8 x64 venv: pypdf 6.19.0, ReportLab
5.0.1, Pillow 12.3.0, charset-normalizer 3.5.2. Hash-required binary-only install,
exact versions/import origins and pip check passed. All 66 isolated shell-tool
hashes/sizes matched M0-T04's saved manifest. Actual PS5.1.26100.9444/PS7.6.5
imported Pester 6.2.0/PSScriptAnalyzer 1.25.0 from absolute pins. Child process-only
RemoteSigned and host-owned modules left stored policies unchanged. No elevation
or global dependency/module/policy changes. No current Calibre probe/conversion.

Variables below redact actual trusted absolute paths. Raw artifacts and PDFs
remain external; all input content is authored/generated synthetic data.

| Command/check | Actual result |
| --- | --- |
| Fresh external venv; hash-required binary-only dev install; `$DevPython -I -B $M1T05ExternalRoot\verify_environment.py` | Correct capture exit 0, exact runtime/import origins/pip check |
| `$DevPython -I -B $M1T05ExternalRoot\before_plan.py` before changes | Exit 0: six real-writer injected planner defects plus old callable preview/source mismatch |
| `$DevPython -I -B $M1T05ExternalRoot\before_plan_output_safety.py` before changes | Exit 0: owned source alias and existing chapter overwritten |
| `$DevPython -I -B $Repo\tests\python\test_plan.py`; stable external capture | 18 tests pass/native exit 0; before/after engine/test/fixture hashes equal |
| Targeted `tests/run_tests.py --layer plan --report <new-external-path>` from unrelated CWD | Four passing receipts; final 04 matches final digest; nine structural/two preview/three real mode/two source aggregates plus deleted-path execution/300 seeded plans |
| Runner unittest checks in final full | 20 tests pass; malformed promised evidence must fail |
| `$DevPython -I -B $Repo\tests\run_tests.py --layer full --tool-root $ToolRoot --shell-path $PS51 --shell-path $PS7 --report $M1T05ExternalRoot\full-final.json` from unrelated CWD | Final exit 0, seven stages, 105 Python tests and nine Pester tests per host, zero skips |
| Both-host syntax/new-scaffold checks | Zero errors/findings; 60 inherited application observations per host remain separate limits |
| Manual/Level 1/Level 2 regression layers | 22/4/4 CLI targets; 8/5/6 aggregates; 250/150/150 seeded plans; 25/10/10 real writer samples; exact complete page IDs/order |
| Existing actual owned entrypoints | Six probes including three concurrent; manual and corrected six-range BM-03 bytes/IDs match; sources/neighbors/decoys/shared TEMP preserved and owned outputs removed |
| Plan validator and intended/staged whitespace checks | VALID 36 tasks/107 acceptance cases/30 oracles; exit 0 |

Full raw SHA-256: `a95dd06cdcacec376bd87a7427ad792fd32ea58b8f58e3324e62eb3d923e301c`.
Shared-plan child SHA-256: `3eddbae9f7cf1f9059b36f6995f35f1d7a43c602fd83273d5ed7593bcb4022ac`.
Every new real slice is reopened and compared against preview filename/range/
order and physical IDs. Full extracted-page content hashes match original pages,
including after source replacement; replacement/moved bytes remain unchanged.
Preparation/preview tree preservation proves no chapter writes. Execution is
patched to fail if it reopens or replans. Seed 20261009 yields 300 independently
expected complete partitions, 100 per mode; the 18 units add 100 seeded cases.

| Acceptance | Executed evidence |
| --- | --- |
| AC-027 | Nine structural aggregates, type/metadata/sequence/cache defects, 300 seeded all-mode plans and 100 unit partitions; six injected invalid production-mode plans reject before writer |
| AC-028 | Manual/Level 1/Level 2 actual prepared writers match exact preview entries, filenames/ranges/order/IDs/content; deeply immutable detached preview/result and no preview artifacts |
| AC-029 | In-place replacement, rename/repoint and deletion retain original captured source/content; units also use same-page-count reversed content and reject private stream changes/cross-job substitution |

## Earlier observations, review and limits

The first environment capture failed because root redirection precreated the
verifier's exclusively created JSON path; its log is retained. Correct capture
passed without changing dependencies. An earlier 16-test native run passed but
its receipt consistency guard failed as production changed concurrently; it is
not final-source acceptance. The stable 18-test receipt supersedes it. The first
full run passed seven stages/105 tests at digest `a73aab92...`, then review made
ordinary page-content evidence mandatory in the report checker. Its receipt is
preserved separately; the corrected final full uses the digest above. Four
retained targeted plan runs passed. No production test failure is
hidden by these observations.

Independent source/harness/raw/public-evidence review found no remaining blocker.
Seven expected pypdf NullObject unit warnings, first-use PS5.1 CLIXML progress,
both PS7 launcher RawUI invalid-handle stderr and legacy analyzer observations
are retained, not empty-stderr/full console acceptance. No submitted GitHub
review or CI run is claimed. Historical guards/oracles/generator/earlier evidence
are unchanged; launchers/README/license/artwork still match M1-T01 blobs.

Branch `codex/winbooksplit-v1-m1`; [draft PR #9](https://github.com/PikkuJanne/WinBookSplit/pull/9).
C normal push/clean live equality was observed at `2026-10-09T05:45:17.796145+00:00`
(07:45:17 Europe/Berlin). PR #8 had already merged at `2026-10-09T05:22:09Z`
to main `828b232`; normal fast-forwards reconciled feature/local main. Live audit
found no protections/rulesets/workflows/runs/tags/releases; settings unchanged.
PR #9's later timestamped read is separate from that pre-code audit.

Ordinary discovery/all launcher/error paths, actual corrected Level 2 PS5.1/BAT
and Level 1 console, full interactive preview/fallback/Explorer, Calibre,
rendered fidelity/PDF feature policy, output transactions, clean OS/extracted
package, CI and final release/download remain open. Windows10/ARM/UNC/other
Python remain unclaimed. No merge/tag/release created here; keep PR draft through
cumulative M1-T06. Next: **M1-T06 — Distinguish no-outline and invalid-document results**.
