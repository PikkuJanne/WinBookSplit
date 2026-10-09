# M1-T06 — Structured diagnostics and reviewed M1 checkpoint

Task: M1-T06; AC-030 and AC-031 **PASSED**.
Date: 2026-10-09T07:14:25.174834+00:00 UTC; user timezone Europe/Berlin.
Author: Codex. Independent local read-only cumulative/source/harness/receipt review passed;
no submitted GitHub approval or human interaction pass is claimed.
Base: `aada4e5b7c018831d6f0971f570957f22931141a`, after normally reconciling externally merged PR #9.
Implementation C: `dee49ac233650e12511a3597077523527de76c5e`; tree `f1dae1f084b7f25c7d3d3c5d863c7c71a3bdff5e`.
Tested actual 46-path SHA-256 map digest: `26d476f3a99370727834ca97f6365f03de68e18bc755de31c7d5c8ed0264c625`.
Full raw receipt SHA-256: `e2f15465918b9fc079950cef2ff5234b9a200a7c659287956672560c796f4f58`.
[Machine evidence](M1-T06-diagnostics.json) contains exact paths, commands, identities,
case/category/choice results, prior failures and observed synchronization receipts.

## Changes and scope

The existing manual/Level 1/Level 2 planners, captured-source plan and page writer
remain intact. A frozen structured outcome separates no outlines, unusable Level 1
destinations, absent usable direct Level 2 bookmarks, zero-page/invalid input,
unreadable input and output failures. Each CLI invocation emits one versioned JSON
result instead of the hidden sentinel. Native compatibility codes remain 0/1/55
until M3 implements the final CLI contract.

The actual PowerShell handler validates protocol/mode/native-exit agreement and
positive successful outputs. Only applicable explicit manual/Level 1 retries are
offered; pending/invalid/cancel decisions never execute a fallback. Failed retries
and cancelled no-plan runs remain nonzero and never print Done. Both process
streams drain safely, native quoting preserves literal user data, and CLI streams
use UTF-8. BAT remains byte-identical; its final native call preserves the code.

Pre-change evidence recorded seven actual engine/API/CLI cases and native argument
loss under both shells. Quoted invalid manual text could become valid page 1 or
extra ignored arguments. Exact quoting plus CLI argument-count rejection closes
that bypass. No private documents were used or committed.

## Actual commands and outcomes

All commands used explicit supported Python 3.14.8 regular-GIL x64 in a fresh
hash-required developer venv, pypdf 6.19.0, ReportLab 5.0.1, Pillow 12.3.0 and
charset-normalizer 3.5.2. Exact imports and pip check passed. Windows 11 Pro 26H2,
build 26300.9457; actual PS5.1.26100.9444 and PS7.6.5; Pester 6.2.0/PSScriptAnalyzer
1.25.0 with all 66 recorded module-file sizes/hashes reverified. Stored execution
policies were unchanged. Calibre was not exercised in this task.

| Executed command / scope | Actual outcome |
| --- | --- |
| `$DevPython -I -B $Repo\tests\run_tests.py --layer full --tool-root $ToolRoot --shell-path $PS51 --shell-path $PS7 --report $ExternalRoot\full-final.json`, unrelated external CWD | Exit 0; eight stages; 118 Python tests, zero skips; nine Pester tests per actual host; zero syntax/scaffold findings. |
| Final `--layer diagnostics` with both actual hosts | Exit 0; 21 actual CLI requests plus two callable injections; three writer controls reopen exact physical IDs/order/ranges/filenames, including Chinese/å path and titles. |
| Real handler per host | 21 category records × 11 explicit choices; 18 malformed/missing/conflicting protocol rejections; no automatic execution. |
| Native argument and AST-isolated actual `Run-PythonSplitter` per host | Literal empty/space/quote/slash/Unicode/shell-like data preserved; both 200K streams and final failure drained/logged. |
| Earlier targeted diagnostics and final full unit subsets | 12 targeted diagnostics passed on their recorded earlier source snapshot; 21 runner units and all 12 diagnostics passed in the final full gate, which covers final bytes. |
| Actual full PS5.1/PS7/BAT controlled failure probes, repeated after UTF-8 change | 12/12 passed: each entrypoint cancelled no-plan = 55, invalid manual retry = 1, corrupt = 1, quoted invalid manual = 1; no chapters/Done. |
| Existing M1 regression routes | Manual 22 targets + eight edge cases, Level 1 four targets + five aggregates, Level 2 four targets + six aggregates; 250/150/150 seeded plans and 25/10/10 writer samples. Shared plan nine structural/two preview/three real writer modes/two source aggregates plus deletion, 300 seeded plans. |
| Existing actual successful launchers | Six bounded launches, including three concurrent; manual and corrected two PS7 Level 2 BM-03 probes preserve complete output bytes/IDs. |

Every final run preserved actual source/input/neighbor hashes and verified owned
cleanup. Historical original/extraction guards, fixtures/oracles and earlier
evidence stay unchanged. Legacy application analyzer findings remain observations;
they are not a runtime static pass. PS7 RawUI stderr from redirected Clear-Host,
dependency diagnostics and any first-use host progress are retained in raw receipts.
Redirected Read-Host does not expose prompt text; decision logs and result/message
checks prove only the documented controlled flow.

Earlier development failures are retained: two checkpoint-tool invocation mistakes,
incorrect assumed PS7 paths in external preflights, a reviewer PS5.1 root-array
capture failure, mid-edit old-sentinel/invalid-plan/runner assertions and a temporary
unnecessary BAT change that tripped the preserved guard. BAT was restored to its
original bytes. The new harness initially selected the wrong no-Level2 fixture and
serialized a PS5.1 root array with ETS metadata; final fixtures/clean array capture
passed. A 12-case launcher receipt initially required unavailable redirected prompt
text; meaningful category/message/decision assertions passed on fresh owned runs.
These earlier failed receipts do not count as acceptance passes.
Public evidence review also caught escaped local paths inside repr/JSON strings
and an overbroad targeted-unit label. Both were corrected before staging; the
serialized record now asserts that concrete profile, Documents and root paths
are absent. The final full remains the evidence for final source bytes.

## Limits and next work

No-plan/invalid-plan rejections and observed controlled first-write failures created
no new chapters. A later writer failure may still leave partial files; failure
written_count counts confirmed successful records, not a rollback guarantee.
Transactional staging/promotion/cleanup is **M2-T01**. Existing early input,
dependency and conversion exits, ordinary discovery, comprehensive stream/
filesystem/cancellation checks, complete interactive UX, Explorer, actual Calibre,
PDF feature/fidelity gates, release package and CI remain open later tasks.
Human checks: NOT RUN. No required future gate is waived. No tag/release published.

## Git, review and checkpoint

C pushed normally and observed clean/live SYNCED at `2026-10-09T07:21:51.780643+00:00`.
[PR #10](https://github.com/PikkuJanne/WinBookSplit/pull/10) merged normally at `2026-10-09T07:23:47Z` into
main `24a447c01ac590ceb7dcc456ae580cea6dd47e6a`, after independent local cumulative M1 review and fresh repository gate
inspection. There were no required checks, branch protection/rulesets or workflows;
no CI run, submitted approval or protection bypass is claimed.
Merged main was clean/live SYNCED at `2026-10-09T07:23:51.802430+00:00`; every tested actual
source path still matched. Full tests preceded C; no clean-C full rerun is claimed.

This following documentation-only E records already observed C/merge/main facts.
Its own normal push/live receipt stays in the thread/PR, outside its own contents.
Only intended source/test/documentation paths were staged; raw reports, synthetic
PDFs, tools and venv stayed outside Git.

Next: **M2-T01 — Isolate runs and stage validated output**, after fresh live check.
