# Project status

Project: WinBookSplit first public v1.0.0
Overall: IN PROGRESS; handoff imported, application unchanged, release unverified.
Completed task: M0-T01; next dependency-ready task M0-T02 after fresh verification of the evidence checkpoint.
Next task: M0-T02 after the verified M0-T01 evidence checkpoint.
Live baseline confirmed 8 October 2026: `6edbed7c1a0c94968999882c5a46d90492d3c327`; all seven original file blobs match.
Active branch: `codex/winbooksplit-v1-m0`.
Origin fetch/push: `https://github.com/PikkuJanne/WinBookSplit.git`.
Last verified implementation commit C: `ac09c024589158e1402ae63263818466d3e77d5c`.
Historical C synchronization: SYNCED at `2026-10-08T16:43:50.035807+00:00` (18:43:50 Europe/Berlin), clean local HEAD equals live branch. Evidence checkpoint E receives a separate final receipt in the thread/[PR #1](https://github.com/PikkuJanne/WinBookSplit/pull/1); recheck live equality next thread.
Public release: NOT PUBLISHED; live tags/releases were empty on 8 October 2026.
Application Windows/Calibre/Explorer execution evidence: NOT RUN.

## Evidence and limitations

[M0-T01 evidence](evidence/M0-T01-handoff.md) and [live audit](evidence/M0-T01-live-audit.json) record the missing initial Git metadata, safe reconciliation, validated import and preserved file hashes. No pre-existing guidance differed. The bundle helper suite ran 57 tests: 56 passed, one symlink test skipped for unavailable host permissions. These are helper checks, not application acceptance.

pypdf is absent from the current Python 3.14.7; Calibre was not found on PATH or at current application discovery locations. Supported dependency versions and real application tests remain later M0 work. These do not block this handoff task's checks.

## Completed tasks

M0-T01: live reconciliation, safe handoff import and verified feature-branch code checkpoint. TASKS.json references C and actual evidence. This evidence update must receive its own normal push/live verification; the final receipt lives outside its own commit. Historical receipts do not prove current synchronization. M0 review/merge remains M0-T04; PR #1 stays draft until that gate.

## Finish rule

Only RELEASE_RUNBOOK.md's verified public v1.0.0 release plus final synchronized main allows COMPLETE. The user authorized final publication after all mandatory gates; no intermediate releases are permitted.
