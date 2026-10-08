# Project status

Project: WinBookSplit first public v1.0.0
Overall: IN PROGRESS; handoff imported, application unchanged, release unverified.
Current task: M0-T01, in progress pending its synchronized implementation checkpoint.
Next task: M0-T02 after the verified M0-T01 evidence checkpoint.
Live baseline confirmed 8 October 2026: `6edbed7c1a0c94968999882c5a46d90492d3c327`; all seven original file blobs match.
Active branch: `codex/winbooksplit-v1-m0`.
Origin fetch/push: `https://github.com/PikkuJanne/WinBookSplit.git`.
Last verified code/checkpoint commit: none yet.
Live synchronization: PENDING; the feature branch has not yet been pushed.
Public release: NOT PUBLISHED; live tags/releases were empty on 8 October 2026.
Application Windows/Calibre/Explorer execution evidence: NOT RUN.

## Evidence and limitations

[M0-T01 evidence](evidence/M0-T01-handoff.md) and [live audit](evidence/M0-T01-live-audit.json) record the missing initial Git metadata, safe reconciliation, validated import and preserved file hashes. No pre-existing guidance differed. The bundle helper suite ran 57 tests: 56 passed, one symlink test skipped for unavailable host permissions. These are helper checks, not application acceptance.

pypdf is absent from the current Python 3.14.7; Calibre was not found on PATH or at current application discovery locations. Supported dependency versions and real application tests remain later M0 work. These do not block this handoff task's checks.

## Completed tasks

None until M0-T01's intended commit is pushed and freshly verified. Historical receipts do not prove current synchronization. M0 review/merge remains M0-T04.

## Finish rule

Only RELEASE_RUNBOOK.md's verified public v1.0.0 release plus final synchronized main allows COMPLETE. The user authorized final publication after all mandatory gates; no intermediate releases are permitted.
