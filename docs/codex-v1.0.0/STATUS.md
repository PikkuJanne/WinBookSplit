# Project status

Project: WinBookSplit first public v1.0.0.
Overall: IN PROGRESS; application unchanged, original defects reproduced, release unverified.
Completed tasks: M0-T01 and M0-T02.
Next dependency-ready task: **M0-T03 — Freeze support and dependency decisions**, after fresh verification of this evidence checkpoint.
Active branch: `codex/winbooksplit-v1-m0`.
Origin fetch/push: `https://github.com/PikkuJanne/WinBookSplit.git`.
Last verified implementation C: `8f1f556f318be488839e839830264fa2e5f4fbc8`.
Historical C synchronization: SYNCED at `2026-10-08T17:00:16.304703+00:00` (19:00:16 Europe/Berlin), clean local HEAD equals live feature branch.
This documentation/evidence checkpoint receives its separate final push/live receipt in the thread/[draft PR #2](https://github.com/PikkuJanne/WinBookSplit/pull/2); recheck current equality next thread.
Public release: NOT PUBLISHED; live tags/releases were empty on 8 October 2026.

## M0-T02 evidence and limits

[Baseline evidence](evidence/M0-T02-baseline.md) and [machine report](evidence/M0-T02-characterization.json) record 14 original embedded-engine cases and two actual unchanged Windows launcher error paths. The engine was executed with generated page-marked PDFs; each output was reopened and checked by exact page identity/order. Manual first-page loss, success with zero files, invalid-token filtering, bookmark front-matter omission and Level 2 parent crossing were reproduced. Known-bad observations are separate from the corrected `PLAN_ORACLES.json` targets. This completes AC-004/005/006's reproduction/provenance obligations; it does not pass fixed-engine acceptance.

The original synthetic 10/12-page fixtures reproduce byte-for-byte with the same source/dependencies. Full anonymous fixed metadata, source/license hashes and exact outlines are recorded. All 22 pages were rendered and visually inspected. Original source, inputs and sentinel neighbors remained unchanged. No real documents, generated PDFs or private artifacts were committed.

Actual fixture/engine environment: existing Windows 11 workstation build 26300.9457; explicit bundled isolated Python 3.12.14, pypdf 6.10.0, ReportLab 4.4.9. Orchestrating shell was PowerShell 7.6.5; the bounded launcher probe used actual Windows PowerShell 5.1.26100.9444. Poppler was 26.07.0. These are development observations; supported versions remain M0-T03's decision. No global dependency installation occurred.

Full Windows PowerShell 5.1/7 splitting workflow, human Explorer drag/drop, Calibre conversion, release-package checks and fixed-engine acceptance: **NOT RUN**. Independent read-only source/evidence review found no blocker. No GitHub review submission or CI pass is claimed. M0-T04's cumulative review/test-scaffolding/merge gate remains open.

## Repository reconciliation and earlier work

At this thread's start, clean local/live feature checkpoint `4f550a2cdfb9736be8c2fd35ab83cc65c1f61ffc` was freshly SYNCED. GitHub showed PR #1 already merged at `2026-10-08T16:51:44Z`, with live main `1c8c681cae0019dc2f413803e150951a962287ff`. A normal fetch and fast-forward merge incorporated that actual history into the existing M0 branch. PR #1 could no longer be reused; draft PR #2 continues the remaining M0 work. No history rewrite or runtime edit occurred.

M0-T01 implementation C was `ac09c024589158e1402ae63263818466d3e77d5c`. Its [handoff evidence](evidence/M0-T01-handoff.md) and [live audit](evidence/M0-T01-live-audit.json) retain the missing initial Git metadata, safe reconciliation/import, seven preserved original blobs and actual helper checks (57 tests: 56 passed, one symlink skip). Its final `4f550a2` receipt is historical and was reverified before M0-T02. The recorded original baseline remains `6edbed7c1a0c94968999882c5a46d90492d3c327`; all seven original application files still match it byte-for-byte.

## Finish rule

Only RELEASE_RUNBOOK.md's verified public non-draft/non-prerelease v1.0.0 release, anonymously downloaded matching assets, fixed tag and final synchronized main allows COMPLETE. No tag/release was created in M0-T02. Unsigned status must be disclosed if signing is omitted.
