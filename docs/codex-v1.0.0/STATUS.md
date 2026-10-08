# Project status

Project: WinBookSplit first public v1.0.0.
Overall: IN PROGRESS; application unchanged, original defects reproduced, support/dependency plan frozen, release unverified.
Completed tasks: M0-T01, M0-T02 and M0-T03.
Next dependency-ready task: **M0-T04 — Add local test and evidence scaffolding**, after fresh verification of this evidence checkpoint.
Active branch: `codex/winbooksplit-v1-m0`.
Origin fetch/push: `https://github.com/PikkuJanne/WinBookSplit.git`.
Last verified implementation C: `2e63f259a59438a689e75ff0ff9c1a9fcd600c1f`.
Historical C synchronization: SYNCED at `2026-10-08T17:33:12.039489+00:00` (19:33:12 Europe/Berlin), clean local HEAD equals live feature branch.
This documentation/evidence checkpoint receives its separate final push/live receipt in the thread/[draft PR #3](https://github.com/PikkuJanne/WinBookSplit/pull/3); recheck current equality next thread.
Public release: NOT PUBLISHED; live tags/releases were empty on 8 October 2026.

## M0-T03 evidence and limits

[Support/setup matrix](SUPPORT_AND_SETUP.md), [task evidence](evidence/M0-T03-support.md) and [machine record](evidence/M0-T03-environment.json) satisfy AC-007/008's selection/setup obligations. Frozen application Python is regular GIL CPython 3.14.8 x64 only (minimum = selected), plain pypdf 6.19.0; Calibre 9.15.0 is the mandatory ebook-conversion test target, optional for PDF-only use. Hashed runtime/dev requirements are separate. Developer pins are ReportLab 5.0.1/Pillow 12.3.0/charset-normalizer 3.5.2; stdlib unittest/Pester 6.2.0/PSScriptAnalyzer 1.25.0 are selected for M0-T04. Current primary vendor requirements/advisories were reviewed with precise source links; no automated scanner pass is claimed.

Official Python was extracted without runtime registration/global installs. Fresh runtime/dev venv installs, pip checks and import-origin/version/absence assertions passed. The complete documented runtime setup ran under both actual shell hosts. The Windows PowerShell 5.1 native quoting defect found during review was fixed and rerun; execution policy remained unchanged. All 14 original engine cases and two unchanged early launcher error paths reproduced on the final dependency set, with exact page identities and source/input/neighbor preservation. This remains known-bad characterization, not repaired-engine acceptance. All seven original application files remain byte-identical.

Actual workstation remains Windows 11 Pro 26H2 build 26300.9457 x64; hosts 5.1.26100.9444 and 7.6.5. Existing registered Python 3.14.7 and bundled 3.12.14 are historical/unchanged, not current supported pins. Calibre was not found; no conversion ran. External shell tooling and selected newer renderer are not installed/tested. Full PS5.1/7 splitting, Explorer, real EPUB/AZW3, repaired-engine and release-package checks remain NOT RUN. Windows10/ARM/UNC and other Python versions remain unclaimed. Independent source/evidence review found no remaining blocker; no GitHub review submission or CI execution is claimed. M0-T04's cumulative review/scaffolding/merge gate remains open.

## M0-T02 evidence and limits

[Baseline evidence](evidence/M0-T02-baseline.md) and [machine report](evidence/M0-T02-characterization.json) record 14 original embedded-engine cases and two actual unchanged Windows launcher error paths. The engine was executed with generated page-marked PDFs; each output was reopened and checked by exact page identity/order. Manual first-page loss, success with zero files, invalid-token filtering, bookmark front-matter omission and Level 2 parent crossing were reproduced. Known-bad observations are separate from the corrected `PLAN_ORACLES.json` targets. This completes AC-004/005/006's reproduction/provenance obligations; it does not pass fixed-engine acceptance.

The original synthetic 10/12-page fixtures reproduce byte-for-byte with the same source/dependencies. Full anonymous fixed metadata, source/license hashes and exact outlines are recorded. All 22 pages were rendered and visually inspected. Original source, inputs and sentinel neighbors remained unchanged. No real documents, generated PDFs or private artifacts were committed.

Actual fixture/engine environment: existing Windows 11 workstation build 26300.9457; explicit bundled isolated Python 3.12.14, pypdf 6.10.0, ReportLab 4.4.9. Orchestrating shell was PowerShell 7.6.5; the bounded launcher probe used actual Windows PowerShell 5.1.26100.9444. Poppler was 26.07.0. These are development observations; supported versions remain M0-T03's decision. No global dependency installation occurred.

Full Windows PowerShell 5.1/7 splitting workflow, human Explorer drag/drop, Calibre conversion, release-package checks and fixed-engine acceptance: **NOT RUN**. Independent read-only source/evidence review found no blocker. No GitHub review submission or CI pass is claimed. M0-T04's cumulative review/test-scaffolding/merge gate remains open.

## Repository reconciliation and earlier work

At M0-T02's start, clean local/live feature checkpoint `4f550a2cdfb9736be8c2fd35ab83cc65c1f61ffc` was freshly SYNCED. GitHub showed PR #1 already merged at `2026-10-08T16:51:44Z`, with live main `1c8c681cae0019dc2f413803e150951a962287ff`. A normal fetch and fast-forward merge incorporated that actual history into the existing M0 branch. Draft PR #2 continued that task. No history rewrite or runtime edit occurred.

At M0-T03's start, clean local/live feature checkpoint `869218c3cc4c8b7f6b47d057a5e8e4a414924173` was freshly SYNCED. PR #2 was already merged at `2026-10-08T17:16:55Z`; live main `a66c8f95f922c36c58b47b6cbfba82be399a552a`. A normal fetch/fast-forward incorporated that history before implementation. New draft PR #3 continues the remaining M0 work; closed PR #2 is not reused. The earlier PR merges do not complete the M0-T04 gate.

M0-T01 implementation C was `ac09c024589158e1402ae63263818466d3e77d5c`. Its [handoff evidence](evidence/M0-T01-handoff.md) and [live audit](evidence/M0-T01-live-audit.json) retain the missing initial Git metadata, safe reconciliation/import, seven preserved original blobs and actual helper checks (57 tests: 56 passed, one symlink skip). Its final `4f550a2` receipt is historical and was reverified before M0-T02. The recorded original baseline remains `6edbed7c1a0c94968999882c5a46d90492d3c327`; all seven original application files still match it byte-for-byte.

## Finish rule

Only RELEASE_RUNBOOK.md's verified public non-draft/non-prerelease v1.0.0 release, anonymously downloaded matching assets, fixed tag and final synchronized main allows COMPLETE. No tag/release was created in M0-T02. Unsigned status must be disclosed if signing is omitted.
