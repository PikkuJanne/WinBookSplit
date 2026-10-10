# M0-T02 — Original behavior and safe fixture evidence

Task: M0-T02, capture baseline behavior with safe fixtures.
Date: 8 October 2026, Europe/Berlin (UTC+02:00).
Author/reviewer: Codex implementation and independent read-only agent review.
Starting checkpoint: `4f550a2cdfb9736be8c2fd35ab83cc65c1f61ffc`.
Reconciled base: `1c8c681cae0019dc2f413803e150951a962287ff`, tree `51daa9b7c1a3a0eb11808ed8b14deb81357b40df`.
Implementation C: `8f1f556f318be488839e839830264fa2e5f4fbc8`, tree `888e3a4df72c34ce738df5d82e8518ff5108350c`.
Tested source-path digest: `b86df43bc289003877949b922be934a4cc99c0afeca9ec8b256339ce70fa6ee1`.
Embedded engine SHA-256: `44e583cbe022b070c537f2229e20fc39c02dd76acef1caebc9f29a136a306cd9`.

The machine report [M0-T02-characterization.json](M0-T02-characterization.json) records the honest pre-commit tested path hashes, original source identity, fixture provenance, versions, stdout/stderr, exit codes, exact output page IDs/ranges/hashes and target oracles. Its digest is SHA-256 of the UTF-8, sorted-key `json.dumps(tested_path_sha256, sort_keys=True)` mapping. The clean-C run was compared with that report: tested source digest, full fixture provenance, fixture bytes, all 14 engine observations and launcher probe results match. The report's `base_commit` identifies the pre-commit base, not a claim that its new files already existed in that tree.

## Live reconciliation and scope

At thread start, `check_sync.py` exited 0: clean local/live feature HEAD `4f550a2cdfb9736be8c2fd35ab83cc65c1f61ffc`, observed `2026-10-08T16:52:38.071392+00:00`. Live GitHub showed PR #1 already merged at `2026-10-08T16:51:44Z`; main was its merge commit `1c8c681cae0019dc2f413803e150951a962287ff`. A normal fetch and `git merge --ff-only origin/main` advanced the existing M0 branch to that live history. No reset, stash, force push, settings change or application edit occurred. The earlier instruction to reuse draft PR #1 was stale; a new [draft PR #2](https://github.com/PikkuJanne/WinBookSplit/pull/2) continues the remaining M0 work. This does not mark the M0 milestone review/merge gate complete.

Added only development fixture/characterization code, documentation and safe generated evidence. No runtime extraction or bug fix was performed. All seven original application raw blobs still match `BASELINE.json` using `git hash-object --no-filters`; the embedded body is extracted to a per-run temporary file, with a guard against PowerShell here-string interpolation. It is executed unchanged with the explicit isolated Python interpreter and real pypdf writers.

## Actual environment

- Windows platform probe: `Windows-11-10.0.26300-SP0`; registry build `26300`, UBR `9457`, DisplayVersion `26H2`. This is the existing workstation, not a clean OS installation.
- Orchestrating PowerShell: 7.6.5. Actual early launcher probes use `powershell.exe`, Windows PowerShell 5.1; exact version is captured in the JSON report.
- Explicit bundled Python: `<BundledPython>`, 3.12.14, MSC v.1944 64 bit. The identifying local path remains only in the external raw receipt; redacted during M5-T04.
- pypdf 6.10.0 and ReportLab 4.4.9 from that bundled runtime; children use `-I`. No global dependency installation occurred.
- Poppler `pdftoppm` 26.07.0 from the bundled native runtime.
- Default PATH Python remains 3.14.7 without the inspected pypdf dependency; it was not used for fixture/engine execution. No Calibre conversion was run. M0-T03 must select supported versions using current primary sources and real tests; these development observations do not freeze that matrix.

## Commands and actual outcomes

`$FixturePython` below is the explicit bundled path above. `$ReportPath` is an explicit new file under this run's external temporary area. Commands used the repository root unless another working directory is stated.

| Actual command/check | Exit / observed outcome |
| --- | --- |
| `& $FixturePython -I tests/baseline/characterize_original.py --launcher-probes --report $ReportPath` | 0: 14 original engine cases and 2 launcher boundary probes reproduced. Every PDF slice reopened and its exact page IDs/ranges checked. |
| Same command with absolute script path, from `<UserTemp>` | 0: same tested-source digest, fixture/provenance bytes and all original observations; no dependency on current directory. |
| Same command on clean C, new external report | 0: same source/provenance/result comparisons after committing C; actual C recorded in that external report. |
| Same command with an existing report path | 1: `Report path already exists`; prior report preserved. |
| Generator CLI `--output-dir <new owned directory>` | 0; 10/12-page fixtures reopened with exact IDs and outline trees. |
| Generator repeated against that nonempty directory | 1; every existing fixture/provenance hash preserved. |
| Generator without required `--output-dir` | 2; argument rejection before generation. |
| Independent second generation into a different owned directory | PDF bytes and provenance bytes identical for both fixtures. |
| `pdftoppm -scale-to 500 -png <simple-10.pdf> <owned-prefix>` and the same for `nested-12.pdf` | 0; 10 and 12 PNGs produced; contact sheets inspected visually by Codex. First/last pages also inspected at larger resolution. |
| `python tools/codex-handoff/validate_plan.py --plan-root docs/codex-v1.0.0` | 0, VALID: 36 tasks / 107 acceptance cases / 30 split oracles; structural validation only. |
| Raw original blob comparison against `BASELINE.json` | All seven unchanged. Initial use of filtered `git hash-object` disagreed on README due to local `core.autocrlf=true`; repeating with `--no-filters` proved exact original bytes. |
| Current SHA-256 versus all machine report tested/provenance source hashes | All equal. |
| `git diff --cached --check` | 0; only seven intended implementation/evidence paths staged for C. |
| `git push origin HEAD`, then `check_sync.py` on C | Both 0; clean local HEAD equals fresh live branch at the recorded time below. |

During harness development, two runs reached report creation and then failed on report-version bookkeeping (missing module attribute, then a local variable shadowing the metadata-version function). Both exited nonzero and produced no success report. They were corrected before the successful reported runs. Poppler emitted missing-display-font configuration warnings for fonts not used by these fixtures; all 22 rendered pages have legible markers and no visible clipping, missing text or overlap. These warnings are recorded, not represented as a warning-free renderer pass.

## Acceptance reproduction results

Ranges here use physical 1-based inclusive pages for readability. JSON ranges use 0-based half-open endpoints from `PLAN_ORACLES.json`.

| ID / scenario | Actual original files / behavior | Required corrected result |
| --- | --- | --- |
| AC-004 / MAN-01: `1` on 10 pages | Exit 0, zero files, all pages missing | One file, pages 1–10 |
| AC-004 / MAN-02: `1,4,7` | Pages 4–6 and 7–10; pages 1–3 missing | Pages 1–3, 4–6, 7–10 |
| AC-004 / MAN-03: `4,7` | Pages 1–3, 4–6, 7–10; positive coverage control | Same complete partition |
| Invalid-token cases MAN-09/10/11/12/13/14/15/16 | `abc`, `1-4`, `1,abc,7`, `0,999`, `11`, `1,,4`, `1,`, `,4` all exit 0; invalid tokens silently discarded. Some also omit early pages or write zero files. | Reject each whole request before writing (`invalid_start_pages`) |
| AC-005 / BM-01: Level 1 first bookmark at page 4 | Pages 4–6 and 7–10; front matter 1–3 omitted | Pages 1–3, 4–6, 7–10 |
| AC-005 / BM-02: Level 1 parents at 3/9 | Pages 3–8 and 9–12; front matter 1–2 omitted | Pages 1–2, 3–8, 9–12 |
| AC-005 / BM-03: Level 2 children at 4/7/11 | Pages 4–6, 7–10, 11–12. Pages 1–3 omitted; A2 consumes B-opening pages 9–10. | Pages 1–2, 3, 4–6, 7–8, 9–10, 11–12 |

AC-004 and AC-005 are satisfied as **baseline reproduction obligations**, not repaired-runtime tests. Thirteen of the fourteen engine cases differ from the fixed target; MAN-03 is the matching control. Fixed-engine acceptance remains NOT RUN. No runtime defects were resolved in this task.

AC-006: both generated fixtures contain only original synthetic text, visible `WBS-PAGE-001` style identifiers and anonymous fixed metadata. No source PDFs, real textbooks, workplace documents, downloads or private content are used. The generator and generated content use the repository MIT license. Full source/license hashes, exact metadata, page IDs and nested outline trees are retained in the JSON provenance. SHA-256: simple10 `9b948f882e5a57f11fd37d30851fb13946e9fdcea0a7ab4faa2c07d17a6c8fa3`; nested12 `168ccef79eeefd524da87deb97bf1add724acdf7b93f88c41972cd9777ceac43`. Reproducibility applies to the same source bytes and dependency versions, not every future pypdf/ReportLab version.

For every engine case, source launcher hashes, fixture input hash and a neighboring synthetic prior-output sentinel hash were unchanged. Generated PDFs, extracted engine, slice outputs and sentinels are inside the run-owned `TemporaryDirectory`, cleaned on completion. Render QA artifacts are outside Git in a dedicated generated-only temporary folder. No private documents were processed or uploaded; no generated PDF/binary artifact is committed.

## Launcher evidence and open layers

Two actual unchanged early launcher paths were executed with redirected stdin/stdout/stderr and a timeout: Windows PowerShell 5.1 on an absent PDF printed `Invalid file` and returned 0; the batch launcher with no arguments printed its no-file message and returned 0. Both stopped before output preparation or the shared TEMP engine. They are bounded error-path reproductions, not full Windows/Explorer acceptance.

Source observations still requiring later executable checks include wildcard-aware input validation, reused output directories, neighboring conversion PDF collisions, shared temporary engine paths, stderr redirected but not drained, invalid menu defaults, ignored fallback results and unconditional `Done.`. No success-path PowerShell 5.1/7 splitting, human Explorer drag/drop, Calibre conversion, cancellation, hostile-path or release-package matrix result is claimed. No M0-T02 acceptance ID was skipped; those broader layers remain open in their scheduled tasks.

## Git, review and continuation

Branch: `codex/winbooksplit-v1-m0`; [draft PR #2](https://github.com/PikkuJanne/WinBookSplit/pull/2).
Normal push of C succeeded. Fresh live receipt at `2026-10-08T17:00:16.304703+00:00` (19:00:16 Europe/Berlin): worktree clean, local HEAD = live feature SHA = `8f1f556f318be488839e839830264fa2e5f4fbc8`.
Independent read-only review found no blocking source, safety, scope or acceptance/evidence issue. No GitHub review submission or CI execution is claimed. M0-T04 still owns milestone review, test scaffolding and merge. Tags/releases remain empty; no intermediate release was made.

The next evidence/status checkpoint references C and the receipt already observed above. Its own final push/clean/live receipt belongs in this thread/PR, outside its own tracked files. Recheck live synchronization at the next thread start. Exact next task: **M0-T03 — Freeze support and dependency decisions**. Read its task/specs, probe actual Windows/shell/Python/pypdf/Calibre and current primary documentation, and choose supported/runtime/developer dependencies in isolation. Keep this task's known-bad harness separate from repaired-engine acceptance. No blocker prevents starting M0-T03.
