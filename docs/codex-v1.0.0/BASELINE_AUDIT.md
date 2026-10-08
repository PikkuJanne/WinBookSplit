# Baseline audit and reproduction obligations

Observed on 7 October 2026: `main` was `6edbed7c1a0c94968999882c5a46d90492d3c327` and the GitHub releases collection was empty. The seven-file tree and blob identifiers are in `BASELINE.json`. Repository references are in `SOURCES.md` [R1-R5]. The old display label is `2.1 (AZW3 Support)`, not a published GitHub release identifier.

Re-read the live checkout in M0. Do not reset newer work to this snapshot. These findings come from source inspection; Codex must run executable regression reproductions and Windows tests. No claim of end-to-end baseline validation is supplied.

| Finding | Code location/behavior | Reproduction or check |
|---|---|---|
| Manual first-page omission | `if 1 not in raw_nums` plus `idx > 0` | Ten pages: `1` yields no planned sections; `1,4,7` drops 1-3. |
| Invalid tokens silently filtered | `isdigit()` filtering before validation | `abc`, `1-4`, `0,999`; reject these in fixed version. |
| Front matter omitted | Auto slices begin at first selected destination | First Level 1 bookmark at physical page 4. |
| Level 2 crosses parent boundaries | Selected level is flattened globally | Child at 4 in parent 1, next child's parent starts at 9 and child at 11. |
| Overwrite/stale outputs | Reused book-name directory; `open(...,"wb")` | Repeat same book or two distinct books with same basename. |
| Neighboring PDF collision | `ChangeExtension(InputFile, ".pdf")` for conversion | Existing `Book.pdf` beside `Book.epub`. |
| stderr ignored | Redirected but never drained | Child writes enough stderr before stdout/exit. |
| False completion | unconditional `Done.`; fallback result not captured | Fail engine, fallback, or writing and inspect batch exit status. |
| Temporary-engine collision | shared `WinBookSplit_Engine_v2.py` | Two overlapping runs. |
| Literal-path gap | wildcard-aware `Test-Path`, no Leaf test | Brackets in filename; directory named `Book.pdf`. |
| Dependency gap | bare `python`; fixed Calibre paths | Missing/wrong Python/pypdf; nonstandard Calibre install. |
| UI/docs mismatch | menu defaults and README claims | Invalid selection, extra dropped files, stated custom ranges/logging. |

The current source is retained by Git, not copied into this bundle. M0 should characterize the original behavior before extraction and isolate changes. MIT license and existing artwork remain in the repository; do not relicense or redesign them as part of this release.

## Live reconciliation on 8 October 2026

M0-T01 confirmed that live main and all seven supplied application files still match the recorded commit/tree/blob identities. The supplied directory initially lacked Git metadata; live history/index metadata was attached without checking out or changing existing files. No newer application work or existing handoff guidance was found. See [the factual task evidence](evidence/M0-T01-handoff.md) and [machine-readable live audit](evidence/M0-T01-live-audit.json) for commands, environment, limitations and checkpoint state. The findings above still require executable reproduction in M0-T02; they are not passing regression tests.

## M0-T02 executable characterization on 8 October 2026

[Task evidence](evidence/M0-T02-baseline.md) and [actual machine report](evidence/M0-T02-characterization.json) now record 14 unchanged embedded-engine runs on original generated 10/12-page fixtures plus two actual launcher early error paths. Exact output page IDs demonstrate manual first-page loss, invalid-token filtering, front-matter omission and nested Level 2 parent crossing. `4,7` is the complete-coverage control. Original expectations are explicitly separate from corrected `PLAN_ORACLES.json` targets; reproducing known-bad behavior is not a fixed-runtime acceptance pass. Input/source/sentinel hashes and repeatable fixture bytes/provenance were verified; no runtime code changed or private documents were used.

PR #1 had already been merged by this task's start; actual main was `1c8c681cae0019dc2f413803e150951a962287ff`. That live history was safely fast-forwarded into the feature branch and draft PR #2 continues M0. All seven original raw application blobs still match the recorded baseline. Broader launcher/process/output hazards in the source-level table remain for later executable tests; full Windows/Explorer/Calibre/package acceptance remains unrun.
