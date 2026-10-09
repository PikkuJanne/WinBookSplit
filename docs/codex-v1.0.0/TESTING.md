# Test strategy and evidence

## Test layers

Characterize existing behavior in M0 with generated safe inputs, then convert findings into regression tests. Python unit tests cover parser/bookmark normalization/plan invariants and filename rules. Integration tests create PDFs with actual pypdf, execute the engine, reopen slices and verify page identity/order. Pester covers PowerShell binding, paths, discovery, child process handling and exit propagation. Real Windows tests cover launcher/Explorer, encoding, cancellation, filesystem behavior and actual EPUB/AZW3 conversion.

M0-T02's executable original-behavior characterization is documented in `tests/baseline/README.md`; fixture generation/provenance is in `tests/fixtures/README.md`. Its `expected_original.json` contains known-bad observations, while `PLAN_ORACLES.json` continues to contain corrected target behavior. A successful characterization is evidence of reproduction, not a fixed-engine acceptance pass. The JSON report records actual page identities, both expectation sets, source hashes, dependency versions and explicitly unrun layers.

`ACCEPTANCE_CASES.json` and `PLAN_ORACLES.json` are requirements and expected outcomes, **not executed tests**. M0/M1 connect these IDs to actual tests. Test harnesses use stdlib unittest, Pester, PSScriptAnalyzer and development-only fixture/render tools, pinned to reviewed versions. Keep runtime dependencies minimal and separate. The local runner is `tests/run_tests.py`; M4-T05 must use the same entry point in CI. Failed native commands fail the runner.

M0-T03 selects stdlib unittest on regular CPython 3.14.8, Pester 6.2.0 and PSScriptAnalyzer 1.25.0 for M0-T04, with ReportLab 5.0.1/Pillow 12.3.0/charset-normalizer 3.5.2 only in `requirements-dev.txt`. `SUPPORT_AND_SETUP.md` distinguishes frozen targets from historical observations and documents absolute paths/hash-verified isolated setup. See `tests/README.md` for full/targeted invocations and `tests/powershell/README.md` for exact-version external shell tools. M0-T04 evidence records actual both-host harness outcomes, deliberate failure probes and repeated unrelated-directory runs; these do not pass repaired-engine, real conversion, Explorer or full launcher acceptance.

M1-T01 added an explicit extraction-equivalence route in
`tests/extraction/characterize_extraction.py`; see its README. It compares the
immutable original Git engine with the then-shipped engine, including all 14
understood known-bad cases, complete output bytes/page identities and safe
import. The unchanged historical baseline harness must refuse the modified
PowerShell source; `--layer baseline` therefore intentionally fails after
extraction. At M1-T01, `--layer extraction` and `--layer full` selected that
comparison; M1-T02 changes the current full route as described below. With both
actual host paths, that checkpoint ran six bounded actual PS5.1/PS7/batch probes
covering unrelated CWD and parallel launches. Their generated GUID
outputs use the workstation's observed Documents folder with exclusive owner
markers, synthetic neighbors and strict direct-file cleanup checks. No private
Documents entries are enumerated or processed. Successful observed cleanup
does not claim cleanup after every possible ownership/initialization failure;
invalid ownership or unexpected members cause safe refusal and preservation.
These controlled stdin/PATH/process-policy probes do not certify Explorer,
ordinary discovery, all launcher paths or release-package behavior.

M1-T02 deliberately transitions the current `--layer full` route to strict
manual acceptance (`--layer manual`), documented in `tests/manual/README.md`.
All 22 unchanged manual oracles execute against the shipped engine, including
whole-document output, every invalid token and zero-page rejection. Exact page
identities/order, normalization notices, eight extra CLI edge cases, 250 seeded
plans and 25 actual writer samples are checked. At that checkpoint, safe import
and three unchanged known-bad bookmark observations remained separate checks. Historical
AC-011 AST evidence reads immutable M1-T01 C; the original/extraction harnesses
and their guards remain unchanged. Their explicit legacy layers now refuse the
changed launcher/processing ASTs as expected. No current mechanical manual
equivalence is claimed. Optional actual entrypoint probes retain the bounded
GUID/marker-owned method and its original limits.

Reports use new absolute paths outside the checkout, include actual source-byte and child-evidence SHA-256 digests, and retain every step's native exit status. A missing/invalid promised child report also fails. Each run owns one unique temporary directory; source and synthetic-neighbor hashes are checked before cleanup. AC-009 requires observed native/Python/Pester failures followed by success while the runner stays nonzero. AC-010 requires two actual full runs from unrelated directories, matching deterministic outcomes and owned cleanup. Acceptance IDs printed by a runner alone are not acceptance passes. The shell layer gates syntax and selected new-scaffold defect/security/compatibility rules while separately recording unchanged application analyzer findings; legacy observations are not a runtime static pass.

M1-T03 added `--layer bookmarks` as the fifth stage of the then-current full
route. It checks corrected Level 1 BM-01/02/05/08 targets, invalid destinations,
depth/lineage and bounded malformed traversal, 150 seeded plans and ten actual
writer samples. Every output slice and flattened page sequence must preserve
physical identities. At that checkpoint, the manual route retained only the unchanged
Level 2 BM-03 observation; M1-T02's three-case record above is historical.
Original/extraction guards, oracles, generator and prior evidence remain
unchanged. This bookmark route runs Python CLI/callable checks; the six existing
actual shell/batch probes then covered manual MAN-03 and historical Level 2 BM-03.
Corrected Level 1 console/Explorer and full application acceptance are separate
later gates. See `tests/bookmarks/README.md` for exact bounds and evidence scope.

M1-T04 adds `--layer level2` as the sixth full stage. Corrected BM-03/04/06/07
execute through the actual Python CLI, including parent opening/fallback slices,
outside-child warnings and structured no-plan rejection. Six hierarchy aggregates,
150 seeded parent/child plans and ten real writer samples verify exact page
identities, selected-parent ownership and complete coverage. Manual/Level 1
regressions remain in the full route. The current manual harness no longer
compares any repaired bookmark mode with historical defective behavior: its
new corrected BM-03 CLI reference supplies the existing six owned launcher
probes. Their two actual PS7 Level 2 probes now match corrected complete
output bytes/page identities. This is controlled entrypoint evidence; ordinary
discovery, full interactive fallback/Explorer, Calibre and output transactions
remain separate later gates. Prior bookmark comparison records are historical;
original/extraction guards, fixtures/oracles and evidence are unchanged.

## Risk-based execution

At each task run targeted tests and related regressions. At a milestone run the relevant full layer and dependency/host compatibility checks. At release run the entire claimed matrix against exact release payload bytes. Pure Python tests may run on the available development host, but do not certify Windows behavior. Windows PowerShell 5.1 and PowerShell 7 require separate runs; a pwsh-only pass is not both. Document the actual Windows build, Python, pypdf, Calibre, shell and fixture versions.

Generate unique per-page markers and known outline trees. Test nonempty disjoint complete coverage plus page identities; a matching total count alone could conceal a duplicated page. Add seeded randomized start lists/bookmark arrangements to check invariants, while keeping exact deterministic fixtures for bugs. Test failures after some writes and verify source/neighbor hashes remain unchanged. When runtime rejection is the policy (encryption/forms/signatures), assert clear rejection before outputs rather than demanding support.

## Windows release-package smoke

Use a fresh extracted ZIP, fresh supported venv, controlled runtime discovery/PATH and an unrelated working directory. No importing the developer checkout, global pypdf accidents or copied untracked engine files. A clean Sandbox/VM is preferable when available; a current-workstation isolated setup is acceptable under D11 if documented honestly, not called a clean OS. Dependencies may be preinstalled by explicit setup. Then test normal offline processing. No new hardware or live UNC lab is mandated for unsupported platforms.

Run actual Explorer drag/drop for PDF, EPUB and AZW3, no-input selection/cancel, multiple-file rejection, filenames with spaces/brackets/unicode, automatic/manual/Level 2, conversion failure, no bookmarks fallback, open-output opt-in and a repeat run. Record human results when interaction cannot be automated. A skipped interaction is an open mandatory item, not a pass. User-supplied evidence can complete the item when it identifies package hash, environment, procedure and result.

Real conversion fixtures: author a tiny original EPUB with TOC/chapters, and generate a redistributable AZW3 fixture with documented Calibre settings; use both as actual input formats in independent runs. Record conversion settings and inspect generated PDFs/chapter coverage. Do not test AZW3 only by renaming an EPUB. Use an explicit expected-failure case for DRM/unsupported conversion without acquiring or bypassing protected material.

## CI

Use a modest Windows workflow for static checks and automated tests, with explicit separate `powershell` and `pwsh` host executions. Configure dependency/tool installation reproducibly, validate current actions/runner support before pinning, and expose test results on failure. Isolate a real Calibre integration job or verified local release evidence when runner installation is not suitable. CI cannot be claimed to have run a human Explorer check. Do not disable failed jobs to make release green.

## Evidence discipline

Fill `templates/TASK_EVIDENCE.md` with commands, exit codes, test counts, tested code commit/tree digest, dates, actual environment, scope and skipped IDs. Keep large or private raw artifacts outside Git; commit a redacted summary or stable CI artifact reference with hashes. Do not put a commit's own SHA inside that commit. `GITHUB_WORKFLOW.md` separates code commit C, evidence checkpoint E and fresh live receipts.

Mandatory cases must pass or be satisfied by their specified explicit rejection outcome. Only cases explicitly marked conditional may be N/A with evidence. A release gate may not be waived solely because the environment cannot run it. Record a precise blocker and next action. Documentation support claims must match evidence; Windows 10/UNC are unclaimed by default, not secretly 'passed'.
