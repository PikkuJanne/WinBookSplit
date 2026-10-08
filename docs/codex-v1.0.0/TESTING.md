# Test strategy and evidence

## Test layers

Characterize existing behavior in M0 with generated safe inputs, then convert findings into regression tests. Python unit tests cover parser/bookmark normalization/plan invariants and filename rules. Integration tests create PDFs with actual pypdf, execute the engine, reopen slices and verify page identity/order. Pester covers PowerShell binding, paths, discovery, child process handling and exit propagation. Real Windows tests cover launcher/Explorer, encoding, cancellation, filesystem behavior and actual EPUB/AZW3 conversion.

M0-T02's executable original-behavior characterization is documented in `tests/baseline/README.md`; fixture generation/provenance is in `tests/fixtures/README.md`. Its `expected_original.json` contains known-bad observations, while `PLAN_ORACLES.json` continues to contain corrected target behavior. A successful characterization is evidence of reproduction, not a fixed-engine acceptance pass. The JSON report records actual page identities, both expectation sets, source hashes, dependency versions and explicitly unrun layers.

`ACCEPTANCE_CASES.json` and `PLAN_ORACLES.json` are requirements and expected outcomes, **not executed tests**. M0/M1 connect these IDs to actual tests. Test harnesses may use pytest, Pester, PSScriptAnalyzer and development-only fixture/render tools, pinned to reviewed versions. Keep runtime dependencies minimal and separate. Add one convenient local runner using the same commands as CI; failed native commands must fail the runner.

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
