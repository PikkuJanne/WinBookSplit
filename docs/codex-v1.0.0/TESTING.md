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

M1-T06 adds `--layer diagnostics` as the eighth full stage. Generated flat,
zero-page, empty, corrupt, truncated and outlined documents exercise structured
results, no-plan distinctions and zero-output failures. Actual PS5.1 and PS7
handler probes validate category-specific messages, permitted choices, explicit
decisions and rejection of malformed or conflicting protocol records. The
current exit codes remain 0/1/55 pending the final M3 CLI contract. Safe process
draining, native argument preservation and final failure propagation are narrow
console integration checks; they do not certify Explorer, conversion, full
interactive UX or output transactions. See `tests/diagnostics/README.md` and the
task's actual evidence for executed bounds and limits.

M2-T01 adds `--layer output` as the ninth full stage. Four repeat/same-basename
CLI runs and four actual barrier-synchronized processes under identical timestamps
must publish distinct complete manifests while preserving earlier run hashes.
Five failures cover partial write, PDF reopen, wrong count, manifest and promotion;
eight cleanup cases cover only-owned removal and unexpected/held/reparse refusal,
including actual controlled Windows junctions. Every successful slice is reopened
for exact page identities/content/order, expected positive counts, size and hash.
Current mode/plan/diagnostic tests consume only the explicit returned final folder
and exact manifest membership. Existing loose files in the base are now preserved
while a new child is published; M1's existing-output refusal was historical.

The current actual launcher helper uses only emitted final/log paths and marked
GUID parents, exact members and held identities. It reacquires delete handles and
checks replacement identity before removal; no recursive private Documents scan
or manifest-directed deletion is allowed. Six success launchers and separately
recorded controlled failure probes retain their bounded process/discovery/UX scope.
Original/extraction harnesses, fixtures/oracles and older receipts remain unchanged.
See `tests/output/README.md` and M2-T01 evidence for commands and final outcomes.

Windows sharing probes must distinguish actual data protection from compatible
metadata access. Already observed empty-stage reparse mutation rejects before
owner creation, but check/create is not atomic. Closing child locks before final
checks/rename is necessary; no arbitrary same-account process sandbox or crash
recovery guarantee is claimed. Source/neighbor preservation and safe held-object
cleanup remain required. These local counts do not certify Calibre, Explorer,
PDF rendering/features, all cancellation paths or a release ZIP.

M1-T05 adds `--layer plan` as the seventh full stage. Structural cases test the
shared coverage validator; all-mode prepared previews execute through the real
writer and are reopened for exact filename/range/physical-page identity parity.
Nested preview mutation attempts, controlled same-path replacement/repointing
and captured-source binding remain distinct from immutable generated fixtures.
Seeded complete partitions supplement exact cases. The callable preview creates
no chapters; console preview/confirmation still belongs to M3. At the M1-T05 checkpoint, source/output
alias and existing-output refusals were narrow pre-writer checks; M2-T01 now
provides the staged output lifecycle described above. Historical mode/original/
extraction guards, oracles and prior evidence stay unchanged.

At each task run targeted tests and related regressions. At a milestone run the relevant full layer and dependency/host compatibility checks. At release run the entire claimed matrix against exact release payload bytes. Pure Python tests may run on the available development host, but do not certify Windows behavior. Windows PowerShell 5.1 and PowerShell 7 require separate runs; a pwsh-only pass is not both. Document the actual Windows build, Python, pypdf, Calibre, shell and fixture versions.

Generate unique per-page markers and known outline trees. Test nonempty disjoint complete coverage plus page identities; a matching total count alone could conceal a duplicated page. Add seeded randomized start lists/bookmark arrangements to check invariants, while keeping exact deterministic fixtures for bugs. Test failures after some writes and verify source/neighbor hashes remain unchanged. When runtime rejection is the policy (encryption/forms/signatures), assert clear rejection before outputs rather than demanding support.

## Windows release-package smoke

Use a fresh extracted ZIP, fresh supported venv, controlled runtime discovery/PATH and an unrelated working directory. No importing the developer checkout, global pypdf accidents or copied untracked engine files. A clean Sandbox/VM is preferable when available; a current-workstation isolated setup is acceptable under D11 if documented honestly, not called a clean OS. Dependencies may be preinstalled by explicit setup. Then test normal offline processing. No new hardware or live UNC lab is mandated for unsupported platforms.

Run actual Explorer drag/drop for PDF, EPUB and AZW3, no-input selection/cancel, multiple-file rejection, filenames with spaces/brackets/unicode, automatic/manual/Level 2, conversion failure, no bookmarks fallback, open-output opt-in and a repeat run. Record human results when interaction cannot be automated. A skipped interaction is an open mandatory item, not a pass. User-supplied evidence can complete the item when it identifies package hash, environment, procedure and result.

Real conversion fixtures: author a tiny original EPUB with TOC/chapters, and generate a redistributable AZW3 fixture with documented Calibre settings; use both as actual input formats in independent runs. Record conversion settings and inspect generated PDFs/chapter coverage. Do not test AZW3 only by renaming an EPUB. Use an explicit expected-failure case for DRM/unsupported conversion without acquiring or bypassing protected material.

## Current M2-T02 path acceptance

`--layer paths` requires both actual supported hosts; `--layer full` now includes
it as the tenth stage. See `tests/paths/README.md` and
`evidence/M2-T02-paths.md/json`. Three actual PS5.1/PS7/BAT successes retain literal
special-character mixed-case/read-only input metadata and exact content against a
wildcard decoy. Eight actual rejects cover directory/provider/parser/held-source
failures. Nine title cases and 120 actual outputs per mode verify deterministic
Unicode/fallback/width/lexical identity. A bound long preview executes unchanged
within independently measured UTF-16 budgets; direct-API traps reject any planner
or source-reader reopen. Five failures cover impossible bound/unbound bases,
changed base, actual sharing-denied base and explicit ENOSPC after a real slice.
Reports/strict runner require actual host/read-only observations, no-replan proof,
complete content/manifest/source/neighbor/policy checks and authenticated cleanup.
Held controls are not ACL tests, and ENOSPC is not a filled physical disk.
Console-init statement-block/native reservation receipts have separately limited
scope. Historical source guards/fixtures/oracles and raw failed/earlier passing
receipts remain unchanged. Explorer/Calibre/CI/package gates remain open.

## Current M2-T03 conversion acceptance

Conversion/full requires both actual hosts and explicit pinned Calibre 9.15.0;
the T03 full route had eleven stages; T04 adds runtime acceptance as the twelfth.
PDF-only target layers remain converter-free.
See tests/conversion/README.md and evidence/M2-T03-conversion.md/json. Eight real
whole-PS format/host/retention runs, two independently inspected captured-reader
controls, eight actual native-zero invalid controls and actual BAT dependency
failure are required. All three EPUB/four AZW3 physical pages and three chapter markers
must survive once; exact captured source-to-slice content matches independent real references.
Default source vectors are explicitly engine metadata; direct/retained PDFs have
independent byte observations. Exact conversion-only membership includes optional
fullPDF; inherited PDF guards stay strict. Original/read-only ebook/neighbor hashes,
manifests, retained bytes, policies and known cleanup are required. Scoped native
fake-child and parent/grandchild timeout/injected-cancel checks supplement acceptance.
These do not certify human Ctrl+C/Explorer/GUI, expanded compatibility, remote
resources/DRM, package or release. Successful portable BAT discovery is separately
required and observed in M2-T04 runtime acceptance.

## Current M2-T04 runtime acceptance

`--layer runtime` requires both actual distinct supported hosts and the explicit
real pinned Calibre executable; full now has twelve stages. See
`tests/runtime/README.md` and `evidence/M2-T04-runtime.md/json`. The exact 33 actual
whole-PS/BAT cases cover ten Python/pypdf refusals, eight runtime selections, two
book/CWD/PYTHONPATH shadow controls, two PDF-without-Calibre successes, six converter
refusals and five real portable EPUB successes including unchanged BAT PATH.
Missing/wrong/non-Python/real unsupported 3.14.7 candidates reject before output;
explicit/app-venv/multiple-PATH/authored native launcher-listing selections use
actual supported Python/pypdf and isolated `-I -B` engine execution.

The independent strict validator requires exact IDs, actual host/version/argv,
meaningful native probe/reason and chosen paths, no late output/conversion on
dependency failure, complete physical-page/content/manifest parity, read-only
source/neighbor identity/hash/attribute parity, effective decoy positive controls,
absent acceptance markers and known ownership cleanup. Promised missing/malformed
receipts fail. Four synthetic validator tests cover 47 contradictions plus missing/
duplicate cases; they are structural checks, not actual Windows acceptance.
Full Pester includes existing nine plus 19 runtime tests per host. The deliberate
Pester failure still returns nonzero after later native success. Static/syntax
gates include the helper; legacy application findings remain observations.

Authored py listing controls do not claim actual manager registration/install.
Wrong pypdf alters only an owned copied package declaration; the real pinned dev
package remains unchanged. Real 3.14.7 rejection is tested; broader unsupported
runtime compatibility is not. ENGINE logs are launch intent; actual frames and
reopened chapter PDFs prove execution. Parent-only preflight timeout never claims
stopped descendants or authorizes cleanup. Existing owned-job conversion remains
separate; M2-T05 owns remaining engine supervision/UTF-8 and cumulative review.
Human Explorer/clean-OS/CI/package/release checks remain open. Earlier failed/partial
receipts remain distinct, including two rejected partial workspace removals outside
Git; actual application output and final suite cleanup are checked separately.

## Current M2-T05 process acceptance

`--layer process` requires both distinct actual supported hosts and no converter.
The full route has thirteen stages and preserves real Calibre conversion/runtime
regressions separately. See `tests/process/README.md` for the native and whole
entrypoint controls. Both UTF-8 streams require exact totals, bounded final tails,
blank/no-newline preservation, separate valid result capture and literal argv.
Timeout/injected cancellation must prove the owned job stopped, observe authored
PIDs stopped and preserve an unrelated process. Real Unicode/hostile PDFs require
reopened physical page/content and filename/manifest parity through PS5.1/PS7/BAT.

Dependency preflight now shares the owned supervisor; Runtime Pester checks
inherited descendants after parent exit. Its old parent-only receipts above are
historical. Fixed Python probe and engine use explicit `-X utf8` with `-I -B`.
The diagnostics regression now requires bounded retained bytes/truncation and
full drain counts rather than unbounded copies of its authored 200K streams.
Malformed, duplicated, missing or oversized process receipts must fail. Focused
Pester and synthetic receipt-validator checks retain their own narrower scope.
Automated token cancellation is not a human Ctrl+C/Explorer pass; interrupted
engine stages remain for inspection without outer-supervisor filesystem cleanup.

## Current M3-T01 outcomes acceptance

`--layer outcomes` requires both distinct actual supported hosts and exercises
the shipped BAT bytes in controlled copied applications. `--layer full` adds this
fourteenth stage to all inherited page/plan/output/path/runtime/process checks
and real Calibre EPUB/AZW3 checks. See `tests/outcomes/README.md` for the exact
failure/fallback/cancellation/finalization matrix and ownership recovery method.
The strict independent validator requires native and final structured outcomes
to agree, selects the final attempted fallback, proves every authored process
stopped and unrelated Python survived, and binds source/application/native-control
bytes and source/neighbor/prior-output identities. Missing or contradictory
promised receipts fail. Validator unit tests are structural evidence only.

Token cancellation and actual OS CTRL_C_EVENT in hidden private consoles have
distinct controls; neither is a human keypress or Explorer claim. Outer abrupt
interruption retains marked staging when no engine ledger is available. Known
owned failure cleanup, unsafe-cleanup refusal, manifest/publish errors and
published-output retention after finalization errors are separately checked.
Focused Pester verifies mapped protocol and secondary-log-failure handling under
each actual host; these function controls are narrower than application acceptance.

Corrected regression helpers authenticate historical original Git bytes and
unchanged fixtures/oracles while permitting the reviewed current BAT to evolve.
Original baseline/extraction guards remain untouched and still reject a changed
launcher; historical AC-011 reads immutable original/extracted Git source.
Code 7, full CLI binding, preview/version, menu/Explorer, PDF feature/fidelity,
extracted package, CI and release remain later gates.

## CI

Use a modest Windows workflow for static checks and automated tests, with explicit separate `powershell` and `pwsh` host executions. Configure dependency/tool installation reproducibly, validate current actions/runner support before pinning, and expose test results on failure. Isolate a real Calibre integration job or verified local release evidence when runner installation is not suitable. CI cannot be claimed to have run a human Explorer check. Do not disable failed jobs to make release green.

## Evidence discipline

Fill `templates/TASK_EVIDENCE.md` with commands, exit codes, test counts, tested code commit/tree digest, dates, actual environment, scope and skipped IDs. Keep large or private raw artifacts outside Git; commit a redacted summary or stable CI artifact reference with hashes. Do not put a commit's own SHA inside that commit. `GITHUB_WORKFLOW.md` separates code commit C, evidence checkpoint E and fresh live receipts.

Mandatory cases must pass or be satisfied by their specified explicit rejection outcome. Only cases explicitly marked conditional may be N/A with evidence. A release gate may not be waived solely because the environment cannot run it. Record a precise blocker and next action. Documentation support claims must match evidence; Windows 10/UNC are unclaimed by default, not secretly 'passed'.
