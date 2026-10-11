# Test strategy and evidence

## M5-T02 package checks

`tools/release/build_package.py` builds/independently verifies an explicit
commit's 28-file allowlist and three external assets. See its README for the
actual CLI recipe, byte provenance, complete dependency paths, noncircular
hashes, exclusive outputs and incomplete-build marker. Package unit/integration
controls in `tests/python/test_release_package.py` use owned real Git repositories
and forged asset hashes; the shared Python route discovers them automatically.
The source guard includes release tools and their allowlist. The Python child
uses a bounded 600-second deadline; historical failed 300-second M5-T01 evidence
is retained, with no old timeout reclassified as a pass.

AC-084/085/086 require an actual committed-candidate build, full member inspection,
independent Git/file/archive/manifest/checksum comparisons and missing/version/
unsafe-member negatives. Repeated actual builds establish reproducibility only
for the recorded tools. This task does not satisfy M5-T03 or M6's separate
extracted-candidate Explorer/Windows/Calibre, final asset or download gates.

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

The baseline interaction catalogue includes PDF/EPUB/AZW3 drops, no-input selection/cancel, multiple-file rejection, special filenames, automatic/manual/Level2, conversion failure, no-bookmark fallback, viewer opt-in and repeat. Record actual native and human evidence separately; a skipped required interaction is an open item, not a pass. User-supplied results qualify when bound to package/environment/procedure/outcome.

For current M5-T03, the user explicitly authorized fewer human checks or automation: **D13** routes actual interactions through completed **G01-G03** plus mandatory **R01-R04** (combined no-input/no-outline/manual-multiple starts, EPUB, genuine AZW3/manual/generated-PDF open, two-file rejection). Original unrun G04-G16 are SUPERSEDED, never PASS. Normalization/invalid-token/selected-level-fallback/conversion-failure/input-cancel/plan-cancel/identical-mode-repeat variants retain existing native evidence only. Human repeat means three actual same-input runs in different modes with prior outputs preserved; no extra same-mode GUI run is claimed. This explicit user-directed route replaces the16-row human catalogue for this task; all four current flows have now been evidenced as PASS in the M5-T03 human record. Preserve original guide/results, safety outcomes and acceptance IDs. Package/CI/M6 fixed-release and download gates remain intact.

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

## Current M3-T02 CLI acceptance

`--layer cli` requires both distinct actual supported hosts and the real pinned
Calibre executable; full adds it as the fifteenth stage. See
`tests/cli/README.md`. The 72 native controls use unavailable DEVNULL stdin and
the actual host's NonInteractive flag. Application NonInteractive, independent
NoPause and Preview are separately exercised. Explicit manual/Level 1/Level 2
execution preserves exact physical page identity/content/ranges; the documented
EPUB Auto/retention command verifies actual chapters and the retained full PDF.
PDF and real EPUB/genuine AZW3 previews have zero chapter publication, accurate
complete plans, original/generated identity and independently checked cleanup.
Version works in a helper-free copy containing only the entrypoint/shared JSON;
full comment-based help is read under each host.

Contradictory/missing choices, malformed tokens and unknown arguments reject
before discovery or writes. No-plan never prompts or retries; missing dependencies
and out-of-range physical starts preserve mapped native/final statuses. Exact
source/neighbor/prior identities, read-only attributes, copied source hashes,
stored policies and held-object output cleanup are required. The independent
validator refuses missing/contradictory promised receipts. A timed-out test retains
its workspace when owned descendant shutdown is unproved. Structural validator
units and focused Pester checks do not certify native application acceptance.
Current-build human checks are recorded separately in M3-T01-human.md/json;
these tests do not certify future package bytes, PDF rendered fidelity or CI.

## Current M3-T03 launcher acceptance

`--layer launcher` adds 70 redirected-input native controls under actual PS5.1,
PS7 and BAT; full includes it as stage sixteen. See `tests/launcher/README.md`
for exact source-copy controls and scope. It proves literal paths, no-input
selection/cancel130, multi-input rejection2 including quoted-empty argument 2,
exact menu/fallback reprompts and closed-input cancellation. Structured outcomes,
real Calibre ebook conversion, complete physical page contents/manifests,
source/neighbor/prior preservation, policy equality and known owned cleanup are
required. Successful native BAT copies declare only PowerShell parameter-default
changes; multi-input controls declare a PowerShell launch sentinel. They do not
prove Explorer interaction. Focused Pester decision helpers and strict receipt
mutation units are narrower tests.

AC-057/058/059 also require actual Explorer drag/drop and double-click procedures
on byte-identical committed application files and authored books. Record
assistant-operated UI checks separately from historical human M3-T01 receipts and
native stdin controls. Native exit observers must attach real process handles
before completion; an unobserved exit remains unknown. No new human result,
clean-OS/package, CI or rendered-fidelity claim follows from these tests.

## Current M3-T05 local support diagnostics

`--layer support` is stage 18 of full and requires both actual supported PowerShell hosts and pinned real Calibre. See `tests/support/README.md`. Its 66 native controls inspect successful/failed/cancelled/incomplete UTF8 manifests and closed bounded logs, real parser/bookmark warnings, explicit local redacted export, immutable sources and ownership-bound cleanup. The 32 application and 34 export controls include 12 disclosed copied-code finalizer faults and authored secret-bearing records; they are not human/GUI/upload tests.

Physical ranges/output manifests must agree with the captured plan and reopened page content. A completed split plus diagnostic failure is incomplete6 with retained output; primary failure/cancellation130 survive. The final log footer must match the authoritative outcome when usable. Pending publication never becomes an export input. Source/neighbor/prior/policy identities and no-overwrite/alias/reparse guards are required. Unknown or contradictory receipt claims fail closed.

Focused Logging/Support Pester and Python logger/strict-receipt/identity mutation controls are narrower proofs. Preserve failing reports/workspaces; pure retained-row revalidation is not a native rerun. A test-only fix may use reviewed exact production-byte equality plus affected native reruns and explicit retained passing evidence; do not relabel a failed full run as passed. M3-T05 records that precise combination. Human M3-T01 stays closed. Rendered fidelity, broader feature/conversion policy, CI and the exact release package remain separate tasks.

## CI

Use a modest Windows workflow for static checks and automated tests, with explicit separate `powershell` and `pwsh` host executions. Configure dependency/tool installation reproducibly, validate current actions/runner support before pinning, and expose test results on failure. Isolate a real Calibre integration job or verified local release evidence when runner installation is not suitable. CI cannot be claimed to have run a human Explorer check. Do not disable failed jobs to make release green.

## Evidence discipline

Fill `templates/TASK_EVIDENCE.md` with commands, exit codes, test counts, tested code commit/tree digest, dates, actual environment, scope and skipped IDs. Keep large or private raw artifacts outside Git; commit a redacted summary or stable CI artifact reference with hashes. Do not put a commit's own SHA inside that commit. `GITHUB_WORKFLOW.md` separates code commit C, evidence checkpoint E and fresh live receipts.

Mandatory cases must pass or be satisfied by their specified explicit rejection outcome. Only cases explicitly marked conditional may be N/A with evidence. A release gate may not be waived solely because the environment cannot run it. Record a precise blocker and next action. Documentation support claims must match evidence; Windows 10/UNC are unclaimed by default, not secretly 'passed'.

## M4-T01 page fidelity route

`--layer fidelity` is the nineteenth current full stage. See `tests/fidelity/README.md` and [T01 evidence](evidence/M4-T01-fidelity.md). Eight actual redirected app controls cover rich Manual/Level 1/Level 2 and genuine image-only Manual under both supported hosts. They reopen every chapter, compare page/content/image/geometry/static appearance, check serialized Page/content/image objects for outside-segment leakage, validate safe metadata/start bookmark and local destinations/fit/coordinates, and require visible warnings. Two actual explicit support exports retain only permitted categories/tokens. Sources, neighbors, prior files, settings and owned cleanup are authenticated.

Poppler and PDFium compare source/output dimensions/PNG bytes/RGBA at the same scale within each renderer. They are developer tools, not runtime requirements. Full/fidelity requires `--renderer-path` and `--secondary-python-path` with the applicable host/runtime flags. Persisted authored renders are separate from removed fixture work for agent image inspection. Automated comparison, agent inspection and closed historical human checks remain distinct. Synthetic mutations are pure oracle tests, not native application reruns.

T01 uses honest composite targeted/affected evidence: full01 remains FAILED/native1 with 18/19 stages passing; a fresh same-C process-only retry passes 35 controls. Together there are 19 passing layer observations, not a successful standalone full run. Preserve the failed process stderr and unknown case/cause without speculative diagnosis. M4-T02 owns broader unsupported-feature preflight; M4-T05 owns actual CI.

## M4-T02 unsupported-document policy route

`--layer document-policy` is stage20 of full, requiring both explicit actual supported PowerShell paths and isolated developer Python. See `tests/document_policy/README.md`, [policy](../PDF_POLICY.md) and [T02 evidence](evidence/M4-T02-document-policy.md). It runs 25 original inert fixtures,30 app controls per host,13 unchanged-BAT controls and2 explicit local exports(75 total). Its strict validator requires exact native argv/UTF8 stream hashes/PID/status, finite deadlines/closed streams and owned descendants, copied app bytes, observed versions/settings, read-only source/neighbor/prior identities and ownership-bound cleanup. BAT's three copied PowerShell defaults are disclosed and byte-proven; no native/human/GUI claims follow from synthetic receipt units.

Encrypted user and empty-user owner inputs must reject with exit 7 before plan/extraction/output; Preview writes nothing and interactive unsupported controls request no plan/input/fallback. Guarded callable tests separately assert Encryption.read/verify/decrypt uncalled. Forms/XFA/widgets/signature/catalog-action/JavaScript/attachment/associated-file/portfolio/3D/media cases reject with exit 7. Page-AA's explicit visible exclusion and ordinary Manual/Level1/Level2 controls preserve complete physical coverage and reopened content. Signature fixtures test detection rather than cryptographic validity. Rejected ordinary logs/failure manifests can remain; redacted exports must be finalized metadata-only/nullplan without private identities or text.

Malformed/truncated/zero/deep-outline/cyclic-page-tree inputs have bounded native 6 failures, no successful zero-file output and no invalid-token fallback. Generated malformed output remains exit 4, unsupported/7, with converter/owned-cleanup diagnostics; secondary cleanup4 keeps its primary reason. Reduced-cap units verify snapshot/raw-tree/graph bounds, not production-scale 512-MiB/100,000-page stress. Independent copied-page/catalog/action and inherited-resource/geometry regressions preserve their initial failures. Writer fidelity requires captured inherited fonts/boxes/rotation and immutable source. Original C2 full01 remains FAILED/native 1 with 18/20 passes and its outer workspace retained. The policy oracle incorrectly forbade the existing deep-outline outline_limit warning despite correct invalid_outline/6/no chapter or plan. C3 corrects only that validator and its synthetic units. The process failure identifies PS51 detached-pipe missing its authored PID receipt; the underlying cause and actual case status remain unknown. Fresh C3 process01 passes 35 controls. Canonical coverage selects 17 unchanged C2 stages plus 3 fresh C3 routes; no successful standalone full run is claimed. Historical T01 full/process failure and mutable-source Python trial remain failed as recorded. Human M3-T01 stays closed; CI and exact package remain separate gates.

Rejected deep-outline receipts permit exactly one existing `outline_limit` warning with null source order, depth 65 and message `The outline tree exceeds the traversal limit.` All other rejected policy fixtures require an empty warning list. The focused synthetic mutations reject wrong code, depth, order, text, missing/duplicate warnings and unexpected warnings on unsupported fixtures.

## M4-T03 fault and independent partition route

`--layer faults` requires both actual supported hosts and runs existing paths/process/outcomes plus the new fault stage; full includes that stage once as 21. See tests/faults/README.md and [T03 evidence](evidence/M4-T03-fault-regression.md). The common source map must bind every child before AC-075 is satisfied; four source-binding controls alone are insufficient.

AC-073 combines two prior publications, two actual simultaneously held marked stages, a real partial second-write failure 6 and a later successful repeat. It verifies immutable source/neighbor/prior hashes and identities, the held successful stage after failure, full reopened physical-page coverage and exact ownership cleanup. Actual venv launcher/interpreter IDs are separated and bound by boot/retained handle/exact job membership/native exits/jobzero/EOF. Every recovery attempts all known owned jobs and preserves unproved shutdown evidence.

AC-074 uses independent per-page ownership seed 20261010074: 720 accepted plans, 160 invalid manual inputs, 96 malformed outline rejections and 12 callable writer samples. Match exact ordered physical-page membership and L2 parent intervals; counts alone cannot pass omission+duplicate or crossing mutations. Preserve original source/neighbor/prior identity and reopened chapter/manifest bytes.

AC-075 combines both-host real/fake path/stream/exit/cancel suites with four declared replace/delete-after-capture adapter controls. Preserve exact saved engine bytes and original captured markers/hash. Keep engine API return, intended SystemExit code, actual worker/direct venv parent and OS-native supervisor status separate. Native descendants must stop and both streams complete before authenticated cleanup. Source/neighbor/prior files remain immutable; only authored processing copies are mutated.

Failed process/fault receipts retain exact partial raw evidence, current/completed attempts and outer ancestors. Missing/malformed/unexpected-validator child reports fail closed, never authorize ancestor removal. Synthetic oracle/recovery controls and agent PNG inspection do not establish native or human acceptance. Preserve all failed/intermediate and historical reports. Physical disk exhaustion, crash/power loss, arbitrary hostile-document/same-account sandbox, CI and exact release-package gates remain separately recorded. Human M3-T01 remains closed.


## M4-T04 real ebook acceptance route

`--layer ebooks` runs inherited conversion and the new authored EPUB3/genuine AZW3 companion, exactly2 stages. Supply explicit actual PS5.1/PS7, Calibre9.15.0 and developer Poppler26.07.0 paths. Full now contains22 stages, the companion once. Persistent reference PDFs/every-page PNGs stay under the parent report's external `-ebook-renders` directory. Missing/rejected child receipts preserve owning workspaces. See tests/ebooks/README.md and docs/EBOOK_SUPPORT.md.

Clean C **5214bb230590b27d9399d22b80c1f0b5fead2d74**, tree **5adb04ede4f8a852fc3dbe0a60858d9040be1d0e**, passes 453 Python methods (zero skips) and both native ebook stages, with all155 raw mapped files equal to Git, digest **4709450a19fe13bbdd121b7f4c73eb4adf94a2d94f5d293050b5bbf863b9e418**. All15 production paths remain identical to merged PR24 main. Four normal Auto1/Keep ebook cases cover EPUB3/AZW3four physical pages; four real malformed failures return app4 before planning/publication; two loopback EPUB controls observe actual image blocking and zero exact image/CSS endpoint requests bracketed by four positive GETs. No OS-wide network denial claim. Every generated/chapter physical page reopens and matches text/content/geometry/same-renderer pixels; prior same-base publications and immutable input/neighbor/file identities survive authenticated cleanup. The inherited stage freshly covers eight native Manual/retained/default executions, two separate callable preview/capture parity controls with declined working-PDF opening, eight fake-invalid converter outputs and one BAT missing-converter control. No application Preview/plan-only or viewer interaction is claimed.

[T04 evidence](evidence/M4-T04-ebook-conversion.md) and its [machine record](evidence/M4-T04-ebook-conversion.json) record actual commands/runtimes/seals and retained failed new-oracle trials. Source/native/public agent reviews and image inspection remain distinct from human/GUI tests. No fresh full22 or Pester is claimed: historical T03 full21/21/112Pester per host remains separate. Human M3-T01 stays absolutely closed; historical unknown process causes remain UNKNOWN. M4-T05 owns actual Windows CI/review/merge. No package or release closure follows here.

## M4-T05 actual Windows CI and local composite

The pinned least-privilege Windows Server 2025 workflow uses the same tests/run_tests.py for Python and separate explicit PS5.1/PS7 syntax/scoped-analysis/Pester. [C2 branch](https://github.com/PikkuJanne/WinBookSplit/actions/runs/38063779891) and [C2 PR](https://github.com/PikkuJanne/WinBookSplit/actions/runs/38063782954) are actual successful workflows, 475 Python methods and 112/112 Pester per host with zero skips. Reports bind all 161 mapped bytes to C2 **123bb3e9bb1ad6e15e3bf6e060f35fc8a7f15cb4**, digest **268e6530bf3ce4a7c72cef6bbc4587802d9d357dbb7b91cb689f92beb3dd778f**; PR actual synthetic merge is recorded separately. Only five declared JSON reports are uploaded, and exact downloaded artifact/API digests match. Package-input checks cover 15 runtime + 8 support prerequisites without building a package. Ordinary push/PR/manual CI has contents:read and no release or secret context; exact tools/actions are reviewed in tools/ci/README.md.

Clean C1 **6a1b0f37c505b95c20f3ee903a71f4fb6c30cb25** freshly passed 22/22 local full stages, native 0, 475 Python methods/112 Pester checks per local host, real Calibre 9.15/Poppler 26.07/PDFium, source/input stable and owned cleanup. C2 changes only one setup assignment in tests/python/test_output.py; 160 mapped paths and all 15 runtime files are unchanged. Final local acceptance explicitly retains 21 unchanged native stages plus fresh 475 Python methods on C2/native 0. This is not a standalone C2 full run. The same two inherited tests first failed on hosted C1 and equivalent long local TEMP; all 27 output transaction methods then passed under that TEMP after only shortening exclusive case-directory names. Production path limits, all guard/ownership/junction assertions and explicit AC039 long-path controls remain unchanged.

Cumulative review and actual inert serialized reproduction led to the narrow indirect feature-name correction; final focused 22 policy + 11 fidelity methods and six retained native inputs pass their explicit reject/drop oracles. Four retained projected-appearance PNGs have separate agent inspection; blank content is not general text/layout fidelity. Full fidelity separately verifies established rich/image fixtures. Both first hosted workflows and intermediate failed setup/policy/helper/audit trials remain immutable. Historical unknown causes stay UNKNOWN. Human testing stays closed. [T05 evidence](evidence/M4-T05-windows-ci.md) and [machine record](evidence/M4-T05-windows-ci.json) state source, exact raw seals, runtime distinctions, composite scope and ordinary PR26/main verification. No package/signing/tag/download/public-release or new clean-machine/GUI/network-isolation claim follows.

## M5-T01 — source-candidate documentation acceptance

AC-082/083 actual evidence is in evidence/M5-T01-documentation.md/json. FinalC2 has 171 exact mapped Git/raw files and16 runtime inputs. Fifteen distinct current user example origins (README 7,SETUP 2,support 1,help 5) are separately executed on both hosts from explicit C2 source archives:30 native-0 cases, two fresh runtime-only pinned venvs, 14 publications/42 chapters/four previews and ten resealed agent PNG views. Literal bodies/path substitutions/wrapper/native streams/plan-execution and manifest equality/page identity/Producer/startbookmark/version/privacy/immutability/owned-conversion cleanup are checked. Archive SHA/member/Git checks precede extraction; the actual native Get-FileHash example separately runs later from the extracted candidate. The converter template is bound by actual EPUB/AZW3 tablet reference calls; historical/developer/future-publishing recipes are outside the end-user inventory. No public release ZIP is tested here.

Fresh unchanged C2 shared --_python-child suite passed 483 methods/zero skips/native 0 under an external 600-second cap (281.44suite seconds), with source/sentinel unchanged and workspace retained. This does not claim a successful local outer Python-runner retry. Fresh shared shell passed116 Pester each actual PS 5.1.26100.9444/PS 7.6.5/zero skips plus syntax/scoped analysis; legacy application analyzer findings stay observations. Retained original C1 CLI 72/support 66/fidelity 10 native passes are explicitly bound to170 unchanged mapped paths/all 16 runtime inputs and applicable byte-identical validators; no standalone full 22 rerun is claimed. Focused VERSION5/checker15/helper-free native8 retained passes have exact applicable source bindings. C2 branch 38072286908/PR 38072289461 CI actually passed 483 Python/116 Pester per host on Server 2025/PS 5.1.26100.33438/PS 7.6.6; exact five JSON artifacts/API ZIP digests/logs/source maps reviewed. This hosted scope remains separate from local real Calibre 9.15/Poppler/PDFium/source-candidate execution.

First focused checker 14/15 failure, first documented0 case signature-module failure, C1 Python 300-second deadline/child124, C1 local and hosted one-warning scaffold failures, and intermediate pure-auditor assumptions remain failed/retained. Corrections are the explicit VERSION closure edge/negative guards, external process-local inbox module resolution, one test script-scope assignment with assertions unchanged, and pure reviewer parsing/label fixes. Actual first failure root causes are stated only where reproduced; deadline cause beyond the recorded limit is not inferred. No private input/raw local profile data is published. Human M3-T01 is absolutely closed; no new human/GUI/clean OS/full 22/package/signing/tag/release/anonymous-download/network-isolation claim.

## M5-T05 cumulative acceptance review

[evidence/M5-T05-acceptance.md](evidence/M5-T05-acceptance.md) and its machine matrix audit all17 improvements/92 prior required IDs and AC093/094. Reviewed merged main c7aebc53d47d75b06a5ee0fa89be5bec948fb374 retains all176 exact guarded C3 bytes/digest a428c67c95acf794a15767975079d34a79ca2a5b25af38f951cafa04d6d2715b. Fresh read-only saved full revalidation accepts22/22 original native0 stages and21 strict child receipts. No new full/app/converter/renderer/human run occurs. Actual merged-main CI38096308762 five downloaded reports/API digest/source audit passes514 Python/116 Pester per hosted host/zero skips and configured static gates, distinct from71 hosted default observations each. No runtime/build/dependency/test bytes change; final evidence own CI/sync follows externally after commit. Checkout newline/stale-stat/incorrect-plan-flag preparation failures stay retained. Closed M3/T03 D13 routes and all old failed/UNKNOWN limits are preserved; M6 exact final asset/source/publication/download gates remain pending.

## M6-T01 release source freeze audit

AC095/096 use a read-only audit of committed main **73c90d10dc16b1c321737afbdfbe959d61a24573**, tree **62da4fe06b2d07374b27743dd252aab492993c35**, VERSION1.0.0. Correct check_sync/native0 proves clean/fresh live main at the recorded freeze timestamp; live branch/PR/tag/release state is separately inspected. Committed source checker/native0 and independent builder `_source`/Git-byte audit freeze28 payload/all16 runtime/three builder rows,29 prerequisite inputs and27 shipped relative links. All176 guarded files digest **a428c67c95acf794a15767975079d34a79ca2a5b25af38f951cafa04d6d2715b** equal actually tested C3. No build or final-asset claim is made.

Independent prerequisite review audits94 prior IDs,65 tracked evidence seals and16 actual primary receipts. Fresh download of R CI38097123491 validates the exact five reports/archive API digest, native0/Python514/Pester116 each/zero skips, source API/Git maps and scoped static gates. Immutable CI auditor historical label and reported bootstrap-map-only scope are explicit. Read-only retained full audit revalidates22 actual C3 stages/native0 and21 strict children; it does not invoke a fresh application/full/Calibre/renderer/human run. Completed human routes stay closed; exact final R ZIP acceptance remains M6-T03. Preparation failures and all earlier skips/failures/UNKNOWN/security/developer limits remain recorded. See [freeze evidence](evidence/M6-T01-release-source.md) and [machine proof](evidence/M6-T01-release-source.json).

## M6-T02 final package integrity

AC097/098 use two actual explicit-R builds and two verifies from **73c90d10dc16b1c321737afbdfbe959d61a24573** with exact committed tools, absent external final-R-01/repeat-R-02 destinations and preserved separate command/native streams. Allthree asset bytes match across builds; independent exact ZIP/member/Git/mode/metadata/manifest/recipe/dependency/checksum audit passes. Final ZIP414497bytes/SHA e78a1d0b976223832fcff1c7f17d0d53473112324b05eeec3ec7f74e575ccdc3; manifest11547bytes/SHA ea5c29faaaeda9132ae75546e1fe0751d8f6417b109b23047c7590875a2f182f; sums178bytes/SHA 1bdff00a4589049e41901417784d58b43b31371ca2b74cd98f14a5c4f535548b. Focused unchanged real-Git test_release_package.py passes25methods/zero skips/native0; all176 source paths and final asset hashes remain unchanged after tests. Fresh R/prior-main CI artifacts independently bind514Python/116Pester per host/zero skips and scoped static gates; no new full or workflow invocation. Actual hashes/commands/runtimes/sealed audits and preparation failures are in [M6-T02 evidence](evidence/M6-T02-final-assets.md) and [machine record](evidence/M6-T02-final-assets.json). Package integrity is distinct from M6-T03 exact final-ZIP Windows/Explorer/real Calibre/render/human acceptance; closed prior human routes and historical limits stay preserved.
