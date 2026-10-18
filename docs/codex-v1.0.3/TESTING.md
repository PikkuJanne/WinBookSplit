# Testing: prove the route, not just the algorithm

For the separately authorized 18 October security repair, maintain the three
small array/dictionary/content parsing regressions in test_release_security.py:
boundary positives, reduced-limit rejection and default restoration. The old
6.20.0 environment actually lacks these three fields; all eight security/BAT
tests pass under new isolated 3.13.16/3.14.8/3.15.0 environments with 6.21.1.
Keep full supported-runtime, PDF, Calibre, both-shell, origin/hash/setup and CI
evidence source-bound. Changed package bytes require a new freeze and exact-ZIP
native acceptance; original R's accepted T07 is preserved only for its own bytes.
See tasks/08-T08-security-repair.md and implementation/correction/final candidate
evidence. OWNER_SECURITY_REPAIR.md identifies the new exact-ZIP handoff.

Original native result: T07 STORE-037 through STORE-042 PASS for the unchanged
actual owner Windows10 Store return at frozen R/exact production ZIP. Review and
offline checker results are in evidence/T07-owner-return.md/JSON and independent
audit JSON. The owner need not repeat that unchanged test. The new security
candidate requires fresh exact-byte Store evidence, currently NOT_RUN. Actual
enforced-policy STORE-024 remains NOT_RUN; T08 publication stays BLOCKED pending
the real required gates. The T02-T06 observations below describe their historical
component/tooling scope; they do not retract or substitute for the T07 result.

T06 freezes committed source R and exact assets, with one generated owner
launch/report action in `OWNER_KIT.md`. Its default command must fail on a
non-target OS, and its separately flagged Windows11 self-test must block before
fallback/extraction/bootstrap when no compatible existing Store runtime is
verified. Check all seven returned members, six envelope hashes and run/source/
ZIP correlation independently; the offline checker must still refuse a blocked
report. Test execution and raw streams stay distinct from sanitized evidence.

The exact-ZIP developer PDF rehearsal uses explicit offline setup followed by
automatic healthy no-op/read-only Check. It verifies all four page recipes and
the local authored MOTW/RemoteSigned/remedy component with payload data leases
held. It does not establish fresh automatic Store setup or enforced organization
policy. Windows11 and authored Calibre EPUB/AZW3 characterization are separate.

An actual T06 full conversion run exposed the official Manager's informational
welcome-cache write in a redirected missing-converter fixture. The repaired
test-only nested cache schema2 preserves the empty Windows cache chain and
held identities, and admits only the bounded ordinary single-link welcome marker
whose exact bytes match a separate same-run verified installed-only Manager
listing. All other paths/links/content/provenance fail. This observation is never
application/setup selection input. Historical reports and the legacy native
report schema stay unchanged. Broad R1 execution remains labelled R1; explicit
Git-blob equality for all production/kit helpers, final delta checks, full
conversion and final-source CI document the subsequent source freeze honestly.

T05 adds a separate store-existing route, schema2 no-substitution/attempt gate,
bounded UTF8 collector, actual negative launch controls with positive controls,
failed diagnostic return ZIPs and new automatic application PDF recipes. The
normal Python suite maintains collector PS5.1/PS7 pairs and Store scenario probes;
shell syntax/analysis includes the new tools/probes. Three separate supported
explicit runtimes retain policy/origin/package/PDF coverage. See evidence/T05.md
for counts, preserved RED/exploratory failures, skips and source bindings. Broad
suite bytes and the final narrow failure-receipt/README delta are recorded
separately, with focused reruns/analysis for that delta. The actual new PDF recipe
rehearsal used explicit offline developer setup and automatic healthy no-op.
No actual existing-Store/fresh automatic setup/owner exact-ZIP result is inferred.

T04 adds maintained rejection-category, coherent summary, actual emitted setup
Action/BAT, no-op/repair and MOTW/process-policy/privacy cases. These use owned
explicit developer fixtures separately from native automatic discovery. The
application-summary Action is executed unchanged against the healthy authoritative
fixture venv after offline setup, not as fresh online application acceptance.
Actual Windows11/python.org automatic processing and setup readiness consistently
report the unchanged verified 3.14.7 security-floor rejection on PS5.1/PS7. See
`evidence/T04.md` and source-bound receipts. Enforced organization-policy denial,
native Store and owner Windows10 exact-ZIP evidence remain NOT_RUN/T07. The full
shell test controller is bounded at 600 seconds per suite; production inner
probe/setup deadlines are unchanged. Policy comparisons use the same host edition
before/after; PS7 and inbox WindowsPowerShell persistent stores can differ.

T03 extends maintained `Discovery.Tests.ps1` with descriptor-to-concrete-child
regressions, strict manager line-format negatives, distinct empty/rejected/failed
outcomes, optional-failure continuity, authority and a disposable real package-free
interpreter. That interpreter is created only in an owned developer fixture with
`venv --without-pip`; setup-base success, processing `missing_pypdf` identity and
the existing setup summary are checked separately. This is not the owner scenario.
Actual Windows 11/python.org automatic discovery on both shells now reaches its
existing concrete 3.14.7, verifies identity and rejects the reviewed 3.14.8 floor.
No compatible native success or Store-channel launch is claimed by those results.
See `evidence/T03.md` and its hash-bound receipts. Native Store/owner exact-ZIP
requirements remain deferred to T07; T04 still owns full diagnostic/remedy UX.

T02 adds maintained `tests/powershell/Manager.Tests.ps1` to the normal shell
runner. It executes the emitted query against labelled synthetic inventories,
checks both identity tuples and malformed/cross-paired negatives, and inspects
real authored ordinary/hardlink/junction decoys without executing them.
`Process.Tests.ps1` adds an actual marker-child positive control and proves a
package-bound rejection prevents that child from running. T01 reproduction
receipts/tests remain historical old-behavior evidence; their bug-expecting
assertions are not part of the maintained normal runner. T02 real Windows 11
python.org component receipts are separate from those synthetic tests and from
the still-NOT_RUN owner Windows 10 exact-ZIP Store gate. See `evidence/T02.md`.

## Layers

1. Reproduction and unit/contract tests: replay sanitized records, original query
   exclusion, downstream identity rejection, and misleading summary. Before the
   patch the Store-positive regression must FAIL for the expected reason. A
   source-text check alone is not the behavioral regression.
2. Component Windows tests: official identity pairing, manager entry-point
   binding, real alias launch under the supervisor, path/encoding/parser limits,
   explicit/venv authority and negative decoys. Synthetic records are synthetic.
3. Existing application regression: retain split-plan, PDF safety, subprocess,
   setup, diagnostics, BAT and dependency-origin tests. Test actual PS5.1 and PS7;
   preserve supported Python branches using available reviewed environments.
4. Native release-archive test: the existing Store installation on the owner's
   Windows 10 desktop, through normal automatic setup and BAT, is mandatory.
   Server CI, Windows 11, mocked registrations, or supplied runtime paths do not
   satisfy it. A short Win11 no-regression run is also required.

## Reuse, do not recreate the testing platform

The current runner is `tests/run_tests.py`; `tests/runtime/characterize_runtime.py`
and `validate_runtime_report.py` exercise actual processes. Inspect Pester tests
and `tools/compatibility/Invoke-NativeAcceptance.ps1`/`native_smoke.py` before
adding new infrastructure. Use their actual current CLI/help, not guessed flags.
The runtime README's current invocation requires an explicit developer Python;
that is fine for developer tests, NEVER for the owner automatic-discovery gate.

A developer dependency environment can differ from the runtime under test. Do
not add renderers/Pester/dev libraries to the user's runtime to make tests pass.
Run existing ebook smoke with authored inputs and already-approved Calibre; PDF
acceptance must not depend on Calibre. No copyrighted/private book in fixtures.

## Required regression families

`fixtures/registration-cases.json` supplies identity test data. Add invocation
and rejection cases from `ACCEPTANCE_CASES.json`, including mismatched publisher,
wrong signature, empty registration vs rejected registration, malicious alias
binding, unrelated CWD, malformed listings, no installed runtime, multiple
runtimes, healthy/broken venv, and a supported base without the locked library.

A mocked nonzero or timeout needs a positive control showing the stub is really
used. Do not replace failing assertions with weaker ones. Report actual commands,
source commit/tree, host/version, outcomes, skips and preserved negative results.
Missing capability means NOT_RUN/BLOCKED. No repeated manual fidelity programme.

## Helper validation is NOT application acceptance

`check_handoff.py` checks this handoff's files/task links. The separate
`check_native_store_report.py` enforces the narrow required declarations and
hash-bound artifact set; it does not independently attest Windows execution or
make a fabricated report trustworthy. Independently review native receipts.
The supplied native template intentionally fails until real results are filled
by the automated kit. Never make it pass by hand-editing fields.

## Deferred native checks without repeated owner requests

Task dependencies order implementation checkpoints, not repeated owner sessions.
If code/tests can advance but a native case needs the owner's desktop, record
IMPLEMENTED_PENDING_NATIVE and the exact deferred acceptance ID, then proceed
with independent next work. Do not mark the case PASS or the task fully COMPLETE.
Consolidate owner-only cases into T07's single kit. All required cases must pass
before T08 publication. Use an available authorized agent Windows environment
for early alias/supervisor tests whenever possible; do not wait to discover
obvious launch errors on the owner desktop.
