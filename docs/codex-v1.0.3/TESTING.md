# Testing: prove the route, not just the algorithm

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
