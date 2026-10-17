# Testing: prove the route, not just the algorithm

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
