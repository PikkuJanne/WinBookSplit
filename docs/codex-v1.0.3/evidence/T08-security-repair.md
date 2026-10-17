# T08 - pypdf security implementation checkpoint

The owner explicitly authorized this separate repair on 18 October 2026. Source
implementation is complete and independently reviewed. Final candidate freezing,
CI, packaging and transferable native-kit receipts follow separately; this record
is not release clearance. Initial HEAD was b2ab77ed4fb28986d3ab9798c6de788af50476ec,
clean and equal to fresh live origin on codex/v1.0.3-store-discovery.

The exact lock is now pypdf 6.21.1, SHA256
553e2d42f5bb3bb2a785a7d3de8ef90bb7df74b5413eff5a7b4388007ea9254b.
Requirements, compatibility policy 4/review date 2026-10-18, runtime test lock,
package-input/kit pins, current tooling assertions, mutable tests and user docs
agree. Python 3.13.16/3.14.8/3.15.0 floors, Calibre 9.15.0, engine, import/origin,
setup, process and file safety implementations are unchanged. Advisory sources
and limited security claims are in SECURITY.md and the candidate review.

Three maintained parser guard tests include small boundary-positive controls,
reduced-limit failures and default restoration. Against the preserved old 6.20.0
runtime, the test file ran eight cases: five passed and three errored because the new limit
fields are absent. This is the expected bounded RED, not a memory exploit.
All8 security/filter/BAT tests pass on newly created external isolated developer
3.13.16/3.14.8/3.15.0 environments with 6.21.1; pip check passes in each.
Existing installations/venvs were not modified. New environments are separate
explicit developer tests, never inputs to an owner Store scenario.

The full application matrix passes on all three targets: 150 runtime-only unit tests
per target, 10 PDF/conversion cases, 21 native probe records and both PS5.1 and
PS 7.6.6. Existing verified Calibre 9.15.0 performs actual authored EPUB/AZW3 checks.
Each matrix has source_unchanged=true, exit 0 and empty outer stderr. These actual
Windows11 explicit-runtime regressions do not prove automatic Store discovery or
fresh native acceptance. Three full raw reports stay external with byte hashes.

Shared Python and PowerShell suites were started before the last reviewed
runtime-characterization assertion/tooling README correction. Their final raw
results and source-change limits will be recorded in the follow-up candidate
receipt, along with CI on the exact committed source. No pending result is a pass.
Independent review found one stale 6.20 assertion, now policy-driven; its recheck
found no further finding. Current source pins are frozen only after intended
source review and are recorded externally, without a self-referential commit.

Original R 243d31e, production ZIP, kit/carrier, OWNER_KIT/OWNER_TRANSFER contracts,
historical reports and accepted T07 remain untouched. T07 PASS applies only to
old R/exact ZIP. The security repair creates changed candidate bytes requiring a
new freeze, new package and new exact-ZIP Windows10 Store acceptance. STORE-024
actual enforced-policy denial remains NOT_RUN; no target or receipt was supplied.
No policy/PATH/alias/runtime-Manager changes, publish/tag/main merge or old-release
mutation occurred. T08 publication remains gated by the real missing evidence.

The six authored fixtures for the new kit were copied with exclusive creation
from the hash-verified original kit. All six pins match its binding manifest;
this is input provenance reuse, not another execution of the accepted Store test.
Final new package, owner instructions and ordinary push/live equality receipts
belong in the follow-up record and external handover. Stop after completing that
independent preparation checkpoint; do not manufacture native results.
