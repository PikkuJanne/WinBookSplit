# Checkpoint 2 — readable PDF authentication and local opening views

30 October 2026. PDF repair implemented; final application-unit run passed.
Checkpoint 3 setup/consent/resume work has not started.

## Authority and source

Read root AGENTS.md, STATUS_NEXT.md, CHECKPOINT_1.md, the historical v1.0.3
closure/testing references and the supplied repair pack's checkpoint-2 contract.
The instruction ZIP still has SHA256
`f33ab17eab774f4a65cf6e008949012e810cb0c10ce4c9bb9796f0d6524dc70d`.
Only the requested checkpoint is authorized here; there is no publication.

Started on clean `codex/functional-repair`, local HEAD and freshly queried live
origin branch both `bf8dff4725ce748f50bac466d117957264450d19`, tree
`45a62949ed8ac3ff02d95f3159942d9a273f38b9`. Origin remains
`https://github.com/PikkuJanne/WinBookSplit.git`; there is one worktree.
The final changed production/test byte map is recorded in CHECKPOINT_2.json.
Its canonical map SHA256 is
`e2570a1a3347ecbc1ca8e4ae2d3335e9900d6e9e40c945f18b8de64a7c3c02ff`.
Documentation is excluded from that map to avoid a self-referential hash.

## Implemented behavior

- Removed the private blanket `_handle_encryption` override. The supported
  pinned pypdf constructor makes its ordinary single empty-password attempt.
  Shared catalog access recognizes success without treating `is_encrypted` as
  an access verdict or reauthenticating a captured reader. Tests count the
  library attempt and forbid another attempt during captured-reader reuse.
- Genuine nonempty-password inputs return `password_required` / exit 7.
  Unsupported handlers return `unsupported_encryption` / exit 7; unavailable
  crypto backends return `encryption_backend_missing` / exit 3. Malformed
  encryption representations return `invalid_document` / exit 6; converted
  malformed output retains `conversion_output_invalid` / exit 4. Structural
  checks distinguish truncated credentials/invalid key parameters from an
  ordinary failed password attempt, including valid indirect numeric fields.
- Validated local destination arrays and exact local GoTo dictionaries can be
  ignored. Local physical-page membership, fit/arity/parameter types, finite
  coordinates, reference bounds, name-tree limits and unambiguous names are
  checked. Only the validated catalog edge is excluded from generic traversal;
  the same action referenced from copied resources is still rejected.
- Accepted views emit one `opening_destination_ignored` warning per plan.
  Chapter writers create fresh catalogs without OpenAction. JavaScript,
  Launch, URI, remote/embedded GoTo, Next chains, extra action fields,
  malformed destinations and unrelated active-content checks remain rejected.
- Preview, execution, interactive held plans, captured readers and converted
  PDF validation use the same policy. The outcome map, strict PowerShell
  result parser and redacted support warning/export contracts agree.

Encryption access does not establish document permissions. This repair adds no
per-operation permission-enforcement feature or nonempty-password entry feature.
The source is never decrypted/resaved on disk. New chapter PDFs do not preserve
source encryption or the ignored viewer setting. No AES interoperability claim
is made: the unavailable-backend negative is a labelled synthetic exception,
not execution of an AES algorithm with a newly installed backend.

## Executed checks

Existing developer driver: regular Windows x64 CPython 3.14.8 and locked pypdf
6.21.3 on Windows 11 build 26300. No interpreter, crypto dependency or application
environment was installed or prepared. This driver was not injected into normal
discovery. Pester 6.2.0 and PSScriptAnalyzer 1.25.0 were already available.

| Check | Actual result and scope |
| --- | --- |
| Final fixed application-unit roster | PASS, 171 methods in 111.377 seconds; zero failures, errors, skips or expected failures. All ten required modules loaded under `-I -B`, including final-byte policy (31), functional PDF repair (11), conversion (18), fidelity, interactive, outcomes, compatibility, filename, version and parser-warning tests. Final source-map equality was checked after execution. |
| Focused functional PDF suite | PASS, 11 methods. Manual/Auto1/parent-aware Auto2 preview and execution, independently expected physical-page ranges/order/content, unchanged source and prior output, interactive confirmation, converted validation and both encrypted captured-reader variants. Included again in the final roster. |
| Focused PDF policy suite | PASS, 31 methods on the reviewed implementation before the final name-tree key comparison refinement. RC4-40/RC4-128 positives, local direct/indirect/named views, single authentication attempt, genuine password refusal, handler/backend/corrupt negatives and active/malformed controls. Final-byte coverage comes from the final roster. |
| Conversion helper suite | PASS, 18 methods. Actual supervised authored fake-converter children, including encrypted empty-password output acceptance, stream/process/cleanup failures and source preservation. This is not a new real-Calibre run. Included again in the final roster. |
| Wider planning/output/security/report checks | PASS, 223 methods at the initial implementation before the final narrow PDF validation refinements. Manual/Level1/Level2/plan/preview/output/diagnostics/parser-security/support and policy receipts/runner/CI-report checks. No missing native scenario is inferred from receipt tests. |
| PowerShell Outcome + Support components | PASS, 56 methods on each actual host: inbox 5.1.26100.9711 and existing bundled PowerShell 7.6.5. No failed or skipped cases. New error exits/status/output-claim negatives and warning redaction included. |
| Configured PowerShell analysis | PASS, zero findings across the four changed PS files, using the repository settings and analyzer 1.25.0. |
| Independent review | Final authentication, local-page/action-edge handling, bounds and malformed encryption classification reviewed; reported defects repaired. |
| Native fresh first-run / Store / exact distribution / private book | NOT_RUN here. Setup is checkpoint 3. Checkpoint 1's natural Python 3.14.7 security-floor prerequisite remains unresolved; no substitute runtime or prebuilt environment was supplied. Private book unavailable. No package, publication or native Windows 10 pass is claimed. |

The first broad roster run is retained as FAIL: 171 methods, three subtest/test
failures and no errors. Two raw indirect-alias negatives exposed a pypdf parser
AttributeError escaping the opening-reference classification; those are now
narrowly rejected as unsupported opening structures. The third was the existing
LF-only VERSION byte assertion against an unchanged CRLF Windows checkout.
That test now accepts only the parser's existing two line endings for exactly
version 1.0.3. Neither VERSION nor production version parsing changed.
Earlier intermediate receipts are not relabelled as final-byte runs.

## Commands and receipts

`$Repo` is this checkout; `$DevPython` is the existing
`%TEMP%\WBS-623-finish-q6dwc7hk\dev14\Scripts\python.exe`.
Private/local raw output is under `%TEMP%\WBS-functional-repair-checkpoint2-20261030`,
`%TEMP%\WBS-functional-cp2-agent-20261030` and
`%TEMP%\WBS-functional-repair-checkpoint2-protocol`.
CHECKPOINT_2.json retains raw receipt hashes and the tested byte identities.

```powershell
& $DevPython -I -B "$Repo\tests\python\test_functional_pdf_repair.py"
& $DevPython -I -B "$Repo\tests\python\test_pdf_policy.py"
& $DevPython -I -B "$Repo\tests\python\test_conversion.py"
& $DevPython -I -B "$Repo\tests\run_tests.py" --_application-units-child
```

The wider 223-method command loads these explicit unittest modules from the
trusted test directory: `test_manual`, `test_level1`, `test_level2`, `test_plan`,
`test_preview`, `test_output`, `test_diagnostics`, `test_release_security`,
`test_support_receipt`, `test_document_policy_receipt`, `test_runner`,
`test_ci_reports`. It uses `-I -B`, unittest.TextTestRunner and exits nonzero
unless all tests succeed.

The retained `checkpoint2-protocol\run.ps1` component driver imports the existing
pinned Pester module, passes repository/work/test-Python parameters to
New-PesterContainer for Outcome.Tests.ps1 and Support.Tests.ps1, and runs on
both hosts with `-NoProfile -ExecutionPolicy Bypass -File`. Policy handling is
process-only; this does not prove organization-policy or ordinary BAT acceptance.
The project-settings analyzer result is authoritative; a preliminary wrong
`-LiteralPath` invocation was corrected to `-Path`, and exploratory default-rule
findings are retained separately rather than presented as the configured gate.

Setup, discovery, locks/floors, native test prerequisites, published tags/assets
and historical evidence are unchanged. Only authored fixtures were processed.
After final verification, commit/push intended paths and verify clean local HEAD
against the freshly queried live branch. Report that synchronization receipt in
the handover, then stop; checkpoint 3 is the next separate chat.
