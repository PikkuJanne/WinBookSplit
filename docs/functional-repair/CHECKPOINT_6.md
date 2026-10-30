# Checkpoint 6 — agent-only document-policy verification

30 October 2026. The migrated native document-policy route passes. Production
files, dependency locks/floors, VERSION and existing packages/releases are unchanged.
This checkpoint performs no publication. The owner's direct continuation includes
eventual publication and explicitly excludes all human testing, sign-off, owner
reports and kit returns. Unavailable environments remain NOT_RUN, with no owner
action required; they are not invented passes or pending human work.

## Source and authority

Read current AGENTS.md, STATUS_NEXT.md, checkpoint-5 Markdown/JSON and its linked
prior evidence, then the actual fixture/producer/validator/support/runner paths.
The latest supplied AGENTS replaces earlier AGENTS instructions. The owner's
separate direct follow-ups provide the no-human-testing order and eventual
publication authority; the instruction pack alone had not provided that authority.
One checkpoint per chat and normal commit/push/clean-live verification still apply.

Started clean on `codex/functional-repair`, HEAD and freshly queried live branch
both `15519da527e092c19be7e839a86f71893fdeba7d`, tree
`6dda8a7a7b6954507d615f7b7a1fd130528092de`. Origin remains
`https://github.com/PikkuJanne/WinBookSplit.git`; one worktree. An initial remote
query hit a host allocation failure; a fresh successful query established equality
before edits. No reset, fallback interpreter or configuration change was used.

Tests bind reviewed working bytes over that base, before evidence/status edits.
The native report independently binds 232 files; its outer/child/current map
digest is `0941251b1e8f48cf6db1de1bafbb56a50fe485266713e0fdd7d21ae5685ad957`.
CHECKPOINT_6.json records seven changed Python test/tool working hashes and their
staged Git blob identities, including any CRLF normalization, plus all 22 preserved
runtime members. The final checkpoint commit is verified after ordinary push in
the handover, avoiding a self-referential hash in evidence.

## Reproduction and repair

The original actual Windows native route failed after four successful controls:
`PS51-encrypted-user` correctly returned **error / password_required / 7**, while
the obsolete validator demanded unsupported_document. The failed native receipt
and authored workspace remain outside Git.

The test fixture roster still has 25 original generated inputs and the same
73 application controls plus two exports. Genuine password refusal is distinct
from unsupported features; empty-password input is explicitly retained as a
positive execution/preview case. Positive BAT input receives the full method,
starts, confirmation and completion replies. Preview requires its exact
preview/preview_complete/0 result, zero writes/execution/diagnostics, complete
typed physical ranges and captured source identity. Reopened chapter checks
require physical-page/content order, unencrypted output and no catalog OpenAction.
Rejected-input support exports now bind the genuinely password-required input.
Existing malformed, active-content, process/drainage, source/neighbor/prior-output,
ownership-limited cleanup and redaction checks remain enforced.

The shared support receipt oracle gains only the already implemented fixed
`opening_destination_ignored` navigation token. Its unsupported-status predicate
is unchanged. The runner's report description and active guidance match the new
contracts and owner authority. Historical docs/codex-v*, checkpoints 1–5, frozen
test ZIPs, the old kit, and published tags/assets are preserved.

An intermediate synthetic fixture incorrectly used unsupported/password_required;
its first unit run had 21 passes and one error. A speculative shared predicate
broadening then produced intermediate green results against that wrong fixture.
An actual PS5.1 export succeeded with error/password_required/7, disproving the
assumed production defect. The exploratory driver's expected-failure assertion
was wrong. We restored the shared predicate, corrected the fixtures to the real
status, and reran final units/native checks. All intermediate receipts remain
nonfinal evidence. **No production exporter change was made.**

## Actual results

Existing tools only: Windows 11 Pro build 26300, regular x64 CPython 3.14.8,
pypdf 6.21.3, Windows PowerShell 5.1.26100.9711 and PowerShell 7.6.6. Explicit
developer-runtime harnesses are labelled as such; no ordinary Store setup pass
is inferred. No Python, app environment, global package, PATH, alias or persistent
execution policy was provisioned or changed.

| Check | Result and scope |
| --- | --- |
| Original native harness | FAIL as reproduced above; four completed controls and specific actual password failure retained. |
| Final runner document-policy layer | **PASS**, 271.832s, all 75 actual controls: 30 PS5.1, 30 PS7 and 13 BAT application cases plus two redacted exports. Three empty-password execution routes retain all four physical pages in order with plain catalogs; two positive previews produce no output/log; five password routes return error/password_required/7 with no output. |
| Final document-policy receipt units | PASS, 22 methods, zero failures/errors/skips, 0.067s. Includes eight added methods for exact positive/negative outcomes, typed complete preview/source binding, no-write claims, reopened page/encryption/action negatives and refused-export binding. Synthetic oracle tests, not native executions. |
| Final support receipt units | PASS, 18 methods, zero failures/errors/skips, 0.097s. Actual error categories/exit mappings, forged unsupported aliases and fixed warning redaction. Synthetic receipts. |
| Focused runner/PDF regressions | PASS, 47 methods, zero failures/errors/skips, 39.263s: five document-policy runner integration controls, 31 existing PDF-policy methods and 11 existing functional PDF methods. Source hashes stable. No full-suite claim. |
| Independent completed-receipt audit | PASS. Reconstructed child raw SHA256, strict validator, all 75 controls, source/host maps, page/output/cleanup/export invariants and unchanged production independently verified. Audit launches no application or new native test. |
| Review and whitespace | PASS, independent code/guidance review and scoped diff check. No production diff. |

The runner JSON is 3,924,677 bytes, SHA256
`c57177ed7e12db6f794c4db3561f4c46ec3e6e67b2ce60c3807f6f13f3293ef6`.
Its embedded child raw SHA256 is
`aecb318d834a0fad4b999e6d5d0d88c4a10d40c725b61c2c905c9bd557161a33`.
Independent audit SHA256 is
`2f7426b0adf46a57eba0422914bce3d73301c3786ff40255a4fc08585373b737`.
The source map remains stable throughout the native and focused runs.

Ordinary Store first-run/setup/resume/repeat/network, Windows 10 and the private
book remain NOT_RUN for unavailable agent prerequisites. Human testing is
OUT_OF_SCOPE and no owner action is required. No new Calibre, full-suite,
supported-runtime matrix, new-package or publication pass is claimed here;
earlier evidence retains its actual original source/byte scopes.

## Commands, receipts and next checkpoint

Raw receipts and exact argv are under `%TEMP%\WBS-functional-cp6-20261030`.
The independent read-only audit is under
`%TEMP%\WBS-functional-cp6-independent-audit-0bbff1b1d1ea4da09d81bbac4c742fc1`.
Only their hashes/redacted locations enter Git. With the pre-existing developer
Python, actual host paths and a fresh external report:

```powershell
& $DevPython -I -B "$Repo\tests\run_tests.py" --layer document-policy --report $Report --shell-path $PS51 --shell-path $PS766
& $DevPython -I -B "$Repo\tests\python\test_document_policy_receipt.py"
& $DevPython -I -B "$Repo\tests\python\test_support_receipt.py"
```

The retained final-focused-checks.py loads the five named document-policy
RunnerTests methods plus test_pdf_policy and test_functional_pdf_repair through
unittest, with before/after source hashes. The independent audit revalidates the
completed raw report and source map; it does not rerun native application actions.

Live read-only release/tag inventory still contains v1.0.0 through v1.0.3; v1.0.4
is the observed next unused patch, to verify again at release preparation. The
next checkpoint coordinates maintained version/package/test pins, current release
guidance, automated CI and fresh source-bound package acceptance, then publication
of those exact bytes under the direct owner authority. No old v1.0.3 package/kit
may be relabelled as the new release or new Store acceptance. No human testing
or further publication confirmation is requested by this handoff.

Commit intended files only, push normally, verify clean local HEAD against the
freshly queried live branch, hand over and STOP at this checkpoint boundary.
