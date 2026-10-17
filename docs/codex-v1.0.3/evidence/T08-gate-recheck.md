# T08 gate recheck from checkpoint 64e726b

Status: **BLOCKED**. T07 remains COMPLETE for STORE-037 through STORE-042 at
frozen R/exact ZIP. STORE-024 remains NOT_RUN, T04 remains
IMPLEMENTED_PENDING_NATIVE, and STORE-043 through STORE-048 remain NOT_RUN.
No acceptance status changed and publication has not started.

This 17 October 2026 continuation began clean at
`64e726bccad204ad910610cdad08bac21e40bb32`, tree
`288f0cb09d33fd0cde1b6a6936a629d898f3dbb2`, equal to freshly queried live origin
on codex/v1.0.3-store-discovery. Root/ancestor/nested instruction inspection found
only the root AGENTS.md; origin and the single worktree are unchanged.

## Remaining policy gate

Independent review confirms that the existing Store kit cannot produce the
missing enforced-policy acceptance. Its local authored MOTW component initializes
organization_policy_denial as NOT_RUN in
`tools/compatibility/Invoke-StoreExistingAcceptance.ps1:228`; the separate checker
at `tools/codex-v1.0.3/check_native_store_report.py:375` requires that NOT_RUN
declaration and rejects a claimed local-component pass. This is an intentional
evidence boundary, not a reason to alter receipts or weaken the validator.

The accepted owner return records MachinePolicy/UserPolicy Undefined. No new
enforced-policy receipt or authorized target was supplied. Genuine separate
evidence from an already enforced environment is needed to clear STORE-024.
Do not rerun the accepted owner Store test, change desktop policy, or substitute
another launch route. Preserve the existing validated kit. Any necessary future
human check must first have one validated transferable launcher/report handoff;
no new enforced-policy test kit was prepared or validated in this review.

## Security disposition

An independent current official-source review found no material change from
T08-gate-review.md/JSON. Retained pypdf 6.20.0 includes the
[multiple-filter fix](https://github.com/py-pdf/pypdf/security/advisories/GHSA-7mh2-4gg2-979f),
but remains in the affected ranges for
[array/dictionary parsing](https://github.com/py-pdf/pypdf/security/advisories/GHSA-f54x-cmq8-qg57)
and [content-stream operations](https://github.com/py-pdf/pypdf/security/advisories/GHSA-cgf9-327m-v77c),
both naming 6.21.0 as fixed. Parser initialization and indirect-object resolution
still occur before application structure bounds. Process deadlines, stream
bounds and kill-on-close do not supply a heap limit.

The [form-field advisory](https://github.com/py-pdf/pypdf/security/advisories/GHSA-6vvc-x485-93v7)
still reports affected <6.21.0, patched field >=6.2.10, and prose fix 6.21.0;
that upstream discrepancy remains explicit. The
[image advisory](https://github.com/py-pdf/pypdf/security/advisories/GHSA-5hx5-vpf9-5wcv)
still names <6.21.1 affected. Lazy content-stream parsing and the absence of
direct production form/image API calls narrow the demonstrated paths without
proving immunity. Earlier recorded outline/tree residuals remain unchanged.
No exploit or memory-pressure test ran. No new dependency was selected or
security clearance granted. The release disposition remains unresolved/BLOCKED.

## Frozen evidence and preservation

Read-only standard-library archive/hash and raw Git comparisons reconfirmed:

- All 34 production ZIP members equal R and current HEAD; ZIP SHA256 remains
  `70a76bd04f5645e195972bf075ff53d651c9993696a40065241390805d2ee1de`.
- The unchanged 35-member original kit, 34 manifest pins, 21 helpers at R, six
  fixture pins and 35 existing extracted members agree. Kit SHA256 remains
  `e2f2a36a727edec378b891f5b125a18bea89dc8e41946719ac8b8a6d2e1837fc`;
  bindings remain `446ef20b5aad8462f413ab6c43df12baa8c6d46211e921c4933bea3c3e3e818b`.
- The unchanged 41-member convenience carrier contains that exact kit and its
  35 files; SHA256 remains
  `adc6d1cb89fc90123cd5ef9c0f71c121f74761b8205fea7a1c4c6947e8a97745`.
- The accepted owner ZIP is still 377606 bytes, SHA256
  `1ac523bfb8aebc1511c27629af5134fa9000a08db361a6e207c33280d184a06e`.
  All seven member pins match. A new external copy of those unchanged JSON bytes
  passed the checker read from raw R: exit0, CONSISTENT_NATIVE_REPORT, six
  receipts, empty stderr. The prior public 79-check audit is unchanged.

This is offline rechecking of previously accepted same-run native receipts;
there was no new native execution or local PDF reopening. No returned code was
executed. R remains `243d31ee03efac35dc6f7e88c0c1dfd1203081ed`, tree
`ea0db1cc5f5494b508f140deaf25eff87a7e39d9`, VERSION 1.0.3. Any necessary product,
dependency or package change needs a new committed freeze and new exact-byte
acceptance; this review cannot extend the accepted R result to changed bytes.

Fresh old tag objects/peeled commits and release body/asset metadata equal the
prior gate review; v1.0.3 is absent. Live main remains
`4cc6b00b99685607f3214e731edaaec7d65926df`. Existing final-R CI run39043915939 and
all four jobs still report success. These are existing scoped CI results and
metadata preservation, without a new CI run, anonymous byte download, or release
closure acceptance.

## Checkpoint

Only active routing/task/evidence links and this new evidence pair change.
Existing evidence, historical reports, owner contracts, product, package,
controller, fixtures and desktop installations remain unchanged. No app/setup,
runtime/Manager installation, policy/PATH/alias change, release/tag/main merge,
force Git or settings action occurred. Broad implementation suites were not
rerun for this documentation-only checkpoint.

Structural handoff, JSON/evidence links, all 48 unchanged acceptance statuses,
intended paths, privacy, whitespace and raw staged evidence bytes are checked
before commit. The companion JSON records commands, actual results and read-only
lookup failures. Final commit/push and clean HEAD equality with fresh live origin
belong in the external handover, without a self-referential evidence hash.
STOP after that checkpoint; T08 remains blocked by STORE-024 and unresolved
security disposition.
