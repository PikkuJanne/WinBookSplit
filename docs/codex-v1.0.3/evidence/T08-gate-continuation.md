# T08 gate continuation from checkpoint 4711c9e

Status: **BLOCKED**. This gate-only review changes no acceptance result. T07
remains COMPLETE for STORE-037 through STORE-042 at frozen R/exact ZIP. Actual
enforced organization-policy denial STORE-024 remains NOT_RUN (the acceptance
index retains DEFERRED_NATIVE), T04 remains IMPLEMENTED_PENDING_NATIVE, and
STORE-043 through STORE-048 remain NOT_RUN. Publication has not started.

The 17 October 2026 continuation started clean at
`4711c9e6156f532413e071a4d18c0d2958be99a2`, tree
`9b6c59bc660da800f871e956a6e7f6eb1587bba4`, equal to freshly queried live origin
on codex/v1.0.3-store-discovery. The expected origin and single worktree remain;
ancestor/nested inspection found only root AGENTS.md. Required active instructions,
task specifications and earlier evidence were reviewed with three independent
read-only audits of gates, security and frozen receipts.

## Policy evidence remains missing

Direct review of the accepted original owner ZIP reconfirms MachinePolicy and
UserPolicy Undefined before and after execution. The accessible Windows11 host
also has those scopes Undefined. No new enforced-policy receipt or authorized
enforced-policy target was supplied. The retained Store kit's local component
explicitly initializes organization_policy_denial NOT_RUN at
`tools/compatibility/Invoke-StoreExistingAcceptance.ps1:228`; the checker requires
that declaration at `tools/codex-v1.0.3/check_native_store_report.py:375`.
Rerunning the accepted Store test cannot clear STORE-024.

Genuine separate evidence from an already enforced environment is still required.
No new human test kit was prepared or validated, and no owner rerun is requested.
Preserve the existing validated carrier and kit. Any necessary future human check
must first have one validated transferable launcher/report handoff; changing the
desktop's policy or editing a receipt cannot supply the missing observation.

## Current security review

Fresh official-source review leaves retained pypdf 6.20.0's release disposition
unresolved/BLOCKED. The [multiple-filter advisory](https://github.com/py-pdf/pypdf/security/advisories/GHSA-7mh2-4gg2-979f)
fix is included, but the retained version remains affected by
[array/dictionary parsing](https://github.com/py-pdf/pypdf/security/advisories/GHSA-f54x-cmq8-qg57)
and [content operations](https://github.com/py-pdf/pypdf/security/advisories/GHSA-cgf9-327m-v77c),
both fixed in 6.21.0. The [6.21.0 release](https://github.com/py-pdf/pypdf/releases/tag/6.21.0)
and [PR4218](https://github.com/py-pdf/pypdf/pull/4218) corroborate parser limits.
PolicyPdfReader initializes pypdf before structural preflight, and indirect
resolution precedes application dictionary/array size checks. Process deadlines,
stream bounds and kill-on-close do not provide a heap limit.

The [form-field advisory](https://github.com/py-pdf/pypdf/security/advisories/GHSA-6vvc-x485-93v7)
still says affected <6.21.0, patched field >=6.2.10 and prose fix 6.21.0; the
upstream discrepancy remains explicit. The [image advisory](https://github.com/py-pdf/pypdf/security/advisories/GHSA-5hx5-vpf9-5wcv)
names <6.21.1 affected. This continuation also directly checked the official
[6.21.1 release](https://github.com/py-pdf/pypdf/releases/tag/6.21.1), which
corroborates the JPEG-specific fix. This does not clear frozen 6.20.0.

Lazy operations parsing and the absence of direct production form/image APIs
still narrow demonstrated paths without proving immunity. Earlier outline/tree
residuals remain as recorded. No exploit or memory-pressure test ran; no
dependency was selected, security clearance granted or product file changed.
Any necessary dependency/product/package remediation requires a new committed
freeze and new exact-byte native acceptance. Accepted R evidence cannot be
extended to changed bytes.

## Frozen receipts and live metadata

An independent standard-library archive/raw-Git audit reconfirmed all 34 production
members equal R and current HEAD; the original kit's 35 members, 34 manifest pins,
21 helpers, six fixtures and existing extracted files agree. The 41-member carrier
contains the exact original kit and all 35 kit files. Original owner return and
all seven accepted member pins/extracted files agree. Retained hashes are:

| File | SHA256 |
| --- | --- |
| Production ZIP | `70a76bd04f5645e195972bf075ff53d651c9993696a40065241390805d2ee1de` |
| Original kit | `e2f2a36a727edec378b891f5b125a18bea89dc8e41946719ac8b8a6d2e1837fc` |
| Kit bindings | `446ef20b5aad8462f413ab6c43df12baa8c6d46211e921c4933bea3c3e3e818b` |
| Convenience carrier | `adc6d1cb89fc90123cd5ef9c0f71c121f74761b8205fea7a1c4c6947e8a97745` |
| Accepted owner return, 377606 bytes | `1ac523bfb8aebc1511c27629af5134fa9000a08db361a6e207c33280d184a06e` |

The retained checker is byte-identical to raw R (26163 bytes, SHA256
`257cfa9d79ea5fe2779de5572ac634ecc4569d85eee05c65a91a91d9ab7267bb`).
Its new offline invocation against the unchanged retained owner files exited0:
CONSISTENT_NATIVE_REPORT, six receipts, run
`store_f61d4a88eef1415c9ccf811613c6d694`. No stderr was emitted. This is receipt
consistency rechecking, without new native execution or local PDF reopening.
The earlier 79-check native receipt audit remains unchanged, not rerun.

R remains `243d31ee03efac35dc6f7e88c0c1dfd1203081ed`, tree
`ea0db1cc5f5494b508f140deaf25eff87a7e39d9`, VERSION 1.0.3. Fresh API/tag queries
and programmatic comparison reconfirm all six old tag-object/peeled identities,
three release bodies and nine asset metadata records equal T08-gate-recheck.json.
v1.0.3 is absent. Live main remains
`4cc6b00b99685607f3214e731edaaec7d65926df`. Existing final-R CI run39043915939
and all four jobs still report success. No new CI, anonymous asset-byte download
or release-closure acceptance is inferred.

## Checkpoint boundary

Only this new evidence pair and active status/task/continuation/evidence links
change. Historical evidence, owner contracts, frozen product/kit/carrier and
desktop installations remain unchanged. No app/setup/native scenario, runtime
installation/update, policy/PATH/alias change, publication/tag/main merge, force
Git or settings action occurred. Broad implementation suites were not rerun for
this documentation-only checkpoint. Handoff, JSON/links, unchanged statuses,
privacy, intended paths, evidence pins and staged raw bytes are checked before
commit. The companion JSON records actual checks and the harmless corrected
read-only tree-lookup error from the independent audit.

Ordinary commit/push and the final clean HEAD/fresh-origin equality receipt belong
in the external handover, without a self-referential commit hash. STOP after this
checkpoint. T08 remains blocked by STORE-024 and unresolved security disposition;
the owner need not repeat accepted T07.
