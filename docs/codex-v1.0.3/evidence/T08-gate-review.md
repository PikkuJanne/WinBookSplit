# T08 gate review after accepted owner return

Status: **BLOCKED**. No publication, tag or main merge was attempted. T07 remains
COMPLETE for STORE-037 through STORE-042; STORE-024 remains NOT_RUN, T04 remains
IMPLEMENTED_PENDING_NATIVE and STORE-043 through STORE-048 remain NOT_RUN.
No acceptance status changed in this review. The older T07/T08 blocked records
retain their historical scope. Current technical security review is complete,
but the retained dependency's release disposition remains unresolved/BLOCKED.

## Accepted owner evidence retained

The unchanged 377606-byte SEND-THIS-REPORT.zip still hashes to
`1ac523bfb8aebc1511c27629af5134fa9000a08db361a6e207c33280d184a06e`.
Seven ordinary bounded members, six envelope bindings and all 34 production
member hashes agree with frozen R. The separate checker read from the raw R blob
again exited0 with CONSISTENT_NATIVE_REPORT and empty stderr. An independent
receipt review reconfirmed actual Windows10 build19045/inbox PS5.1, pre-existing
Store Python3.15.0, fresh automatic setup, matching venv base, healthy Check/no-op,
all four complete-page recipes, preserved installation and no forbidden fallback
or hidden selection. The owner need not repeat this accepted test.

This is a recheck of supplied same-run native receipts, not new execution or
cryptographic machine attestation. The return contains no PDF binaries for local
reopening. The earlier independent 79-check audit remains unchanged. Both actual
policy observations show MachinePolicy/UserPolicy Undefined; their authored local
MOTW/RemoteSigned component does not prove enforced organization-policy denial.
No authorized enforced-policy target or new STORE-024 evidence was supplied.

## Current dependency security review

Official upstream advisories were reviewed on 17 October 2026 alongside unchanged
production source. The retained pypdf 6.20.0 includes the
[multiple-filter fix](https://github.com/py-pdf/pypdf/security/advisories/GHSA-7mh2-4gg2-979f),
but is in the affected range for
[large arrays/dictionaries](https://github.com/py-pdf/pypdf/security/advisories/GHSA-f54x-cmq8-qg57)
and [many small content-stream operations](https://github.com/py-pdf/pypdf/security/advisories/GHSA-cgf9-327m-v77c);
both name 6.21.0 as fixed. The upstream [6.21.0 release](https://github.com/py-pdf/pypdf/releases/tag/6.21.0)
and [merged parser-limit change](https://github.com/py-pdf/pypdf/pull/4218)
corroborate those fixes.

PolicyPdfReader initializes pypdf before application structural preflight, and
indirect object resolution can parse arrays/dictionaries before the application
checks their sizes. This leaves release-relevant parser exposure unresolved.
The process job supplies kill-on-close, deadlines and stream bounds, without a
process/job memory cap. These checks do not establish a heap sandbox.

Converted-PDF checks call get_contents/get_data. Upstream
[6.20.0 page code](https://raw.githubusercontent.com/py-pdf/pypdf/6.20.0/pypdf/_page.py)
and [ContentStream implementation](https://raw.githubusercontent.com/py-pdf/pypdf/6.20.0/pypdf/generic/_data_structures.py)
show lazy operations parsing; get_data alone does not demonstrate the particular
many-small-operations trigger. A search of production Python PDF sources found
no ContentStream.operations access or use of extract_text, get_fields,
get_form_text_fields, .images, PIL or Image.open.
These narrow static observations do not establish comprehensive exemption.

The [form-field advisory](https://github.com/py-pdf/pypdf/security/advisories/GHSA-6vvc-x485-93v7)
currently has inconsistent upstream metadata: affected `<6.21.0`, patched field
`>=6.2.10`, and description fix `6.21.0`. That discrepancy is preserved explicitly.
The [image-loading advisory](https://github.com/py-pdf/pypdf/security/advisories/GHSA-5hx5-vpf9-5wcv)
names `<6.21.1` affected and `>=6.21.1` fixed. No production form-field retrieval or
image-loading API was found; this does not prove hostile documents harmless.

The companion JSON records exact production line references and advisory ranges.
No exploit PDF or memory-pressure experiment was run. This review grants no
security clearance and selects no new dependency. SECURITY.md, the lock, Python
floors, product and package remain unchanged. Any later required remediation
needs a new committed freeze and new exact-byte native acceptance; the accepted
R result must never be relabelled as acceptance of changed bytes.

## Frozen bytes and remote preservation

Started clean at `eb3034619ad7d8278e4c7cb18fad89fb5466fa8d`, tree
`109fe4227a09a79feb1544955ef96e0a3ba97151`, on the requested branch, equal to
fresh live origin. Only root AGENTS.md was found; existing origin and single
worktree remain. Frozen R stays `243d31ee03efac35dc6f7e88c0c1dfd1203081ed`,
tree `ea0db1cc5f5494b508f140deaf25eff87a7e39d9`, VERSION1.0.3.

Read-only byte checks reconfirmed all 35 original kit members/34 manifest pins,
21 helpers equal to R, six authored fixture hashes, four production-builder
provenance files, all 34 production members equal to R and current HEAD, and the
41-member carrier's unchanged nested kit/35 files. Existing extracted kit bytes
also agree. No archive was rebuilt or native controller launched.

| Retained file | SHA256 |
| --- | --- |
| Production ZIP | `70a76bd04f5645e195972bf075ff53d651c9993696a40065241390805d2ee1de` |
| Original kit | `e2f2a36a727edec378b891f5b125a18bea89dc8e41946719ac8b8a6d2e1837fc` |
| Kit bindings | `446ef20b5aad8462f413ab6c43df12baa8c6d46211e921c4933bea3c3e3e818b` |
| Convenience carrier | `adc6d1cb89fc90123cd5ef9c0f71c121f74761b8205fea7a1c4c6947e8a97745` |

Fresh old tag objects/peeled commits, release text hashes and asset metadata
match the previous accepted-owner review. v1.0.3 remains absent. Main remains
`4cc6b00b99685607f3214e731edaaec7d65926df`. This is metadata preservation,
without new anonymous asset-download or release-closure acceptance. Existing
final-R CI run39043915939 still reports all four jobs successful; no new CI or
native result is inferred.

## Checkpoint boundary and next action

Only active task/status/next-session/acceptance links and these new evidence files
change. Historical reports, accepted return, OWNER_KIT.md's technical contract,
OWNER_TRANSFER.md, original kit/carrier and owner installations remain intact.
No product or dependency change justified broad implementation suite reruns.
Handoff, JSON/evidence links, exact evidence-byte hashes, privacy, whitespace,
staged blob equality and intended-path preservation are checked before commit.

T08 stays BLOCKED pending genuine STORE-024 enforced-policy evidence and a
resolved security release disposition. Preserve the existing prepared kit; do not
ask for a repeat of accepted T07 or manufacture enforcement by changing desktop
policy. Any later necessary human test needs one validated transferable launcher
and report handoff. Publish/tag/main merge remain forbidden while gates are
missing. Ordinary docs checkpoint commit/push and fresh clean branch equality
are recorded externally after this record; STOP at that checkpoint.
