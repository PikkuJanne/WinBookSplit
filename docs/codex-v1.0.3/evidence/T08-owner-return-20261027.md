# T08 - 27 October owner return and current gate review

Decision: **The returned packet is an internally consistent old-R run, not
acceptance of patched candidate F. No acceptance status changes.** The current
upstream security review also identifies a release-relevant issue in F's
pypdf 6.21.1. Hold another F owner run while that security disposition is resolved.
Original accepted T07 and frozen R remain unchanged.

## Returned bytes and offline review

The supplied SEND-THIS-REPORT.zip is 377554 bytes, SHA256
`436441f5967ef752a69e93f0e22d022651ae405ec360d93542ccbd51fd5093b5`.
Run `store_4193de4d051e48918918a5647c9cd918` contains exactly seven ordinary,
bounded JSON files. All six envelope hashes, run/source/ZIP/kind/summary bindings
agree. An exclusive external snapshot and unchanged extracted bytes are retained;
the supplied ZIP remains unchanged. No attachment code or commands were executed,
and attached text was treated as evidence rather than instructions.

The report and every envelope identify original R
`243d31ee03efac35dc6f7e88c0c1dfd1203081ed`, old ZIP SHA256
`70a76bd04f5645e195972bf075ff53d651c9993696a40065241390805d2ee1de`.
All 34 extraction pins match raw Git R. Setup install/Check/no-op and all four
PDF runtime manifests identify pypdf **6.20.0**, policy **3**. This differs from
expected F `afd75b4f41ab7135e13fbb40818ade9e47daba03`, tree
`08812489efdfb9921f87dca5ceba2c3ab5b37a67`, ZIP SHA256
`216c310e395a098e085f8c3a4db994a2c94072e58bc5982234fad1e5db7b9359`,
pypdf 6.21.1/policy 4. Application version 1.0.3 alone cannot distinguish them.

The maintained checker bytes match raw Git F (SHA256
`257cfa9d79ea5fe2779de5572ac634ecc4569d85eee05c65a91a91d9ab7267bb`).
Actual commands, using an existing isolated developer Python only for offline
JSON review, are recorded in the companion JSON:

| Expected binding | Actual exit | Result |
| --- | ---: | --- |
| F / ZIP216c310e | 2 | BLOCKED: source commit mismatch |
| Declared old R / ZIP70a76bd0 | 0 | CONSISTENT_NATIVE_REPORT, six receipts |

The second command checks the already returned data; it does not rerun the
owner test or promote old bytes to F. Both stderr streams are empty. The separate
independent audit records 171 checks: 166 pass and five expected-F binding/package
checks fail. It finds no internal packet inconsistency.

The receipts describe Windows10 build19045/inbox5.1.19041.7802, verified Store
manager26.4.240.0 and existing regular x64 Python3.15.0. Eleven application routes
and 25 unique bounded probes correlate. Manual/BAT outputs are 3+3+4 pages;
Auto1 is 2+6+4; Auto2 is 2+1+3+2+2+2. The same-run oracle records complete physical
coverage, preserved inputs/neighbors and successful engine/console/manifest
correlation. PDFs and the full raw local plan/output tree are not in the return;
they were not reopened or rerun on the review host. This is reviewed owner-native
receipt evidence for old R, not cryptographic remote-machine attestation.

All forbidden substitutions remain false in the packet, and fallback is disabled.
Preservation records correlate runtime/manager/package inventories and environment.
The local authored MOTW/RemoteSigned component passes, but genuine organization
policy denial is explicitly NOT_RUN and MachinePolicy/UserPolicy are Undefined.
Neither an old-R pass nor this local component clears STORE-024.

## Prepared F carrier and new security disposition

The existing prepared WinBookSplit-Windows10-Test-afd75b4.zip remains intact:
3157973 bytes, SHA256
`c9672265f3a2b5a3653e9c36967f8f64a83bd8ac3799489ecdf412b92d0a18a3`.
Read-only inspection verifies D's manifest, F's bindings, embedded kit and exact
F production ZIP. No carrier/product bytes were rebuilt or edited. The return
does not establish which UI path produced it; its own bindings establish old R.

On 27 October, the official upstream advisory
[GHSA-x3rh-gjgm-fmqc](https://github.com/py-pdf/pypdf/security/advisories/GHSA-x3rh-gjgm-fmqc)
identifies xref-stream memory consumption in versions below 6.21.3 and names
6.21.3 as fixed. [PR4269](https://github.com/py-pdf/pypdf/pull/4269) adds a total
cross-reference entry budget. The reviewed F source calls the pypdf superclass
constructor at winbooksplit_conversion.py:217 before application preflight at
the following line. Static
review of official 6.21.1 source traces initialization into xref-stream parsing
before that preflight. Existing input bounds and supervised deadlines do not
supply this missing budget or a heap cap. This is a release-relevant exposure;
no exploit or memory-consumption experiment was executed.

The separate gate-context audit records seven new advisories, their primary
links and narrow reachability limits. Encrypted page-count, strict-mode and
page-label triggers are not simply declared reachable: the application rejects
encryption at its hook, defaults strict=False and has no observed page-label API
use. The opening/xref-stream issue alone justifies holding F's security disposition.
Earlier Oct18 source/CI passes retain their actual scope and dates; they do not
clear subsequently published advisories.

The proposed minimal follow-up is documented separately, including verified
official non-yanked 6.21.3 wheel metadata (409554 bytes, SHA256
`c8ca4d9a360dbd5d4b9bb7a8389e2a9ae6a106bee722b496f918040ba2daa17e`),
Python>=3.9 and unchanged BSD license bytes. This was data-only review: no import,
installation, application execution, package test, production pin change, new
freeze or security clearance. See ../tasks/08-T08-security-followup.md and
T08-remedial-scope-20261027.json. Do not spend another owner run on F before this
security disposition is resolved.

## Preserved gates and checkpoint

Original accepted T07 remains COMPLETE only for original R and the earlier
377606-byte return SHA256
`1ac523bfb8aebc1511c27629af5134fa9000a08db361a6e207c33280d184a06e`,
run `store_f61d4a88eef1415c9ccf811613c6d694`. Its receipts, 79-check audit,
OWNER_KIT/OWNER_TRANSFER contracts, archived R/F bytes and published history are
unchanged. All 48 acceptance statuses and T01-T07 task records remain unchanged.
New F Store acceptance and STORE-024 remain NOT_RUN; T04 remains
IMPLEMENTED_PENDING_NATIVE; T08 publication remains BLOCKED/STORE-043..048 NOT_RUN.
Scripts remain unsigned.

Fresh read-only metadata preserves six old tag identities, three release bodies
and nine asset records; v1.0.3 is absent. Both F and e5c1e00 checkpoint CI have
four successful jobs. This is historical CI/current metadata review, not new
native or anonymous release-byte verification. No source/runtime/policy/PATH,
alias, installation, publication, tag, main merge or settings change occurred.
Broad application suites were not rerun for this evidence/documentation task.

Public audits distinguish received, raw external and condensed public byte
identities. Upstream prose is condensed where stated; local raw-manifest hashes
are not confused with tokenized returned text. Two auditor assumptions about
run-ID namespaces and tokenized raw-manifest hashing were corrected, without
altering any returned data. Commit/push this review checkpoint normally, verify
clean HEAD against fresh live origin, then STOP. Its hash belongs in the external
handover, not recursively in its own receipt. NEXT_SESSION.md has the exact prompt.
