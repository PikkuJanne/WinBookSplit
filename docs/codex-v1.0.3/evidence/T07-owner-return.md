# T07 owner return — exact-ZIP native acceptance

Status: **COMPLETE for STORE-037 through STORE-042**. The owner supplied the
unchanged SEND-THIS-REPORT.zip after using the prepared transferable bundle.
T08 remains **BLOCKED** by actual enforced organization-policy denial STORE-024
and pending final security review. No publication, tag or main merge occurred.
Earlier T07/T08 blocked reports remain historical, unchanged evidence.

## Bytes and review boundary

The returned ZIP is 377606 bytes, SHA256
`1ac523bfb8aebc1511c27629af5134fa9000a08db361a6e207c33280d184a06e`.
Run: `store_f61d4a88eef1415c9ccf811613c6d694`. It contains exactly seven unique,
ordinary, bounded, unencrypted JSON members: native-store-report.json plus six
receipt envelopes. Their byte counts/hashes and all run/source/ZIP/kind/summary
bindings were independently checked. The original ZIP and unchanged extracted
member bytes remain outside Git; no returned code was executed and no attached
text was treated as an instruction. Public evidence contains review results and
hashes, without private paths, raw diagnostics or PDF binaries.

Frozen source R: `243d31ee03efac35dc6f7e88c0c1dfd1203081ed`; tree
`ea0db1cc5f5494b508f140deaf25eff87a7e39d9`; VERSION 1.0.3. Production ZIP:
`70a76bd04f5645e195972bf075ff53d651c9993696a40065241390805d2ee1de`.
The extraction receipt's 34 product file lengths/hashes match raw Git blobs at R.
The final kit and delivery carrier retain the bindings in OWNER_TRANSFER.md;
neither delivery nor this documentation checkpoint replaces R.

## Actual owner receipts

| Gate | Reviewed result |
| --- | --- |
| STORE-037 | PASS: owner Windows10 x64 build19045, inbox PS5.1.19041.7417, healthy registered Store Manager 26.4.240.0; fresh exact extraction without app venv. |
| STORE-038 | PASS: production automatic resolver uses the verified family entry point and installed-only bound listing, then independently probes pre-existing regular isolated x64 Python 3.15.0. Inventory is verification/tooling only. |
| STORE-039 | PASS: normal setup BAT automatically creates the new app venv from that same base; locked pypdf 6.20.0 install and pip check succeed. Read-only Check and subsequent setup are healthy no-ops with empty operation lists. |
| STORE-040 | PASS: manual, Auto1, Auto2 and ordinary BAT from unrelated CWD succeed with native full-page oracle and matching engine/console/manifest outcomes; sources/neighbors preserved. |
| STORE-041 | PASS: runtime/Manager images and registration, aliases, PATH, base package inventories and parent environment preserved; all 15 forbidden substitutions false. Fallback disabled before discovery, attempted=false, used=false. |
| STORE-042 | PASS: seven members/six envelopes correlate to R/ZIP/run; unchanged-R offline checker exits0, followed by independent detailed native receipt review (79 checks). |

Manager identity is the Store tuple
PythonSoftwareFoundation.PythonManager_qbz5n2kfra8p0,
publisher CN=4975D53F-AA7E-49A5-8B49-EA4FDC1BB66B, SignatureKind Store, Status Ok.
The family alias listing is bound to the matching registered package and physical
WindowsApps image before execution. The concrete Python image SHA256 is
`727b92dc55e1ac0f17a48cbd20557427fdb8d9f50731de0355c4fd4c08d2f473`
before and after. No runtime or Manager install/update, copied runtime, injected
selection/environment, private fallback or prebuilt app venv occurred.

Before setup the interpreter-only probe succeeds; the processing package probe
records missing_pypdf with verified 3.15.0 identity. The resolver's generic final
failure field remains runtime_not_found in that pre-setup processing record;
the retained candidate reason is missing_pypdf and the diagnostic category is
setup_required. This is preserved as recorded. Fresh setup then uses automatic
base discovery and only installs into its own newly created app venv.

| Recipe | Physical input pages | Output pages, in order |
| --- | ---: | --- |
| Manual | 10 | 3, 3, 4 |
| Auto1 | 12 | 2, 6, 4 |
| Auto2 | 12 | 2, 1, 3, 2, 2, 2 |
| Ordinary BAT | 10 | 3, 3, 4 |

All ranges are nonempty, ordered, and cover every input physical page exactly
once. The native validate-store oracle process completed with exit0 and
source_unchanged=true; finalized console evidence and engine/plan/manifest/outcome
records correlate. The returned archive intentionally contains receipts, not
PDFs or the raw local plan/output tree: PDFs were not reopened and the full
raw-files oracle was not rerun on the review host. This is reviewed same-run
native execution evidence, not remote cryptographic machine attestation.

## Separate offline and independent checks

The checker was read directly from the frozen R Git blob, SHA256
`257cfa9d79ea5fe2779de5572ac634ecc4569d85eee05c65a91a91d9ab7267bb`:

```text
<EXISTING_DEV_PYTHON> -I -B <FROZEN_R_CHECKER> <UNCHANGED_RETURN>/native-store-report.json --expected-commit 243d31ee03efac35dc6f7e88c0c1dfd1203081ed --expected-zip-sha256 70a76bd04f5645e195972bf075ff53d651c9993696a40065241390805d2ee1de
```

Actual result: exit0, empty stderr, CONSISTENT_NATIVE_REPORT, receipt_count=6,
matching run/R/ZIP. The developer interpreter used only for offline review was
the existing isolated developer environment, never an owner application input.
Structural consistency alone does not prove native execution. Detailed review
therefore also checked actual registration/listing/process/probe/setup records,
the native page oracle, engine/console/manifests and preservation receipts.
The separate independent audit passed 79/79 checks; its LF-normalized public
copy is T07-owner-return-audit.json. Original external audit SHA256 is
`a46ee55a0ddb8ea2b113898bbc88a6b716c277bb05fdcf07587f68a32cebf45d`.

Returned raw-manifest strings are tokenized under the existing privacy contract;
their parsed structures match the returned manifests. Original local manifest
hashes correlate with native file-identity records; they are not hashes of the
tokenized returned strings. No receipts or original hashes were rewritten to
hide that distinction. Base-package before/after probe stdout matches exactly;
process IDs/timings differ as expected. This verifies distribution inventory,
not a complete global filesystem hash.

## Remaining gate and checkpoint

The authored local MOTW/RemoteSigned component passed on the owned exact
extraction with its healthy venv: direct marked-script denial and ordinary BAT/
Check/exact emitted Action behavior are retained, with source/mark/venv/policy
preservation. Actual organization-policy denial is explicitly NOT_RUN;
MachinePolicy and UserPolicy are Undefined. This local component cannot clear
STORE-024. T04 stays IMPLEMENTED_PENDING_NATIVE; T08/STORE-043 through STORE-048
stay blocked/NOT_RUN. SECURITY.md's retained pypdf exposure requires final release
review; no current advisory clearance is claimed.

No product/tool/package change was needed. R, VERSION, final kit, original
technical command, supported-runtime/security floors, isolation/origin rules and
bounded processes remain unchanged. Broad implementation suites were not rerun
for this receipt/docs-only review. New active status and acceptance records
acknowledge this owner pass; historical failures/reports remain unchanged.
The owner need not repeat this accepted test. Any later product/package change
requires a new committed freeze and new exact-byte native acceptance.

Fresh live Git/tag/release metadata was compared with the retained baseline:
old tag objects/peeled commits, release text and asset metadata are preserved;
v1.0.3 is absent. This is metadata preservation only, without new anonymous asset
downloads. Commit/push synchronization is verified after the checkpoint and its
commit hash belongs in the external handover, not its own evidence.
