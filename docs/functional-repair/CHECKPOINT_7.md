# Checkpoint 7 — v1.0.4 functional repair and dependency compatibility

30 October 2026. **v1.0.4 is published and anonymously byte-verified.** The owner explicitly
excluded all human testing, sign-off, screenshots, report returns and kit runs.
Unavailable environments remain NOT_RUN, with no owner action required.

## Authority and source

Started clean on `codex/functional-repair`, local HEAD and fresh live branch both
`458969d7c722840bd3480534864e63de27ebb79c`; live main was
`bafa1868da10ab6ba65f112fe29b4bf4deb909f8`. Origin remains
`https://github.com/PikkuJanne/WinBookSplit.git`; one worktree. Read active AGENTS,
STATUS_NEXT and checkpoint 1–6 evidence/maps. Checkpoint 1 has Markdown evidence,
not a CHECKPOINT_1.json. Historical evidence and published artifacts are preserved.

The initial version freeze `f8093b89002530a977424cbe34ca116e95546d3a` coordinated
VERSION 1.0.4, builder/allowlist metadata, affected version assertions and current
docs. Its unpublished 973,011-byte ZIP (SHA256
`157ef70e5a972dc227059834d12c591eb5269d267acff9bd885e3f4caa72cd3d`) and old-pin
results remain bound to that source; they are never relabelled as final acceptance.

Live advisory review then found new issues affecting the preserved dependencies.
The owner rejected introducing another exact Calibre restriction and explicitly
approved the clarified repair: automatically installed pypdf 6.21.4, unchanged
Python support, and optional Calibre 9.17.0 **or a later stable release**. AGENTS
records this narrow exception to the earlier preserve-lock instruction.

Final release source is **`22745a9d4361a77817754d28c9499cbf8fb33532`**, tree
`37dc829f5dc54aa45a0808ceeefbd996ac091a27`, committed and pushed before packaging.
The JSON records all 52 changed source file blobs/hashes after the initial freeze.
Python policy branches, numeric floors and exclusions are exactly unchanged.
Architecture, entry points, all 34 allowlist paths, page/source/output/process/
origin/privacy guards and earlier dependency fixes remain intact. PDF engine,
converter engine, native job and Windows transport source bytes are unchanged.

## Dependency repair and package

The official pypdf advisories [xref trailer EOF](https://github.com/py-pdf/pypdf/security/advisories/GHSA-5849-834g-hvw5),
[page-tree memory](https://github.com/py-pdf/pypdf/security/advisories/GHSA-cggf-277g-3q54)
and [dictionary cloning](https://github.com/py-pdf/pypdf/security/advisories/GHSA-mfr5-77xh-9j7h)
were published on 28 October, with affected `< 6.21.4` and fixed `>= 6.21.4`.
Reader initialization, physical-page access and chapter cloning make these paths
relevant; no universal exploitability or hostile-document clearance is claimed.
The official non-yanked 410,430-byte wheel is SHA256
`33e4fb41ca352810eb4e9838aa0d82d6888a5f8bf682f43e99cce5faa666255c`.

The [Calibre HTML decoder advisory](https://github.com/kovidgoyal/calibre/security/advisories/GHSA-j84x-7j7f-vwx9)
affects local ebook conversion through 9.16.0 and is fixed in 9.17.0. Production
now compares bounded canonical stable numeric versions with a minimum; later
versions are also accepted by error guidance, diagnostics and export validation.
The package manifest records that minimum. Exact 9.17.0 executable hashes remain
QA provenance only. PDF-only use never requires Calibre. The live review covers
72 upstream pypdf and 28 Calibre advisory records, preserving separate legacy
PyPDF2 and upstream metadata-name discrepancies rather than silently ignoring them.

The existing builder and verifier ran from exact committed tool blobs outside the
checkout, preserving the Windows line-ending identity guard. Two builds and both
verifies pass; all three assets are identical between builds in this environment.
All 34 members match raw Git and fresh extraction; all 22 runtime members bind to
the exact ZIP. Matrices use their existing 21-file subset; setup BAT is separate.

| Frozen asset | Bytes | SHA256 |
| --- | ---: | --- |
| WinBookSplit-v1.0.4.zip | 980,017 | `9d402f72092b20ad4d2a1985d8affef0ccfae525f520817e276185db4920a309` |
| release-manifest.json | 16,641 | `9b6918537e04db12588781faa3517b6ccb844b7dc85fb34dec655fa5d7226d5f` |
| SHA256SUMS.txt | 178 | `f3bd58482cb4c15aba1af8999eaa03ae09b6aa5755bd0d08c74a9ee12eb07b08` |

Fresh explicit laboratory environments use the already available serviced Python
bases and hash-required packages. Existing base/old-env inventories, desktop
Python/Manager, PATH, aliases and persistent policy are preserved. Official Calibre
portable bytes match upstream SHA512 and GitHub SHA256; installer and converter
Authenticode are valid. The final isolated test extraction does not replace 9.15.
The converter is 36,104 bytes, SHA256
`c71723b25a566e3f17763a02184d9241dd948c017330ed979fa475ab6dac4bea`.
No private Python fallback or ordinary-test app environment is provisioned.

## Actual verification

Local OS is Windows 11 Pro build 26300. Actual local hosts are inbox PowerShell
5.1.26100.9711 and PowerShell 7.6.6. Only authored synthetic documents are used.

| Check | Actual result and scope |
| --- | --- |
| Focused affected Python contracts | 136 PASS, zero failures/errors/skips, final working source hashes stable. |
| Security regressions | Three meaningful new tests fail against 6.21.3 with finite read/object watchdogs or the missing aggregate budget. All 14 security methods pass on each actual 3.13.16/3.14.8/3.15.0 + 6.21.4: 42 executions, 14 unique methods, zero skips. No large allocation or unbounded hang. |
| Final native runtime matrices | Three targets each PASS 172 units, 10 application cases and 21 native probes, both hosts and real Calibre 9.17.0 EPUB/AZW3 conversion. Exact ZIP subset, source, package inventories and process/tree/stream guards pass. Explicit runtime evidence, not ordinary Store acceptance. |
| Exact ZIP PDF | 11 functional methods PASS; four unchanged PS launches (two per host) each preserve six physical pages in expected order/ranges, one ignored-view warning, plain output without OpenAction, source/neighbors/prior outputs and all payload bytes. Literal Unicode/bracket/ampersand paths, unrelated CWD, no setup or app environment creation. |
| Exact ZIP ebook PS workflows | Four unchanged PS launches PASS: EPUB three physical pages, AZW3 four; three chapters each, full physical content/marker coverage, two retained full PDFs and two discards. Eight native jobs are bounded, assigned, stopped and fully drained. Source/prior/previous outputs, app and converter preserved. |
| Final document policy | 75 strict-validated controls PASS: 30 unchanged direct PS cases per host, 13 controlled BAT cases and two unchanged exports. The BAT is exact; its copied PS has the declared three default changes plus UTF-8 BOM. Other 20 runtime members stay exact. Original 21-file copies first bind to the ZIP. These controlled BAT cases are distinct from unchanged BAT acceptance below. |
| Final setup/remedy/MOTW | Four actual fresh/stale roundtrips × five stages, three marked setup BAT routes × seven stages, seven reopened-PDF oracles PASS. Two native parent controls explicitly mock organization-policy queries; they do not establish genuine organization enforcement. Bound source/base/backup/app/ADS/policy/process guards pass. |
| Unchanged processing BAT | Two actual CMD→BAT→PS5.1 launches PASS, no runtime/PATH/profile/default injection. Actual explicit setup previously created that fixture's authoritative app env. Each preserves ten physical pages, ranges `[0,3],[3,6],[6,10]`, three chapters, source/neighbors/app/env/marks/policies and prior outputs. Healthy prepared fixture, not ordinary Store first-run. |
| Components and final guidance | Immutable pre-freeze snapshot: 67 core + 38 FirstRun Pester cases PASS per host, zero failures/skips. Later root drift is explicitly the corrected converter summary and policy source prose/line endings. Final raw-source analysis: 11 changed PS files, zero findings. Final actual 9.15 refusal, 9.17 resolution and corrected minimum-version summary PASS on both hosts. |
| Independent review/audit | Version, final compatibility/privacy/source safety and completed package/PDF/ebook receipts PASS independent reviews, including reopening outputs and raw stream/source hash verification; no app rerun is inferred. |
| Hosted CI | PASS, all four jobs and four downloaded strict artifact validators. Reference Python: 759 methods; hosted application targets each pass 172 units and eight PDF cases. Actual Windows Server 2025 evidence, with no Calibre/Store/Windows 10 claim. |

The exact unchanged ordinary BAT also ran with inherited natural discovery.
Desired first-run consent/setup/resume is one legitimate NOT_RUN skip, zero
failures/errors: actual exit 3 / runtime_not_found, because natural Python 3.14.7
is below the unchanged 3.14.8 floor. All 34 copied payload members and input remain
unchanged, with no app environment or chapters. Ordinary Store first-run/resume/
repeat/network, Windows 10 and the unavailable private book remain NOT_RUN; no
owner action is required and no explicit/mock/Server result substitutes for them.

## Retained failures and closeout

The original remedy harness incorrectly expected processing to exit without the
new consent prompt; EOF correctly canceled with 130. A noninteractive probe then
correctly omitted the old processing Action line. The maintained harness now uses
actual read-only public `-Check -NonInteractive` for its structured Action, requires
changed=false/zero operations, validates literal AST/text and executes that exact
action before retrying the original request. It does not invent a document
parameter on public Check. A helper-only optional-property StrictMode error was
also corrected. All three failed attempts, the older corrected-helper pass,
initial focused unit failures and intermediate nonfinal passes are preserved.

The old f809 source full run was safely interrupted after 481.475s for the newly
authorized dependency scope; no finalized full pass is claimed. Two failed
portable-tool acquisition attempts remain recorded; final extraction used the
vendor's supported short path. Earlier evidence is never relabelled as new bytes.

[Public final release](https://github.com/PikkuJanne/WinBookSplit/releases/tag/v1.0.4). All three anonymously downloaded assets match the accepted bytes and pass the committed-source package verifier; published notes match the reviewed file. Annotated tag `e462b7127926368561d095079732a03bc68ce64b` peels to the frozen source above. All twelve historical asset downloads (3,208,799 bytes), four old release bodies/stable records and eight old tag references match their baseline. A separately verified durable copy of the new public assets is saved in `%USERPROFILE%\Downloads\WinBookSplit-v1.0.4-22745a9`; existing installations remain intact.

Scripts remain unsigned. Checkpoint-4 kit, checkpoint-5 artifacts and all historical
records stay under their original bindings; no kit is rebuilt or substituted.
Raw diagnostics, private local paths, PDFs, binary tools and test outputs remain
outside Git. CHECKPOINT_7.json contains redacted receipt locations and hashes.
The final documentation-only closure commit leaves the tag/package at the frozen
source. Normal scoped push and clean HEAD versus freshly queried branch/main
verification are reported externally after commit, avoiding a self-reference.
