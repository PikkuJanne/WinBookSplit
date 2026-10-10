# M3-T05 local manifests and redacted diagnostics

AC-063/064/065 scoped acceptance passes. The verified code checkpoint is C3; actual evidence combines retained C1 passes, clean C2 Python/shell/UX and clean C3 diagnostics. The original full C1 attempt and intermediate C2 diagnostic probe remain failed. No standalone successful 18-stage full run is claimed.

Production implementation C1: `47b878e8a4520ea6846c9c8e1446d609da544228`, tree `2955c87225fff7ffe6eec82eb4d91ae90b74cccf`. All 118 mapped raw files match Git C; digest `2f5ac90917945b1e463451ec37b5fe054fdc23e9a51f4b1c12b9c9f63bf286bb`.
Original clean-C1 18-stage full: **FAILED**, native exit 1; 14 of 18 stages passed, four failed. All 118 source paths remained unchanged. Final implementation/test checkpoint C3: **4ee8d87a9c6c07ea2bdf09f22ab85cf80f9838c2**, tree **6e5b85d5096606d007b95556d741de9c3a979275**, with 118 raw Git-matching paths and source digest **0b15b47be7948b2a01ca24b7b1893343130ece3e24cd778583348edc1b61c46f**. All 15 application paths are byte-identical across C1/C2/C3. C3 was pushed normally; fresh clean/live **SYNCED** was observed on the feature branch at **2026-10-10T08:22:53.189257+00:00**. The evidence-only E push, ordinary PR merge and final main receipts are recorded separately after those actions, without self-referential future hashes.
C1 synchronization: normal push completed; the read-only fresh sync at `2026-10-10T07:44:57.318699+00:00` reported SYNCED, with clean local HEAD and live `codex/winbooksplit-v1-m3` both equal C. [Draft PR21](https://github.com/PikkuJanne/WinBookSplit/pull/21) was created and attached. This is branch equality and draft creation, not merge, CI or release proof.

The immutable T05 baseline was `aaebfbb1f063f73bd9ec14eefc652470ef4f264e`. Actual baseline launches under both supported PowerShell hosts completed an authored six-page PDF split with native exit 0, but produced no `WinBookSplit_Run.json` and had no diagnostic export entry point (`baseline-support01.json`). The source PDF stayed unchanged. The cumulative M3 review separately identified the actual pre-M3 main baseline `85916466ef45f25243806fd6d380bf9c3dcee5d4`, tree `e893c2444b2e69c49895c00c6edb0f3bf5444e09`; neither baseline is a reset target.

Local diagnostics now include UTF-8 versions, normalized settings, captured source identity, validated physical ranges, categorized warnings and the authoritative final outcome. The console log reserves 32 MiB for body plus 16 MiB for final records, capped at 48 MiB. Final manifests are published without overwrite from an owned pending file after the log closes. Diagnostic finalization failure cannot preserve a successful application outcome: completed chapter outputs remain available, while the application reports incomplete/exit 6 and does not print Done. Existing primary failures and cancellation retain their meaningful nonzero outcome. Unpublished pending records are rejected by the exporter. Preview and early argument/dependency refusal retain their no-write behavior; records without an available validated terminal or confirmed plan use null plan and metadata-only identity without inventing a content hash. A displayed pending plan alone is not confirmation.

## Completed acceptance observations

**AC-063:** `support-fifth01.json` completed with actual outer and child native exit 0. Its strict independent audit passed 66 controls: 32 application and 34 export controls, split equally across Windows PowerShell 5.1 and PowerShell 7. The application matrix includes success, unreadable PDF, cancellation, no-plan, parser/bookmark warnings, Preview, preflight refusal and diagnostic finalization faults. Twelve application controls use explicitly disclosed copied-code log/manifest fault injections; these are native fault controls, not unmodified production runs. Actual PDF, authored EPUB and genuine Calibre-generated AZW3 controls verify published chapter counts, complete physical coverage and content against captured references. The reference EPUB PDF has three pages; AZW3 has four. No private book was used.

The audit verified 22 finalized run manifests, 28 console logs, 12 redacted summaries and 40 published chapter PDFs across the matrix. All 118 mapped source paths were unchanged, and authored input, neighbor, prior-output and stored settings guards passed. Authenticated owned test cleanup completed. There were no skipped controls.
The independent `support-fullC01-independent-audit01.json` also passes all 66 support controls on clean committed C1, with native child exit 0 and all 118 raw source paths equal the C1 Git blobs. It retains the same 22 finalized manifests, 28 logs, 12 summaries and 40 chapter PDFs. This stage pass belongs to the failed full01 aggregate; it does not turn that aggregate green. The additive confirmation detail checks the actual `[INTERACTION-REPLY]` marker and proves zero replies for the retained pending-plan timeout, correcting the earlier audit's marker spelling without inventing a new run.

The control-native histogram was 24 exit 0, 24 exit 2, two exit 5, ten exit 6 and six exit 130; the aggregate passed because each matched its expected outcome.

**AC-064:** The explicit local exporter projects a fixed schema of trusted application/runtime tokens, typed settings, numeric page ranges/counts, fixed outcome codes and known warning categories. It excludes document paths, titles, filenames, messages, arguments, environment, identities, hashes, timestamps and run IDs. Local full runtime versions remain in the original record; the support summary uses fixed PowerShell buckets `5.1`/`7`. The 12 native summaries cover actual success, read failure, cancellation, incomplete and no-plan records plus authored secret-bearing copies on each host. Original manifest bytes remained unchanged. Twenty-two invalid export controls refused with native exit 2, including existing destinations, source alias/hardlink, relative paths, duplicate keys, bad UTF-8, wrong protocol, junction output ancestor, pending manifest and malicious selected fields. No file was overwritten; no upload or network operation exists in the export path.

Focused helper tests passed 17/17 on each actual host, native exit 0, with zero failures/skips/not-run. Additional controls cover bounded input, unpaired Unicode, contradictory page/count fields, unsafe ancestors and competing writes. These are helper tests, distinct from application acceptance.

**AC-065:** Each native parser-warning control visibly/logged two categorized `pypdf_parser_warning` records with suppressed count 0. Each invalid-bookmark control recorded `invalid_destination` and retained the correct split result. Parser text remains local; only fixed category/code/severity and bounded counts enter the redacted summary. This does not claim every malformed PDF is recoverable or establish a rendered-page fidelity result.

## Review and preserved attempts

The read-only cumulative M3 production review passed with no remaining blocker. It bound 15 current production hashes to the targeted receipt, compared the pre-M3 Git blobs and confirmed 11 planner/ownership declarations unchanged by AST comparison. It reviewed process streams, dependency resolution, page coverage, output/cleanup ownership, injection risks and public-data projection. This reviewer authored the exporter; the parent separately reviewed that owned helper. Source review is not a certification of every possible process/filesystem interleaving. The fresh `cumulative-M3-production-C-review01.json` binds this review to clean C: 13 production paths are raw-identical to Target05, and the exporter entry point/support helper differ only by exactly reconstructable CRLF-to-LF conversion. All 118 Git blobs match the worktree and C source receipt. Five test-only oracle/fixture/cleanup changes and one test-file line-ending change are tracked separately; no earlier native receipt is relabeled as clean-C acceptance.

A later authored same-byte replacement control reproduced a cleanup identity gap in the inherited test adapter. Authentication now carries the directory/member identities into held deletion and preserves replacement objects. Eight pure filesystem/manifest unit methods passed; focused warning methods 10 and receipt methods 63 also passed with native Python exit 0. Separately, three independently resealed synthetic copies altered an output range, execution source and summary status/code. The earlier strict oracle accepted those contradictions; the corrected oracle rejects all three. Its 14 pure support receipt methods passed, and all 66 original Target05 rows pass read-only revalidation with the stronger oracle. These are test/proof fixes and retained-row audits, not 66 new native runs. The final committed affected checks and reviewed retained C1 passes bind the corrected adapter/oracle bytes with distinct source maps.

An actual ordered helper reproduction found marker creation before directory leases could redirect a new write through an authored junction. The fix holds the base before reservation and the new console directory before its marker. Fresh fixed helper controls passed both hosts; later Directory.Move succeeds after lease disposal. Logging Pester passed 6/6 per host, native exit 0. These are source-bound helper controls, not a full application concurrency or GUI pass.

All earlier receipts remain preserved. The first focused support run failed a privacy assertion that mistakenly matched a permitted count-field name. Logging attempts 01/02 failed directory-handle/rename controls. Native support attempts 01/02/03 stopped on incomplete process-proof expectations, an unrecognized actual failure status and inherited `remove_diagnostic` raising `KeyError: cleanup_complete` for a parser-only engine diagnostic. Attempt 04 reproduced a real corrected-log-footer mismatch under copied-code finalization failure; the application and oracle were corrected before attempt 05 passed. The initial warning unit assertion also confused object identity with the existing frozen mapping reconstruction; its stack is retained only in tool output. The initial cumulative-review sealing helper refused changed Main bytes; its later receipt explicitly binds the newer bytes rather than relabeling earlier tests. Failed workspaces are retained where recorded, and absent raw output is not represented as a saved receipt.

## Preserved clean-C1 full failure and narrow test fixes

`full01.json` is FAILED/native 1 with 18 stages: 14 passed, while `shell-0`, `shell-1`, `diagnostic-regression` and `ux-regression` failed. Its 302 Python methods and native CLI 72, launcher 70 and support 66 controls passed within their recorded stages. The authored outer and failed UX workspace were retained. No original failed receipt or retained file was replaced or deleted.

Both original Pester hosts passed 110/111 with one failure and zero skipped/not-run tests. Fresh detailed reproduction on exact C1 PowerShell sources identified `Outcome.Tests.ps1`'s child-stream renderer test: its mocked application setup omitted `$KeepConvertedPdf` under StrictMode. The only fix adds `$script:KeepConvertedPdf=$false` to its existing BeforeEach, matching the real application's switch default; no assertion or production code changed. Full container retries passed 111/111 under both actual hosts, native 0, zero failures/skips/not-run. Each run also had 18 syntax files with no errors and 17 scaffold files with no findings. Main's 71 static observations per host (68 warnings, three information) remain observation-only; application static analysis is not claimed clean. The initial external PS5.1 reproduction wrapper omitted the per-host builtin module path and failed before report creation; its captured streams and the wrapper failure remain preserved before the corrected dual-host reproduction.

The diagnostic failure was the inherited base-membership helper assuming a parser-only diagnostic owned `record_path`, raising `KeyError: record_path`. The separate adapter reproduction reuses the real same-C parser payload in an explicitly reduced fixture; it is not a new original native frame. The UX failure expected a displayed pending plan to be a confirmed plan despite zero execute replies. The actual native pending timeout was exit 130, with null plan and metadata-only identity, and the retained files stayed byte-identical. These are test/validator adapter corrections, not production behavioral fixes. C2 fixes these adapters and passes its affected Python/shell/UX stages. Its diagnostics rerun then exposed one further inherited probe omission: the trusted AST subset loaded Run-PythonSplitter without the shipped ConvertTo-WinBookSplitDisplayText dependency. Both original host subsets failed; exact external copies adding that existing helper passed the bounded stream/argv subset. C3 adds only the helper name to that probe list and passes the complete diagnostics stage. No production file changed after C1, and neither failed receipt is relabeled. No standalone successful 18-stage full on C1, C2 or C3 is claimed.
## Final targeted and affected verification

| Snapshot | Actual result | Scope |
|---|---|---|
| C1 full attempt | native exit 1;14/18 stages pass | Original failure preserved; 66 support controls, 72 CLI controls, 70 launcher controls and other passing stages retained |
| C2 Python | native exit 0;305 methods | Complete stdlib unit suite, including warning/identity/consent/receipt guards |
| C2 shell | native exit 0;111/111 per host, zero skipped/not-run | Actual PS5.1/PS7 Pester,18 syntax files and 17 selected scaffold analyses each; legacy Main analysis observation-only |
| C2 UX | native exit 0;28 controls | Actual redirected interaction/plan/consent/timeout and real working-PDF controls, zero skipped; no human/viewer result |
| C2 diagnostics01 | native exit 1 | Failed inherited AST subset dependency; receipt preserved |
| C3 diagnostics02 | native exit 0;23 engine cases and 2 real host probes | Complete categories/decisions/protocol guards, literal quoting and production dual-stream function |

Each passing affected outer report proves unchanged mapped source, unchanged synthetic input and removed owned working directory. C1's failed outer/UX workspace remains retained. The independent retained-pass audit accepts all 13 passing C1 child receipts through the final validators, with original raw child seals distinguished from in-memory reserialization. Two actual parser-only payloads pass the final guards and zero-cleanup trap. This pure audit is not a native application rerun. Exact raw C2-to-C3 comparison finds only the one probe-list line; Python/Pester/UX and all production/validator bytes stay unchanged. The source/delta maps and independent review justify reusing those actual passes.

The machine record retains all three source identities, each affected stage's source commit, original failures and hashed raw receipt ledger. C3's nine changed paths relative to C1 are test-only. M3 source review has no remaining production blocker. This meets the task's targeted-plus-affected checkpoint; later CI and package gates remain separate.

## Runtime and command disclosure

Actual environment: Windows 11 Pro x64, build 26300.9457, 26H2; Windows PowerShell 5.1.26100.9444; PowerShell 7.6.5; Python 3.14.8; pypdf 6.19.0; Calibre 9.15.0. These are this workstation's observations, not a clean-install portability pass. No global policy changes were made.

The following are redacted templates of recorded argv, not literal commands executed with placeholder text. `<Python>`, `<PS51>`, `<PS7>` and `<Calibre>` denote the verified runtimes; `<Repo>` denotes the source checkout, and `<OwnedRun>`/`<OwnedCase>` denote authored isolated test directories. `<ToolRoot>` denotes the isolated pinned Pester/PSScriptAnalyzer directory and `<Task>` the external evidence directory. The failed full01 outer runner started from `<Repo>` and created its own unrelated child working directories. The native support child ran from `<OwnedRun>`; each export ran from `<OwnedCase>\c` with closed stdin. Exact original commands/cwds are retained locally, without publishing profile paths.

```text
<Python> -I -B <Repo>\tests\run_tests.py --layer python --report <Task>\affected-python01.json
<Python> -I -B <Repo>\tests\run_tests.py --layer shell --report <Task>\affected-shell01.json --tool-root <ToolRoot> --shell-path <PS51> --shell-path <PS7>
<Python> -I -B <Repo>\tests\run_tests.py --layer ux --report <Task>\affected-ux01.json --shell-path <PS51> --shell-path <PS7> --calibre-path <Calibre>
<Python> -I -B <Repo>\tests\run_tests.py --layer diagnostics --report <Task>\affected-diagnostics02.json --shell-path <PS51> --shell-path <PS7>
<Python> -I -B <Repo>\tests\run_tests.py --layer full --report <Task>\full01.json --tool-root <ToolRoot> --shell-path <PS51> --shell-path <PS7> --calibre-path <Calibre>
<Python> -I -B <Repo>\tests\support\characterize_support.py --report <OwnedRun>\support-regression.json --calibre-path <Calibre> --shell-path <PS51> --shell-path <PS7>
<PS51-or-PS7> -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File <OwnedCase>\a\Export-WinBookSplitDiagnostics.ps1 -ManifestPath <OwnedCase>\WinBookSplit_Run.json -OutputPath <OwnedCase>\support-summary.json
```

The export template substitutes a safe output leaf for the actual authored Unicode/metacharacter fixture name. `RemoteSigned` is a process invocation option, not a global settings change.

## Receipt seals and limits

Raw local receipts contain authored diagnostic details and workstation paths. The public narrative publishes only safe receipt names, counts and hashes.

| Receipt | SHA-256 |
|---|---|
| `baseline-support01.json` | `b8ed4d8770aea49df49669b3702bf2d257ef16e2c51e7a89739d22d3daf4a917` |
| `support-fifth01.json` | `f6b874c49e4c3d5beb9591a5da37d96f9b34f3fa9166062f42eb7896df8d7310` |
| `support-fifth-independent-audit01.json` | `ba49f465218da7f293c48f10c993072569d5d7052d5232e541acad70eef74f57` |
| `implementation-handoff03.json` | `5a4d8c71240906438525f0e7f7f770fd616a4d918dc554e0cd914937a7f3d349` |
| `production-source-review04.json` | `b9b90d5ab756916f3dcfacbf41b0817ae8b4bcacf267264578d63097c847427f` |
| `cumulative-M3-production-review02.json` | `ee8820374c9a1e05940d25f4152d4eb28a1e39b2602b6000407136c592b622e7` |
| `cumulative-M3-production-C-review01.json` | `d7ce2c44750e642815e975f6967ec9e3c611ae2be2be328078f8376b551effe4` |
| `sync-C-command01.json` | `bc4861ce126c10e1292e29f2bb9ecf8591f697083a0ad9da9ef8188411d7d47f` |
| `support-final-oracle-review01.json` | `dbfc7d989c05b927ba4dc5ee476585fef4283bf16ded8fb433dc88fdab41359a` |
| `authenticated-identity-gap05.json` | `93e32681fd963a85bd213edaa41c4015e43f6d8f6fdba83c470c1b67074c9b3c` |
| `identity-candidate05-tests.json` | `174a6dc50d2355e6052d03ead8927343574e6db4e1f898b5ed55843cb220d00c` |
| `full01.json` | `17ec486562f25863407830087078418185aca5fb18a2378cf3c5c2f02ce91caa` |
| `support-fullC01-independent-audit01.json` | `131987f928c023ac48c99def74e489c3ac6f343325b23dc0cf25378c558eaa12` |
| `support-fullC01-ux-confirmation-detail01.json` | `c37300e57d58e0aa182bf14defdf6fb5944712b8c33eed7edb21a7082bca8f13` |
| `full-pester-C1-failure-review01.json` | `5b50a94e9511cf5eac2ca8978a50d84dbc5419cbeb374c63ffb40a5486a96695` |
| `full-pester-test-fix-verification03.json` | `4cace783ced8876d6188b928ad78845e12263f1e04ba9b896217f7773949025b` |
| `full01-adapter-reproduction01.json` | `32d8bedc79c79d5cef5b77d40b2883621f408cdd284198e6af2480cf928a55fe` |

Targeted raw-source digest: `a80bd06f249b46462069a36a6a80ca09bfa6d9b5d95c1cc1e1af650dfb9c1208` for 118 paths. This identifies the earlier targeted working snapshot. The separate C1 digest above binds committed C1 bytes; its actual 18-stage full failure remains failed. Final C3 raw binding is the distinct digest **0b15b47be7948b2a01ca24b7b1893343130ece3e24cd778583348edc1b61c46f** above; C2 passed stages retain their own map and do not become native C3 reruns.

Human M3-T01 controls remain permanently closed and are not relabeled as current source or package proof. T05 used no human, Explorer, viewer or rendered-PDF checks. No CI, clean-OS installation, release package/tag/assets, anonymous asset download or v1.0.0 publication pass is claimed. The code checkpoint has actual scoped acceptance and a fresh normal-push/live-sync receipt. The canonical evidence checkpoint records these completed facts; its own push/merge/main synchronization belongs to subsequent external/thread/PR receipts. Future evidence-only commit hashes and their push/sync belong to subsequent receipts; they are not self-referential claims in this narrative.

## Final receipt seals

| Receipt | SHA-256 |
|---|---|
| `affected-python01.json` | `55fa91460e3d439b3b36e0a248f43c2a1c66c9725f746c5b986e2e9e866341c4` |
| `affected-shell01.json` | `30fbeb823b4ed548ba0db5e23cef99abd4ea158c301f0e8767127790bfccf626` |
| `affected-ux01.json` | `5f962c0b684a0378951caaa4b0c54d7709b1dc161abc2c360bdd042c77bdf8e9` |
| `affected-diagnostics01.json` | `d38503d035ebca565fe286ee806cceab1376e3fdeeea717bd0bd1a37fc700d1a` |
| `affected-diagnostics02.json` | `2d43ec36ba6860949e9a0de4cf80c2a48b65eaf191e500662b66ec7852799c2b` |
| `source-C1-to-C3-delta01.json` | `cba4f200fafcf8e1763ec85d286e7d12ee73eedb11f2f5a6d92460e3ef8e6b88` |
| `source-C3-after-diagnostics01.json` | `76fb35b43fa65159508def14239f2a281e170a4f82d6fa0c216a1a75eff534fc` |
| `sync-C3-command01.json` | `3807220417a4d4978d9be8f31aaa6b289ee3055db0de09e847de9f6a1e91aab1` |
| `final-C03-evidence-independent-audit01.json` | `a323102b99144217622e6c76cf2c62776fbab319e71eabb4ab3171d16ab71f57` |
| `final-C03-evidence-live-detail01.json` | `24f9f0b1c382b57d89418cbd64e4d6e32e4246425e052d3c20d3a7b0c5f3689b` |
| `final-C03-coverage-contract-correction01.json` | `af72752291a612f5a61a59a08d81cb510cd65b2c5c15acf32d3ca26b5a754d8c` |
