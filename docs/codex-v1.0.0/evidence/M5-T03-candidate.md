# M5-T03 — Extracted candidate Windows checkpoint

**Partial; task BLOCKED on AC-088.** AC-087/089 pass their recorded native/output/agent-render scope. Separate user reports and output audits establish **G01-G03 PASS**. Under the user-authorized **D13** consolidation, only **four actual human flows R01-R04 remain NOT RUN**. Original unrun G04-G16 are SUPERSEDED, never PASS; native-only variants remain explicit. See [human record](M5-T03-human-smoke.md), [machine evidence](M5-T03-candidate.json) and [smoke matrix](M5-T03-manual-smoke.md). Date:2026-10-10 Europe/Berlin; raw receipts use UTC. Do not advance to M5-T04.

## Source and candidate

Started clean at milestone branch/live head `925f0eceae9d5b5266cef2487057c517f06c1bbb`. PR28 was already user-merged; ordinary fetch and fast-forward reconciled the branch to live main **a282cf4412494063b91f2a71cb21b283731569f1** (C), tree **01c081dcb74df7017ebee08345242c753dc8cad5**. No application, package recipe or test-source changes were needed. The tested clean C has **175** exact mapped Git/raw files, digest **97aa8a37e04db67a56e86935d04666d8a3ecd62d2964a785ab07ba17607ea4e0**. All 28 payload files and all 16 runtime paths are byte-identical to T02. The 23 historical unmapped newline differences remain preserved; this is not an all-tracked-file raw/Git equality claim.

The actual explicit-C package is `EXTERNAL/candidate-main01`. The ZIP matches T02 byte-for-byte; manifest/checksums correctly bind this merged C and therefore differ from T02 provenance files. An independent stdlib/Git artifact audit passed **518/518** member/hash/raw ZIP/provenance/runtime/link/recipe checks.

| Asset | Bytes | SHA-256 |
|---|---:|---|
| WinBookSplit-v1.0.0.zip | 413958 | 1f29981128c5fa74d17faca9e97d67f2605f008e80068d1f76809cbb79218645 |
| release-manifest.json | 11546 | 125bc2b38bdc6ebcdbb08136a2ea9d7c79375b7849ed1d6fd47e5864496f437d |
| SHA256SUMS.txt | 178 | 44d5b0f822a19c27929772d16647914fd8701a030572ab4d0dc1f4fb7a7ad39b |

Payload digest: **e511be767306327ef0ee81f36c164d88fc43869e85cefb11282c43c4b3990cf2**. These assets remain local, unsigned and unpublished. They are not M6 final release source/assets.

## Actual environment and isolation

Existing Windows 11 Pro x64 workstation, CIM version `10.0.26300`, build **26300.9457**, display version 26H2; Windows PowerShell **5.1.26100.9444** and PowerShell **7.6.5**. The historical registry ProductName says Windows 10 Pro; it is recorded without converting this into a Windows 10 test. Sandbox executable exists, but feature readiness is UNKNOWN because the read query requires elevation. No elevation, enabling or clean OS/Sandbox/VM acceptance is claimed. This is D11's isolated current-workstation method.

Final run02 made two new extractions (`app51`, `app7`), two unrelated working directories, and a fresh runtime-only `.venv` per host using the actual extracted `docs/SETUP.md` block with literal path substitutions. Regular GIL AMD64 CPython **3.14.8**, pip **26.2.1**, plain pypdf **6.19.0**; pip and pypdf are the only installed distributions. Imports resolve to each candidate `.venv`, outside the checkout. Candidate execution uses only shipped runtime code. External development helpers are frozen from C, hash-bound and excluded from the payload.

Each application child uses controlled shell/Windows inbox PATH and inbox PSModulePath, an explicit real Calibre path (**9.15.0**), and the application-venv Python selection. There is no developer/setup Python, checkout or Calibre fallback on that controlled PATH. Process-local RemoteSigned leaves persistent execution policies unchanged. Native wrappers preserve literal argv, both streams and finite deadlines, recording parent/stream completion; application receipts separately record owned child-tree completion. No policy/settings change, source-file replacement or installation during processing is inferred.

A read-intended inventory command `py list` unexpectedly auto-updated the Python install manager to **26.4**. The preceding manager version and that individual command's native exit are UNKNOWN; its combined shell command returned 0. This side effect is retained and was disclosed to the user. No runtime install is claimed from it. All later setup, processing and audits use explicit executable paths. No alias was invoked during the independent final audit.

After the user requested clear step-by-step Explorer instructions, scoped dependency preparation exclusively copied the existing trusted Calibre **application subtree** into the previously absent supported per-user known location. All **1,345 files/83 directories/662,902,308 bytes** were verified byte-for-byte; source remained unchanged, with no reparse/hardlink/overwrite. Library/settings siblings were excluded. No installer, PATH, registry or execution-policy change. Separate **four additional default-discovery previews** (EPUB/AZW3 on each host, no explicit Python/Calibre flag) passed native0 with `application_venv`/`known_location` identities, exact three/four physical reference-page content, no publication/log/stage, and complete source/copied converter maps unchanged. This preparation does not relabel the canonical38-case report or prove actual Explorer inheritance/interactions. The current guide checks the actual human child resolver lines. The earlier separate-Explorer PATH experiment checklist is superseded historical preparation.

## Commands and observed results

Roles: ROOT is the checkout; EXTERNAL is the retained local M5-T03 workspace outside ROOT; DEV is the explicit retained CPython 3.14.8 development venv. Exact local argv, absolute dependency paths, raw stream bytes and retained output trees remain external. Machine evidence publishes safe relative names, byte lengths and SHA-256 seals.

| Actual command / check | Outcome and scope |
|---|---|
| DEV -I -B record_command.py --receipt build-candidate-main01.json --timeout 120 -- DEV -I -B tools/release/build_package.py build --repo ROOT --commit C --output EXTERNAL/candidate-main01 | Native 0, exact committed-C assets; source unchanged |
| DEV -I -B prove_source.py EXTERNAL/source-proof-main01.json | Native 0; clean C, 175 mapped files exact, unrelated historical raw differences retained |
| DEV -I -B independent-candidate-assets-audit.py | Native 0, 518/518 checks; no project imports |
| DEV -I -B run_extracted_candidate01.py | **FAILED/native 1**, 17 completed passing cases; retained external-oracle failure described below |
| DEV -I -B reproduce_driver01_failure.py | Read-only reproduction of the retained malformed-EPUB diagnostic-directory oracle failure |
| DEV -I -B run_extracted_candidate02.py | **PASS/native 0**, 38/38 processing cases on fresh separate extractions/venvs; all raw streams and work retained |
| DEV -I -B independent-native-audit-main01.py | **FAILED**, raw CRLF versus normalized report text comparison; retained separately |
| DEV -I -B independent-native-audit-main02.py | **PASS/native 0**, 65,818/65,818 independent retained-artefact checks; no application rerun |
| DEV -I -B record_renderer_version.py | Explicit Poppler-only `pdftoppm -v`, actual native 0; version 26.07.0/binary SHA bound to final renders |
| DEV -I -B prepare_explorer_calibre.py | Exclusive supported known-location application copy; full byte/ownership/source stability checks pass; no private library/settings |
| DEV -I -B run_default_discovery_preflight01.py | Separate four default-discovery preview cases PASS/native0; no explicit dependency flags; zero publication and complete source/copy stability |
| view_image on final PS51 PDF/EPUB/AZW3 first/last output PNGs | Six actual agent views, readable expected content; sealed visual receipt, no human viewer claim |

**AC-087 PASS:** 19 actual processing cases per supported host, 38 total; 72 native operations including setup/probes/fixture generation/references/rendering/support. Actual processing exit counts: **0:26, 2:2, 4:4, 5:2, 130:4**. Nonzero statuses are expected tested errors/cancellations, not harness passes masking application failure. Dependency-free Version works on both extractions before venv setup.

Cases include manual `1`, `1,4,7`, `4,7` with explicit page-1 addition, sorted/deduplicated starts, automatic levels 1/2, preview levels 1/2, rejected `1,,4`, missing outline, redirected confirmation/input cancellation, redirected manual and selected-level fallback, repeat-run preservation, genuine EPUB/AZW3 conversion and malformed real converter failures. Redirected controls prove CLI behavior only; they do not prove Explorer interactions or BAT's actual GUI exit behavior.

There are **22** successful publications, **74** chapter PDFs, **230** reopened physical output pages, four zero-publication previews and two validated redacted support exports. Thirty actual finalized logs are independently authenticated. Input, neighboring `Book.pdf`, prior publication, extracted payload, assets, helpers and tool identities remain stable. Four malformed conversions preserve real Calibre native 1 and application exit 4, zero chapters, no final publication/stage, and exactly authenticated owned failure diagnostics.

**AC-089 PASS:** original MIT authored inputs only. The PDF has 12 physical page markers, two front-matter pages, parent boundaries and a childless last parent. Level1 physical ranges are `1-2 / 3-8 / 9-11 / 12`; Level2 ranges `1-2 / 3 / 4-6 / 7-8 / 9-10 / 11 / 12`. Every successful chapter is reopened and compared to its physical reference pages. EPUB produces three reference pages. AZW3 is a genuine Calibre-generated KF8 input, not a renamed EPUB; its conversion has four physical pages, including the generated final TOC. Manual `1,2,3` preserves ranges `1 / 2 / 3-4` and that fourth page.

Eighteen retained PNGs comprise six independent format-reference slices and twelve first/last output slices across both hosts. All **12** same-renderer reference/output comparisons match pixels exactly. Six PDF reference/output render operations returned native0 with the same retained **1,289-byte font-diagnostic stderr**, SHA256 `74095b267fcd746d2f088f0aea89f337e472614c4726d4626eb919cbb3107ff8`; no warning-free/all-font claim is made. The agent visually inspected final PDF first/last markers, EPUB first/last chapter text, and AZW3 first chapter/final TOC. No missing or clipped expected content was observed. The AZW3 TOC emits the documented `cross_chapter_link_dropped` warning for links outside its final chapter; physical TOC content is retained, not a promise that cross-chapter links survive. These small original fixtures do not establish general conversion fidelity or reopen historical M3 human testing.

## Retained failed trials and defect disposition

Run01 stopped at the PS51 malformed EPUB case because the external oracle forbade every new directory except the console receipt. The supported application intentionally retained an owned `.WinBookSplit-failed-<run-id>` containing only owner and failure records. Actual app exit 4, written count 0, no final/execution publication and cleanup were correct. Read-only reproduction and independent review establish this external oracle defect; no production change was made. Driver02 narrowly authenticates exact ownership/schema/run/code/message/source/conversion/record-path members and absence of the stage, while still rejecting partial-looking output. It also requires exactly one finalized log for successful noninteractive publication; the original successful cases did already emit their actual logs. The original failed driver/report/work remains retained and FAILED.

Auditor01 incorrectly compared raw CRLF log text to the driver's `read_text` universal-newline LF text. Raw hashes/byte lengths matched. Auditor02 explicitly normalizes only that text comparison and preserves raw log/stream hash and size assertions. The original failed auditor remains retained. The canonical successful native report SHA256 is **b17b0d77a685c4846e7ff7df5766a7fe2d197140d2330b8d1dc2d5fad33dd05b**; independent audit02 SHA256 is **b77d4ffcc3259ea7b45ae700281990875b52cc8088922808c0abfa841120043e**. Machine evidence seals both failed trials and their corrections.

**AC-088 BLOCKED:** G01-G03 reports and matching output/log audits pass their stated scope. The user then asked to combine/cut the16-row program or automate; D13 records a four-flow actual human route in `EXPLORER_SHORT_STEPS.md` / `EXPLORER_SHORT_RESULTS.md`. R01 combines no-input/no-outline/manual-multiple starts; R02 is EPUB; R03 is AZW3/manual/generated-PDF open; R04 is two-file rejection. All four remain NOT RUN. Original guides and G04-G16 remain superseded history, with no fabricated PASS. Removed cancellation/conversion-failure/normalization/selected-level/same-mode-repeat variants retain native evidence only; human repeat is the three actual different-mode runs with prior files preserved. Acceptance IDs and safety outcomes remain; TESTING/task/template/AC-088 routing is updated explicitly. Actual UI/viewer behavior is never inferred from automated controls or hashes. Historical M3 human testing remains absolutely closed.

No shared full22/Python/Pester suite was rerun in this evidence-only task; T02 history stays distinct. No clean OS, OS network-isolation, human approval, M5-T04 security/advisory review, M5-T05 integration or M6 tag/signing/publication/anonymous-download claim is made.

## Git checkpoint and continuation

Active branch/upstream: `codex/winbooksplit-v1-m5` / `origin/codex/winbooksplit-v1-m5`; origin remains PikkuJanne/WinBookSplit. C is already user-merged main. Its [main CI](https://github.com/PikkuJanne/WinBookSplit/actions/runs/38078153689) was observed completed/success at C; this task inspected metadata only and makes no new CI artifact audit claim. The evidence checkpoint's own actual commit, ordinary push, fresh clean/live equality and scoped draft PR receipt follow externally/thread after those actions; no self/future SHA is embedded.

Next action remains **M5-T03 / AC-088**. Reinspect live Git/PR/check state, rehash the three assets and candidate/runtime/input identities, collect actual Explorer checklist outcomes, authenticate the resulting logs/manifests/physical outputs, record defects and rerun only affected checks. Keep task blocked until every mandatory current interaction is evidenced. Preserve both retained automated workspaces and all historical failures; advance only after a synchronized completed task checkpoint.
