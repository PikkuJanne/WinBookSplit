# M4-T04 — real EPUB/AZW3 evidence

AC-076/077/078 pass with original authored books, actual Calibre and both supported Windows hosts. Human M3-T01 stays absolutely closed. The [machine record](M4-T04-ebook-conversion.json) binds commands, source, runtimes, raw seals, results and limitations.

## Source and verification

Clean C **5214bb230590b27d9399d22b80c1f0b5fead2d74**, tree **5adb04ede4f8a852fc3dbe0a60858d9040be1d0e**, contains **155 mapped raw files** identical to Git blobs, digest **4709450a19fe13bbdd121b7f4c73eb4adf94a2d94f5d293050b5bbf863b9e418**. All15 application/dependency paths match user-merged PR24 main **d360bfbff007dc6b5ed49a9e352f5079977b4f22**. Only tests and documentation changed; the shipped tablet compatibility option is unchanged. C was normally pushed and clean/live **SYNCED at 2026-10-10T14:01:17.184045+00:00**. E's own actual push/live receipt follows externally in the thread/PR after commit; no self/future SHA is embedded.

| Fresh command / acceptance | Actual outcome on clean C | External sealed records |
|---|---|---|
| Pinned Python `-I -B tests/run_tests.py --layer python` | Native0, **453 methods**, zero skips, source stable | `python-C01.json`, `python-C-command01.json` |
| Same Python `--layer ebooks`, explicit PS5.1/PS7, Calibre and Poppler | Native0, **2/2 stages** (inherited conversion + new companion) | `ebooks-C01.json`, `ebooks-C-command01.json` |
| Original EPUB3/TOC; independently generated genuine AZW3 | Four normal Auto1/Keep runs and four actual malformed failures across the two hosts | Same source-bound companion |
| Live loopback resource observation | Two actual EPUB canary runs, four positive same-endpoint GET controls; zero conversion requests | Same companion, native diagnostics and timestamp ledger |
| Independent source/native/public audit and agent render inspection | PASS within explicitly recorded scope | Review receipts in machine record |

Actual runtimes: Windows11 Pro x64 26H2 build26300.9457; PS5.1.26100.9444 and7.6.5; Python3.14.8/pypdf6.19.0; ReportLab5.0.1/Pillow12.3.0/charset-normalizer3.5.2; Calibre9.15.0 and Poppler26.07.0. Exact installed executable bytes and selected-format help are recorded. Online Calibre documentation may describe a newer version.

## Format, physical pages and behavior

The original EPUB3 has nine bounded archive members, first uncompressed mimetype, ordered three-chapter spine, matching HTML nav/NCX and closed local resources. Its actual generated KF8 AZW3 is independently identified by BOOKMOBI PalmDB, MOBI8 header, valid ordered records and no encryption; it is not a renamed EPUB. Original material is MIT and contains only unique authored chapter markers. No private ebook/PDF is processed or committed.

Actual independent tablet/default PDF references have EPUB3 physical pages and AZW3 four, both Letter612x792 with three flat chapter destinations0/1/2. AZW3 page4 is Calibre's generated TOC. Normal Auto1 ranges are EPUB `[0,1),[1,2),[2,3)` and AZW3 `[0,1),[1,2),[2,4)`: the TOC remains in chapter3, covering every physical page exactly once. Expected dropped cross-chapter link warnings are checked rather than silently discarded. Default references are diagnostics and do not justify changing the application profile.

The companion reopens **14 normal full-PDF pages / 12 chapters**, plus **6 canary pages / 6 chapters**. Every source/output page preserves extracted text, captured content, geometry and same-renderer PNG/RGBA pixels. Four independent references add **14 physical pages**. **54 PNGs and four reference PDFs** remain outside Git; application originals are inspected before authenticated owned cleanup. Agent inspection of retained representative pairs is separate from automated comparisons and human testing.

Four actual malformed EPUB/AZW3 cases reach real nonzero Calibre status, return app4/conversion_failed with actionable bounded diagnostics, and publish no plan/chapter. The inherited stage freshly covers eight actual Manual/retained/default executions under both hosts, two separate callable prepare/preview-to-execute parity controls with declined working-PDF opening, eight declared fake-invalid converter outputs and one BAT missing-converter control. Those are distinct native/callable/fake controls; no application Preview/plan-only or viewer interaction is claimed.

Literal conversion argv is `<resolved ebook-convert.exe> <original ebook> <owned generated PDF> --output-profile tablet`. Selected-format local help confirms the tested option. Each app case and the generation/reference scope uses new owned process-local config/cache/temp directories; inherited environment values restore afterward. Read-only original ebook identity/bytes/attributes, neighboring PDFs, prior files, earlier same-base publications and their console logs, copied app/tool bytes, stored execution policies and source maps remain unchanged. Removal authenticates only this run's expected owner/directory/file identities. Unexpected native shutdown or receipt rejection retains the owned workspace.

Self-contained fixtures operate with installed dependencies and no tool-initiated dependency installation/upload. The separate bounded canary supplies only an image and CSS on an exclusively owned live loopback listener. Actual Calibre reports blocking the image URL under each host. Zero requests to either exact endpoint occur during conversion; before/after same-endpoint GETs return expected bytes. There is no CSS blocking diagnostic, so its cause is not inferred. This does **not** prove OS network disconnection, all-egress denial or security for arbitrary ebooks/plugins. See [support notes](../../EBOOK_SUPPORT.md) and [route details](../../../tests/ebooks/README.md).

## Preserved trials and limits

The first new fixture-checker trial failed three negative CSS/inline-style subtests; its escaped/inline resource false passes are fixed before C. Its exact tested worktree digest was not captured and remains UNKNOWN; the known base is not substituted. Independent fully resealed receipt trials exposed nonboolean page/geometry aliases and foreign failure/config paths; narrow guards reject them now. Converter stream tails now reject contradictions with observed byte counts. These are new harness defects, not production conversion defects.

The initial newline helper stopped on a wrong all-CRLF assumption and its native status was masked by a subsequent shell command; it remains UNKNOWN. The first C proof found16 inherited files converted by Git checkout to CRLF. Guarded raw restoration changes only line endings to actual C blobs, and index refresh leaves the same clean commit; final proof03 passes. Failed first restore/proof attempts remain preserved. The PS5.1 baseline policy probe was partial; one correction used an unsupported module parameter and failed native1, then the explicit shipped-module correction passed native0 without policy changes. The focused behavior recorder's FAIL label came from CRLF summary matching despite actualnative0/3methods; additive raw review preserves both facts.

An initial synthetic receipt resealer trial failed native1 with one `KeyError run_id` among11methods; its traceback remains only in task tool history and is disclosed in unit01. Independent raw hash and exact source digest were not captured and remain UNKNOWN. Pre-C PDFs/canary/unit receipts remain on their own source identities. Historical T01/T02 full failures and T03 initial isolation/combined failures remain immutable; unknown historical process causes remain UNKNOWN. Fresh passes do not diagnose or relabel them.

No fresh22-stage full run or Pester is claimed here: only test/docs paths changed and all15 production paths remain identical. Historical T03 full21/21 and112/112 Pester per actual host remain separate evidence. No new human/GUI/Explorer/viewer, CI, clean-machine/Windows10/ARM/UNC, arbitrary book/DRM/native ebook output, package/signing/tag/assets/download or public release completion is claimed.

## Next checkpoint

Branch `codex/winbooksplit-v1-m4`, [scoped draft PR25](https://github.com/PikkuJanne/WinBookSplit/pull/25). Next: **M4-T05 — Add least-privilege Windows CI**, AC-079/080/081, including cumulative M4 review and passing integration/main sync. Stop at that boundary; preserve historical receipts and human closure.
