# M5-T03 — Actual candidate smoke matrix

**Partial record, AC-088 BLOCKED. G01 actual human procedure plus authenticated output PASS; G02-G16 NOT RUN.** See [current human record](M5-T03-human-smoke.md). This copies the manual smoke template with each kind of evidence explicitly labeled. [Native evidence](M5-T03-candidate.md) and [machine record](M5-T03-candidate.json) contain the observed automated results; they do not substitute for the missing Explorer observations.

- ZIP: WinBookSplit-v1.0.0.zip, 413958 bytes, SHA256 `1f29981128c5fa74d17faca9e97d67f2605f008e80068d1f76809cbb79218645`.
- Source: `a282cf4412494063b91f2a71cb21b283731569f1`.
- Date/automated tester: 2026-10-10, Codex; separate native/asset auditor and agent visual inspection. Human G01 tester: user directly in this chat,2026-10-10 Europe/Berlin; separate name not supplied.
- Isolation: existing Windows 11 Pro x64 workstation, fresh candidate extraction/venv per host and unrelated CWD; no clean OS/Sandbox/VM claim.
- Windows: build 26300.9457; PowerShell 5.1.26100.9444 and 7.6.5.
- Runtime: explicit application `.venv` CPython 3.14.8 regular GIL AMD64, pip 26.2.1, pypdf 6.19.0; installed distributions only pip/pypdf; no checkout imports. Actual Calibre 9.15.0 with explicit path for native tests, Poppler 26.07.0.
- Human preparation: verified existing Calibre application-only copy in the supported per-user known location; library/settings excluded, no installer/PATH/registry/policy changes. Four additional native default-discovery previews pass separately with `application_venv`/`known_location`; actual human dependency lines still must match. Original separate-Explorer PATH experiment is superseded.
- Fixtures: original MIT generated 12-page outlined/flat/parent-only PDFs, three-chapter EPUB, genuine Calibre-generated KF8 AZW3, neighboring canary and malformed originals. Exact hashes are in machine evidence. No private input or remote resource.
- Retained native work: `EXTERNAL/automated-candidate02`; current local human instructions/results: `EXTERNAL/EXPLORER_STEPS.md` and `EXPLORER_RESULTS.md`; original checklist retained. Exact paths/raw logs stay local. If human testing already began against retained run01's same-hash extraction, record that identity and its exact fixture/runtime hashes; do not discard those observations or silently conflate the runs.

| Template check | Actual observed automated result | Actual Explorer/human result |
|---|---|---|
| Extracted package Version without dependencies | PASS, actual native0 before `.venv`, both hosts | NOT RUN |
| Standard PDF Level1 drag/drop | Direct PS1 Level1 passes four complete ranges | **G01 PASS — user procedure confirmation plus authenticated4chapters/all12pages** |
| Nested Level2 parent/opening/front matter boundaries | Direct PS1 Level2 passes seven complete ranges | **NOT RUN** |
| Manual 1; 1,4,7; 4,7; invalid tokens | Direct PS1 passes full12, three chunks, page1 addition, exit2 invalid token | **NOT RUN** |
| No-bookmark/selected-level fallback | Redirected PS1 controls pass; missing outline noninteractive exit5 | **NOT RUN** |
| Actual EPUB conversion and PDF chapters | Real Calibre; three physical pages/three chapters per host | **NOT RUN** |
| Actual AZW3 conversion and PDF chapters | Genuine KF8; four physical pages/three chapters per host, final TOC preserved | **NOT RUN** |
| Generated PDF open/physical page guidance | Converted/source physical ranges validated; no viewer launch exercised | **NOT RUN — explicit viewer opt-in required** |
| No-file selection and cancel | Redirected blank input exit130; actual selection untested | **NOT RUN** |
| Two dropped files explicitly rejected | No current automated BAT or Explorer multiple-drop claim | **NOT RUN** |
| Spaces/brackets/Unicode/special filenames | Direct literal PS1 input `Input & [Å_日本] manual.pdf` passes both hosts | **G01 PASS — confirmed special-name PDF workflow** |
| Noninteractive CLI and Preview, both hosts | PASS actual CLI; four previews/zero publications | NOT RUN through Explorer; native PASS stands separately |
| Repeated run, old outputs and neighboring PDF unchanged | PASS direct native repeat/hash preservation, exclusive new publication | **NOT RUN** |
| Failure/cancel nonzero exit, no final-looking partial success | Direct native expected2/4/5/130; zero publications, owned converter diagnostics | **NOT RUN — no inferred BAT exit from window disappearance** |
| Representative first/last page identities/rendered fidelity | PASS all230 output pages reopened;12 pixel matches; six final output agent views | **G01 PDF markers confirmed; ebook/other human checks NOT RUN** |
| Logs/manifests accurate and support export redacted | PASS independent30 finalized logs/22 manifests/two redacted exports | G01 actual finalized log/manifests independently authenticated; remaining human runs pending |
| Unrelated working directory/fresh venv, no checkout dependency | PASS actual native controlled dependency paths/two fresh runtime-only venvs | Human child resolver identities must be recorded |
| Publicly downloaded same-hash ZIP minimal smoke (M6) | NOT RUN; local unpublished candidate | NOT RUN; belongs to M6 |

## Required actual Explorer continuation

The local checklist contains these concrete rows. **G01 PASS** is separately supported by the current human record; G02-G16 are **NOT RUN**. Procedure cells describe expected checks rather than inventing remaining outcomes. Record tester/date and row PASS/FAIL/BLOCKED, exact extraction/input identities and emitted output/log paths. Open first/last chapter PDFs for PDF, EPUB and AZW3; verify expected physical ranges and readable markers/text. Accept ebook runs only when their actual log binds the intended real converter/version and candidate-venv runtime. Any missing dependency or unresolved environment inheritance remains a failed/blocked check.

| Row | Actual action required | Current outcome |
|---|---|---|
| G01 | Drop special-name PDF onto BAT, Auto1, confirm, opt into output-folder open; ranges1-2/3-8/9-11/12 | **PASS — user confirms procedure/open/markers; exact output/log audit passes** |
| G02 | Drop PDF, Auto2; seven ranges1-2/3/4-6/7-8/9-10/11/12 | NOT RUN |
| G03 | Drop PDF, Manual1; all12 physical pages in one chapter | NOT RUN |
| G04 | Drop PDF, Manual1,4,7; ranges1-3/4-6/7-12 | NOT RUN |
| G05 | Drop PDF, Manual4,7; explicit page1 addition and complete coverage | NOT RUN |
| G06 | Drop PDF, invalid1,,4; visible rejection/no publication | NOT RUN |
| G07 | Drop no-outline PDF, Auto1; explicit manual fallback | NOT RUN |
| G08 | Drop parent-only PDF, Auto2; explicit fallback to level1 | NOT RUN |
| G09 | Drop real EPUB, Auto1; actual conversion/three complete physical pages | NOT RUN |
| G10 | Drop genuine AZW3, Manual; opt into generated-PDF viewer, inspect physical guidance, starts1,2,3 retain finalTOC | NOT RUN |
| G11 | Drop malformed EPUB; real failure/no chapter publication | NOT RUN |
| G12 | Double-click BAT without input; select quoted literal PDF path and finish | NOT RUN |
| G13 | Double-click BAT without input; blank input/Cancel produces no publication | NOT RUN |
| G14 | Drop two authored files; explicit multiple-input rejection without processing | NOT RUN |
| G15 | Drop PDF; cancel the proposed plan, no publication | NOT RUN |
| G16 | Repeat G01 and decline output open; distinct new publication/prior outputs unchanged | NOT RUN |

AC-087/089 native and agent-render evidence passes only its stated scope. AC-088 remains a mandatory blocker. Historical human M3-T01 is absolutely closed; this current candidate checklist does not reopen it. No clean OS or general conversion/release certification follows from these authored local fixtures.
