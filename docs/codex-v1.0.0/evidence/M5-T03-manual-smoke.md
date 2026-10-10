# M5-T03 — Actual candidate smoke matrix

**Partial record; AC-088 BLOCKED. G01-G03 PASS; four D13 flows R01-R04 NOT RUN.** The user authorized combining/cutting the original16 human rows. Original unrun G04-G16 are SUPERSEDED, never PASS; native coverage remains explicitly distinct. See [human record](M5-T03-human-smoke.md).

- ZIP: WinBookSplit-v1.0.0.zip, 413958 bytes, SHA256 `1f29981128c5fa74d17faca9e97d67f2605f008e80068d1f76809cbb79218645`.
- Source: `a282cf4412494063b91f2a71cb21b283731569f1`.
- Date/automated tester: 2026-10-10, Codex; separate native/asset auditor and agent visual inspection. Human G01-G03 tester: user directly in this chat,2026-10-10 Europe/Berlin; separate name not supplied.
- Isolation: existing Windows 11 Pro x64 workstation, fresh candidate extraction/venv per host and unrelated CWD; no clean OS/Sandbox/VM claim.
- Windows: build 26300.9457; PowerShell 5.1.26100.9444 and 7.6.5.
- Runtime: explicit application `.venv` CPython 3.14.8 regular GIL AMD64, pip 26.2.1, pypdf 6.19.0; installed distributions only pip/pypdf; no checkout imports. Actual Calibre 9.15.0 with explicit path for native tests, Poppler 26.07.0.
- Human preparation: verified existing Calibre application-only copy in the supported per-user known location; library/settings excluded, no installer/PATH/registry/policy changes. Four additional native default-discovery previews pass separately with `application_venv`/`known_location`; actual human dependency lines still must match. Original separate-Explorer PATH experiment is superseded.
- Fixtures: original MIT generated 12-page outlined/flat/parent-only PDFs, three-chapter EPUB, genuine Calibre-generated KF8 AZW3, neighboring canary and malformed originals. Exact hashes are in machine evidence. No private input or remote resource.
- Retained native work: `EXTERNAL/automated-candidate02`; current local human instructions/results: `EXTERNAL/EXPLORER_SHORT_STEPS.md` and `EXPLORER_SHORT_RESULTS.md`; the original `EXPLORER_STEPS.md`, `EXPLORER_RESULTS.md` and checklist remain preserved history. Exact paths/raw logs stay local. If human testing already began against retained run01's same-hash extraction, record that identity and its exact fixture/runtime hashes; do not discard those observations or silently conflate the runs.

| Template check | Actual observed automated result | Actual Explorer/human result |
|---|---|---|
| Extracted package Version without dependencies | PASS, actual native0 before `.venv`, both hosts | NOT RUN |
| Standard PDF Level1 drag/drop | Direct PS1 Level1 passes four complete ranges | **G01 PASS — user procedure confirmation plus authenticated4chapters/all12pages** |
| Nested Level2 parent/opening/front matter boundaries | Direct PS1 Level2 passes seven complete ranges | **G02 PASS — reported checklist result; authenticated7chapters/all12pages** |
| Manual 1; multiple starts; normalization; invalid tokens | Direct PS1 passes full12/three chunks/page1 addition/exit2 invalid token | **G03 Manual1 PASS; R01 multiple-start interaction pending; remaining variants native-only under D13** |
| No-bookmark/selected-level fallback | Redirected PS1 controls pass; missing outline noninteractive exit5 | **R01 no-outline interaction pending; selected-level variant native-only under D13** |
| Actual EPUB conversion and PDF chapters | Real Calibre; three physical pages/three chapters per host | **R02 NOT RUN** |
| Actual AZW3 conversion and PDF chapters | Genuine KF8; four physical pages/three chapters per host, final TOC preserved | **R03 NOT RUN** |
| Generated PDF open/physical page guidance | Converted/source physical ranges validated; no viewer launch exercised | **R03 NOT RUN — explicit generated-PDF viewer opt-in required** |
| No-file selection and cancel | Redirected blank input exit130; actual selection untested | **R01 selection pending; cancel variant native-only under D13** |
| Two dropped files explicitly rejected | No current automated BAT or Explorer multiple-drop claim | **R04 NOT RUN** |
| Spaces/brackets/Unicode/special filenames | Direct literal PS1 input `Input & [Å_日本] manual.pdf` passes both hosts | **G01 PASS — confirmed special-name PDF workflow** |
| Noninteractive CLI and Preview, both hosts | PASS actual CLI; four previews/zero publications | NOT RUN through Explorer; native PASS stands separately |
| Repeated run and old outputs preserved | Native same-mode repeat/hash preservation passes | **G01-G03 actual same-input/different-mode runs; all21 prior files preserved; same-mode GUI variant retired under D13** |
| Failure/cancel nonzero exit, no final-looking partial success | Direct native expected2/4/5/130; zero publications, owned converter diagnostics | **Native-only variants under D13; no inferred BAT exit from window disappearance** |
| Representative first/last page identities/rendered fidelity | PASS all230 output pages reopened;12 pixel matches; six final output agent views | **G01 markers confirmed; G02/G03 checklist PASS reported; R02/R03 ebook views pending** |
| Logs/manifests accurate and support export redacted | PASS independent30 finalized logs/22 manifests/two redacted exports | G01-G03 actual finalized logs/manifests independently authenticated; R01-R04 pending |
| Unrelated working directory/fresh venv, no checkout dependency | PASS actual native controlled dependency paths/two fresh runtime-only venvs | G01-G03 actual child resolver identities authenticated; R01-R04 must be recorded |
| Publicly downloaded same-hash ZIP minimal smoke (M6) | NOT RUN; local unpublished candidate | NOT RUN; belongs to M6 |

## Current actual human route — D13

The user explicitly requested combine/cut or automate the16-row program. Follow `EXPLORER_SHORT_STEPS.md`, not the preserved original guide. Existing G01/G02/G03 reports and output audits remain PASS; original unrun G04-G16 become SUPERSEDED, never PASS. The complete original-to-current/native mapping is in the machine human record.

| Current row | Actual remaining workflow | Outcome |
|---|---|---|
| R01 | Double-click BAT without input; select quoted No outline.pdf; Auto1/manual fallback; starts1,4,7; confirm/decline-open; ranges1-3/4-6/7-12 | NOT RUN |
| R02 | Drop real EPUB; Auto1; conversion and first/last output view; three physical pages | NOT RUN |
| R03 | Drop genuine AZW3; Manual; generated-PDF viewer opt-in/physical guidance; starts1,2,3; finalTOC preserved | NOT RUN |
| R04 | Drop two authored files together; explicit rejection/no processing | NOT RUN |

R01 replaces G04/G07/G12; R02 replaces G09; R03 replaces G10; R04 replaces G14. G05/G06/G08/G11/G13/G15 retain their existing native coverage only. G16's identical-mode repeat is native-only; human repeated-input safety is the three completed different-mode runs with previous outputs preserved. No removed GUI observation is inferred or passed.

The four flows preserve remaining PDF/ebook/no-input/multi-input/manual/fallback/viewer coverage. Automatic audits check every output page/log/dependency; human reports still supply actual UI/viewer observations. AC-088 remains blocked until all four current flows are evidenced. This authorized evidence-route change does not waive package, CI or release gates.

AC-087/089 native and agent-render evidence passes only its stated scope. AC-088 remains a mandatory blocker. Historical human M3-T01 is absolutely closed; this current candidate checklist does not reopen it. No clean OS or general conversion/release certification follows from these authored local fixtures.
