# Approved review -> implementation

All 17 review improvements are mapped. Concrete defaults in SCOPE_AND_DECISIONS.md choose strict starts (not ranges), encrypted-feature rejection (not password support), no-overwrite outputs and tested Windows11 support. These were explicit permissible choices in the review, not missing work.

| Review item | Improvement | Tasks |
|---|---|---|
| 1 | Strict manual splitting | M1-T02 |
| 2 | No omitted/misassigned bookmark pages | M1-T03, M1-T04, M1-T05 |
| 3 | Protect existing files/separate runs | M2-T01, M2-T03 |
| 4 | Subprocess streams/failures | M2-T05, M3-T01 |
| 5 | Literal paths/Unicode | M2-T02, M2-T05 |
| 6 | Dependency preflight | M0-T03, M2-T04 |
| 7 | Remove shared temp engine | M1-T01, M2-T01 |
| 8 | Split preview | M1-T05, M3-T04 |
| 9 | Explicit safe ebook behavior | M2-T03, M4-T04 |
| 10 | Keep/polish launcher | M3-T03 |
| 11 | Small noninteractive CLI | M3-T02 |
| 12 | Naming/logs/manifest | M2-T02, M3-T05 |
| 13 | Difficult PDF handling/fidelity | M4-T01, M4-T02 |
| 14 | Regression suite | M0-T02, M0-T04, M4-T03 |
| 15 | Windows CI/release packaging | M4-T05, M5-T02, M5-T03 |
| 16 | Accurate docs/license | M5-T01 |
| 17 | Version/single published release | M5-T01, M6-T01, M6-T02, M6-T03, M6-T04, M6-T05, M6-T06 |

Deferred by scope: website/deployment, GUI rewrite, native EPUB/AZW3 output, OCR/AI, DRM removal, installers/updaters and paid signing. These are not required to finish the project and must not become hidden M7 work.

## M5-T05 closure

All17 rows were audited against implementation, supported rejection policies and actual evidence on 2026-10-11 (Europe/Berlin). [Closed acceptance matrix](evidence/M5-T05-acceptance.json) records AC001-092 and AC093/094 within their scopes; [review](evidence/M5-T05-acceptance.md) records integrated candidate/current CI. Row17 publication/download/final closure remains M6, AC095-107 pending. Historical reproduction and superseded human rows are not current fresh passes.
