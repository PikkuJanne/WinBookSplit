# M4-T01 — ordinary PDF fidelity and chapter navigation

The scoped implementation has passing targeted and affected evidence. The
standalone 19-stage full run **failed**, outer native1, with 18 stages passing.
A fresh process-only retry on the same C passed native0. Together these provide
19 passing layer observations; they do not create a successful standalone full
run. Final independent source, retained-child, process and fidelity reviews pass
with no scoped blocker; they preserve the original failed full result.

Implementation C is `84ff807461409392ff2743480e9a175d0da5398e`, tree
`2557b20d76d89c9b06c4bd2d8407766dfbe7cdbb`, on `codex/winbooksplit-v1-m4`.
All 125 mapped raw source paths match committed Git bytes; their source digest is
`ea70ee9eae682e6a5b73086a94531bf84a1148d8a945edfe259d0366662f6418`.
Clean local C equalled the live branch at `2026-10-10T09:08:27.953349Z`.
[Draft PR #22](https://github.com/PikkuJanne/WinBookSplit/pull/22) is the M4
integration vehicle. Main remains pre-task `f0c86a8908ee854f78b029a1c0046d05b1f8fb06`;
M4-T05 owns cumulative milestone review, CI and merge. No tag or release was
created. A later documentation checkpoint must receive its own actual push/live
receipt; no future checkpoint SHA is claimed here.

Original synthetic fixtures reproduced missing chapter metadata/start outline
and outside-chapter content serialized through annotation references. Rich
three-page baseline outputs contained eight serialized Page objects and all six
source page markers/image streams. Direct pypdf probes confirmed ordinary
`add_page` can copy outside-page references; `append(import_outline=False)` still
imports AcroForm catalog data. Native0 in these baseline probes did not mean the
result was correct.

The shared validated plan, captured reader and owned publication path remain in
use. A shallow page view excludes source annotations, article references and
additional page actions before copying selected content/resources. The writer
rebuilds the permitted Text, Highlight and Square annotations and valid local
direct/GoTo/named destinations against chapter pages. Tested fit/coordinates
remain intact. Malformed, unresolved, unsupported and cross-chapter navigation
is omitted with fixed warnings in the prepared plan before chapter writes.
Source pages, annotation dictionaries and input bytes remain unchanged.

Each chapter receives a safe filename-derived title and first-page bookmark,
optional bounded source author/title attribution, and WinBookSplit/pypdf
creator/producer. No original catalog, signature or complete outline promise is
made. Explicit support exports contain fixed warning codes/categories/counts,
excluding source text, metadata values, annotation text and paths.

| Acceptance | Actual fidelity evidence |
| --- | --- |
| AC-066 | Selected page identities/order, extractable text, vectors and images preserved; genuine image-only manual output works. |
| AC-067 | Mixed sizes, boxes, rotations and permitted static appearances match structurally and in same-renderer images. |
| AC-068 | Safe chapter/source metadata and first-page bookmark match the declared contract. |
| AC-069 | Valid local direct/named destinations and fit/coordinates remap; unavailable navigation is warned and omitted. |

Targeted02 and the passing fidelity stage in full01 each exercised eight actual
application controls and two actual diagnostic exports under PowerShell
5.1.26100.9444 and 7.6.5. Each run reopened 20 chapter PDFs and checked 44 physical
output-page observations, retaining 176 source/output PNGs. Same-renderer PNG
and RGBA comparisons matched. Rich fixtures cover Manual, Level 1 and Level 2;
image-only fixtures use manual physical starts. Synthetic inputs and neighbors
were preserved, and owned temporary fixture work was removed as recorded.

Pinned processing runtime was Python 3.14.8/pypdf 6.19.0. Developer renderers were
Poppler 26.07.0 and PDFium 153.0.7999.0 through pypdfium2 5.13.0/Python 3.12.14.
Comparisons are source versus output within each renderer, not equality between
different renderers. Neither renderer is an application dependency. The root
assistant inspected six representative targeted02 source/output pairs, 12 PNGs,
using `view_image`. Human testing remains closed; no human or GUI viewer pass is
claimed. Targeted02's retained PNGs were independently reread and rehashed.

Targeted02 used an unchanged earlier worktree map. All 15 application files and
123 of its 125 mapped paths match C; two final test-only oracle paths subsequently
added typed/versioned export-frame guards. The final strict oracle revalidated
the retained targeted02 rows, and 13 focused synthetic methods passed native0.
Current engine evidence also passes 11 fidelity and 12 inherited manual methods.
Independent fixed callable probes confirm malformed Rect/QuadPoints, hidden
action and page-AA exclusions prevent the reproduced outside-page clone paths,
while source snapshots remain immutable. These are bounded ordinary-PDF checks,
not a general document sanitizer or unsupported-feature policy.

Full01 passed 330 Python methods and 112/112 Pester methods on each host, with no
skipped/not-run Pester methods. Its passing layers include real Calibre 9.15.0,
72 CLI, 70 launcher, 28 UX and 66 support controls, plus fidelity. The process
stage alone returned native1, stderr `Native exit status changed`, without a
child receipt. The failing case and cause are **unknown**; no flake, concurrency
or production-root-cause explanation is inferred. The aggregate stays FAILED.

The fresh same-C process01 retry passed all 35 actual controls: 26 supervisor
and nine application controls, under both supported hosts. Both reports bind
the same 125-path C map, source/synthetic-input stability and owned outer cleanup.
This yields composite targeted/affected coverage of 19 layers while preserving
`successful_standalone_full=false`. Independent review verifies the retained 18
passing layers, fresh process controls and identical C source maps. A separate
independent full01 fidelity/support audit rechecks all eight app controls, two
exports, 66 support controls and all 176 full01 PNGs without a native rerun.

| Retained report | Actual status | SHA256 |
| --- | --- | --- |
| full01.json | FAILED; outer native1; 18/19 stages passed | `d2ef77974d0fea3ce129b1677596bb0e318ecccf062356a0ac1d8f46090d46b1` |
| process01.json | PASS; outer native0; 35 controls | `b68e5e4efa942531972a455b64a32b3886c6b235ccdfa9816654fea377ab29c0` |
| independent-C-composite-coverage-review01.json | PASS; composite review, original full failure preserved | `b098575b49600355403097ade17be559134adb47a4120d6b5a33501f39bfd1e4` |
| independent-failed-full-retained-review02.json | PASS; strict retained-child review | `3ff8ed193dc1917a768db09d40ed89dbc59236bed124cb3064b3dccfde65d5ec` |
| fidelity-full01-independent-audit01.json | PASS; full01 fidelity/support/PNG reread | `a7a58747f7a10b5f4b536296b65ebb98723a374198912af0edb8033c72b79574` |

Preserved history includes early local baseline assertions that incorrectly
assumed intra-links were unresolved, initial `/P`/source-reopen test mistakes,
and the failed first native fidelity target. That target had an actual
application0/export0 but failed its ordinary-NI process oracle; retained-case
revalidation exposed the inherited bookmark-only support category oracle.
Both narrow test adapters were corrected; failed ancestors remain failed.
A focused recorder mislabelled native0/nine OK methods as FAILED/zero because
its regex missed CRLF. An additive correction records nine methods without a
rerun or overwritten ancestor. The default whitespace native2 observation
remains separate from the CRLF-aware native0 check. Full01's process failure is
also retained, separate from process01's successful retry, for broader M4-T05
CI/review. No original process failure cause or resolved diagnosis is claimed.

Redacted executed command shapes substitute selected developer tools and owned
evidence paths. Private receipts retain the exact argv.

```powershell
& $DevPython -I -B tests/run_tests.py --layer full --report $FullReport --tool-root $ToolRoot --shell-path $PS51 --shell-path $PS7 --calibre-path $Calibre --renderer-path $Poppler --secondary-python-path $RendererPython
& $DevPython -I -B tests/run_tests.py --layer process --report $ProcessReport --shell-path $PS51 --shell-path $PS7
& $DevPython -I -B tests/run_tests.py --layer fidelity --report $TargetReport --shell-path $PS51 --shell-path $PS7 --renderer-path $Poppler --secondary-python-path $RendererPython
& $DevPython -I -B -m unittest discover -s tests/python -p test_fidelity_receipt.py -v
```

No private source documents were used. These fixtures do not establish universal
PDF compatibility, OCR or preservation of forms, encryption, signatures,
attachments, portfolios and active catalog behavior. M4-T02 owns broader
unsupported-document policy. M4-T05 owns cumulative review/CI/merge; this evidence
does not establish clean-machine, new human, release-package or publication
acceptance.
