# Original PDF page fidelity acceptance

`characterize_fidelity.py` runs eight actual redirected Windows application controls: rich Manual/Level1/Level2 and genuine image-only Manual, under each supported PowerShell host. `tests/run_tests.py --layer fidelity` is also stage 19 of full. Pass both `--shell-path` values, `--renderer-path` for the selected Poppler `pdftoppm.exe`, and `--secondary-python-path` for the isolated developer Python containing PDFium. The runner supplies an exclusive external `--render-directory`; retained source/output PNGs remain available for separate visual inspection after owned fixture cleanup.

The original MIT fixtures use pinned developer ReportLab/Pillow/pypdf. They contain text, vector strokes, embedded images, image-only pages, stable physical page IDs, mixed page/crop/trim/bleed/art boxes, four rotations, static Text/Highlight/Square appearance streams, non-ASCII metadata and direct/named navigation. Forward/back intra-segment links, cross-segment links, null/unresolved destinations and an authored cross-page annotation `/P` distinguish page-tree correctness from hidden serialized-page/resource leakage. No private source document is read or committed.

Every output is reopened. Exact text/content/image/geometry/static appearance snapshots, first/last/order coverage, all serialized Page objects, all serialized content markers and every serialized image stream must bind to the selected physical pages. Safe chapter/source metadata and a page0 chapter-start outline are required. Supported direct/named intra-segment destinations must point to actual output page references and preserve their resolved fit and coordinate arguments; cross-segment/invalid links must be absent with visible fixed warnings. Each host's rich Manual run is exported through the actual local diagnostics exporter and checked for category/token-only redaction.

Both separate native renderer routes render source and output at the same scale. PNG bytes, dimensions and decoded RGBA pixels must match exactly within each renderer. Renderer probes, versions, executable/native-library/script hashes, exact argv and process results are recorded. Automated pixel comparison is distinct from human/GUI/Explorer/viewer testing; no such test, clean OS, CI or release-package claim follows.

The strict pure validator checks source/application/neighbor/prior identity, finalized log/run/output bindings, nonempty physical plans, actual job/stream completion, no-prompt execution, and authenticated owned cleanup. Receipt mutation tests are synthetic oracle controls, never new native application passes. A deadline or unexplained failure retains the authored workspace and records cleanup as unsafe. Persistent render artifacts are deliberately outside the removed fixture workspace and never overwritten.

Example (verified runtime placeholders must be replaced):

```text
<DevPython> -I -B tests/run_tests.py --layer fidelity --report <ExternalReport> --shell-path <PS51> --shell-path <PS7> --renderer-path <pdftoppm> --secondary-python-path <BundledPython>
```
