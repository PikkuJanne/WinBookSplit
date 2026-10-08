# PDF fidelity, support, and explicit limits

## Required v1.0.0 support

Ordinary unencrypted PDFs, including image-only scanned pages, must support manual starts and bookmark splitting when actual outline data exists. Copy page objects rather than rasterizing. Test text, embedded images, mixed sizes, rotations/crop boxes, ordinary annotations and non-ASCII metadata. A valid split plan doesn't prove readable output: reopen every output, verify count/order, and compare representative rendered pages visually. Development-only renderers/fixture generators do not become runtime dependencies without justification.

Set useful new document metadata: chapter title, source author when safely available, tool/producer information without falsely attributing a signature, and a chapter-start outline entry. Do not blindly clone the original document catalog. Rebase intra-chapter navigation only where the chosen pypdf APIs and tests support it; explicitly warn/document cross-chapter destinations that are dropped or become unavailable. No general promise that 'all metadata/bookmarks/links survive'.

## Conservative feature policy

Detect/reject encrypted PDFs in v1.0.0 before page extraction or opening any output stream. Do not attempt even empty-password decryption to bypass the declared policy. No plaintext password arguments/logs or extra AES support is required; encryption support would be a separate tested feature [S6].

Detect/reject interactive AcroForm/XFA and signed documents when the splitter cannot safely preserve the required behavior. Digital signature preservation is not promised. Embedded attachments, document-level scripts/actions, portfolios and unsupported active features are not copied silently; either reject with an unsupported-feature reason or implement a narrow, documented exclusion with visible warnings and acceptance tests. Never execute document actions to inspect them. Do not call WinBookSplit a sanitizer or security sandbox.

Broken outlines, invalid bookmarks and no-bookmark PDFs are handled at planning, with explicit warnings or manual fallback, not mislabeled 'corrupt PDF'. Truly unreadable/truncated files fail clearly. A zero-page PDF is invalid for splitting. Limit pathological recursion/resource consumption and offer cancellation; do not promise unlimited-size inputs or complete hostile-PDF defense.

## Proof of fidelity

Generated test pages should carry a stable visible/text page ID, unique sizes/markers as appropriate, plus source/output hashes. For each supported fixture, validate page identity/order and first/last inclusion. For representative visual cases render input/output with the same renderer and compare; human inspection catches issues not captured by page counts. Use a second renderer for difficult cases when claiming support, not as a blanket dependency requirement. Never commit real textbooks, workplace documents or private identifiers. Fixtures are generated or redistributable with provenance.
