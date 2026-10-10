# PDF page fidelity and navigation

WinBookSplit copies PDF page content and resources; it does not render text pages
into images or add OCR. Image-only scans can be split with manual physical start
pages. Automatic modes require actual PDF outline destinations.

Each chapter receives a new document title and a bookmark pointing to its first
page. Their text uses the safe chapter filename label, without its sequence
prefix or `.pdf` suffix; it may be shortened or normalized for Windows filenames.
The source author is retained when it is a readable text value. The source title
is recorded as a `Source: ...` subject when available. Source text is bounded and
control characters are normalized. Creator/producer identify WinBookSplit and
its PDF library. Other source metadata, signatures and the original document
catalog are not promised to survive.

Supported local page links are rebuilt against the pages in each chapter. This
includes tested explicit `/Dest`, `/GoTo` action and named destinations that
resolve to a page inside that chapter. Destination coordinates and fit modes
are retained for the tested valid forms. A destination in another chapter is
dropped and produces `cross_chapter_link_dropped`; an invalid or unsupported
destination produces `navigation_link_dropped`. The chapter-start bookmark is
new; the complete original outline and named-destination catalog are not copied.
Article navigation is excluded with `article_navigation_dropped`.
Additional page actions are omitted with `navigation_link_dropped`; their page
references are not copied into another chapter.

Tested static `/Text`, `/Highlight` and `/Square` annotations retain their
appearance data. Unsupported
or malformed annotations and excluded relationships produce categorized
warnings. Popup/reply relationships and document actions are not a general
preservation promise. Interactive forms, signatures, attachments, portfolios
and active document behavior are outside this fidelity claim. Do not depend on
the splitter to preserve them or to sanitize untrusted documents.

The acceptance fixtures are original generated documents containing text,
vectors, embedded images, genuine image-only pages, rotations, crop boxes,
mixed sizes, static annotation appearances and local links. Structural checks
compare every selected page and its resources; same-renderer checks compare
source/output pixels using Poppler and PDFium as development tools. These
fixtures establish the documented cases, not compatibility with every PDF.
The application does not require either renderer.

Warnings appear in preview/results and local run records. The explicit support
export includes only fixed warning categories and codes; author, chapter title,
annotation text and paths remain excluded. Source documents are immutable.
