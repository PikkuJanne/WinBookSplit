# One split plan, shared by preview and execution

## Fundamental representation

Internally use zero-based half-open ranges `[start,end)`. Display physical pages as `start+1` through `end` inclusive. Plan entries have `sequence`, `title`, `start`, `end`, `parent_id` (nullable), `reason`, `filename`, and warnings; the plan has total pages, source identity, mode, normalized inputs, and coverage summary. These are planning requirements, not a second general-purpose framework.

For any accepted document with N > 0 pages: first start = 0, final end = N, every entry satisfies 0 <= start < end <= N, and each next start equals the previous end. The union is all N pages and there are no overlaps. Verify page **order/identity** too, not only the sum of lengths. Zero-page inputs fail. User-requested omissions are deferred, so no coverage waiver exists in v1.0.0.

Build/validate once, preview that plan, execute that same plan. Any input/conversion change invalidates it. Keep an open reader or verify identity/content before execution to avoid stale previews. Do not infer chapters from text or OCR.

M1-T05 retains the existing logical mode planners and adds a common callable
contract in the shipped engine. `prepare_split(path, mode, manual_data)` probes
the PDF, validates one complete plan and returns a `PreparedSplit` bound to
pypdf's captured in-memory reader. `preview_plan(prepared)` exposes the same
recursively read-only mapping/tuple data without creating chapter files;
`execute_split(prepared, output_dir)` consumes its exact entries without
recalculating boundaries, titles or filenames. UI integration belongs to M3.

The common plan includes `mode`, `total_pages`, `entries`, derived `ranges`,
`normalized_inputs`, `notices`, `warnings`, `source_identity` and derived
`coverage` (`complete`, `covered_pages`, `section_count`). Each entry retains
its existing mode metadata. `validate_plan(mapping)` rejects invalid page
types/bounds, gaps, overlaps, empty entries, incomplete coverage, nonconsecutive
sequences, unsafe/duplicate filenames or disagreement with declared ranges;
it freezes a detached copy. The result reports `written_count` and the same
entry metadata in `outputs`, adding each expected `page_count`.

Source identity records the requested/resolved PDF path, captured byte count
and SHA-256 with `binding: reader_snapshot`. Replacing, deleting or repointing
that filesystem path after preview cannot change the prepared reader or plan;
execution safely uses the captured original PDF. Preparing a different source
or newly converted PDF creates a new job. Execution rejects a replaced plan,
a reader from another job or altered captured stream before writing. This is
source binding, not a sandbox against arbitrary Python memory manipulation.
All planned paths are checked for source aliasing/existing outputs before the
first writer; exclusive file creation also refuses later collisions. Complete
staging, rollback, directory ownership and publication remain later tasks.

## Manual starts

Accept a nonempty comma-separated string of ASCII `[0-9]+` tokens, allowing surrounding whitespace. Integers must be in 1..N. Leading zeros are allowed and canonicalized. Reject signs, negative numbers, non-ASCII numeric glyphs, empty tokens, ranges (`1-4`), decimals, letters, or any invalid/out-of-range token; do not silently filter.

Only after validating every token: sort/dedupe starts, add page 1 if absent, convert to zero-based indices, pair with next boundary/N. Inform about sorting, duplicates, or adding the first page in preview. `1` on a 10-page PDF yields `[0,10)` with a 'one section; no internal split' notice and one validated output. `4,7` and `1,4,7` both yield `[0,3),[3,6),[6,10)`. Empty input is not an instruction to copy.

## Bookmark normalization

Collect usable internal destinations with titles, source outline order, depth and parent lineage. Validate finite integer pages in 0..N-1; skip invalid/external destinations with diagnostic warnings, never a bare `except: pass`. Keep a traversal bound/visited guard for malformed structures. Preserve deterministic first-in-outline-order selection when several bookmarks share a page, recording aliases/ignored duplicates. Sorting for physical order is allowed with a warning; do not invent chronology from titles.

## Level 1

M1-T03's implementation bounds outline and destination-name retrieval to depth
64 and 10,000 raw nodes per tree; named definitions are also capped at 10,000.
Normalized traversal counts both destination entries and child-list containers
against 10,000. Cycles, reused containers/contradictory lineage, unreadable trees
and exceeded bounds reject the request before writing. Safely recoverable
invalid destinations are warned and skipped while retaining their lineage.

Sort valid top-level starts; dedupe positions deterministically. Parent interval end is next top-level start or N. Add `[0, first_parent_start)` as Front matter when nonempty. Emit each parent interval using its title. No valid Level 1 outline yields structured `no_bookmarks`/`no_usable_bookmarks`, not success with zero files.

## Level 2

Construct the Level 1 parent intervals first; do not flatten every depth-2 bookmark into one global sequence. Within each parent, select its own valid direct children, deduped and physically ordered. A child outside its parent interval is ignored with a warning, not used to cut a neighboring parent.

When first child starts after parent start, emit `[parent_start,first_child)` as `<Parent> - Opening pages`. Children split only until the next child or that parent's end. A parent without usable children is one fallback segment carrying the parent's title. Preserve front matter as above. At least one usable Level 2 child somewhere is required for Level 2; otherwise report `no_bookmarks_at_level` and offer Level 1 interactively. A non-interactive call returns the documented no-plan code without fallback. Parents with unusable starts cannot define reliable intervals: warn and omit their child hierarchy; never orphan those children into another parent. Reject a plan if coverage cannot be made unambiguous.

Duplicate parents at the same destination use the first parent and its subtree; warn about ignored alias subtrees. Depth 3+ is not a split boundary in Level 1/2. A malformed tree or contradictory parentage is a diagnostic condition, not a mandate to recurse indefinitely.

M1-T04 uses the normalized Level 1 parent IDs to select direct children before
deduplication/sorting. Ignored alias/invalid parent subtrees remain in metadata
with diagnostics; none of their descendants can satisfy the usable-child gate.
Outside children cannot satisfy it either. With usable parents but no retained
direct child, the noninteractive CLI emits `no_bookmarks_at_level`, the existing
`[NO_BOOKMARKS_FOUND]` compatibility sentinel and exit 55 without writing or
falling back. Empty/no usable parent errors remain distinct. The console's
interactive fallback choice belongs to M3; this does not certify that UI.

## Exact example: 12 physical pages

Parent A starts at 3, with children A1 at 4 and A2 at 7. Parent B starts at 9, with child B1 at 11.

Level 1: `[0,2), [2,8), [8,12)`.
Level 2: `[0,2), [2,3), [3,6), [6,8), [8,10), [10,12)`.

Physical display: front matter 1-2; A opening 3; A1 4-6; A2 7-8; B opening 9-10; B1 11-12. A2 must not absorb pages 9-10. `PLAN_ORACLES.json` holds this and other expected outcomes.

## Documenting limitations

Page boundaries refer to the PDF's physical sequence, not printed roman/arabic page labels. Ebooks are reflowed by conversion: show/open the converted PDF before manual starts. Invalid bookmarks produce warnings with location/reason. Images/scans with usable bookmarks can split automatically; OCR is irrelevant to outline existence.
