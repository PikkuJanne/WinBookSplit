# Output safety and ebook conversion

## Lifecycle

Validation -> dependency preflight -> owned working directory -> conversion if needed -> probe -> plan -> preview/confirmation -> stage slices -> reopen/validate every slice -> publish final directory -> final manifest/summary -> cleanup.

Do not modify, rename, delete, or write beside the source. Resolve both requested source and destination; reject an output that aliases a source. Default to a new child of the real Documents directory, not a hardcoded English folder path. An explicit `-OutputDirectory` is a base directory. Append a sanitized book stem, a timestamp, and an unguessable/random run component; reserve a unique run directory atomically and retry collisions. Never reuse an earlier run's folder. An atomic same-volume rename into a not-yet-existing final path can publish a completed staging directory; no replace/overwrite operation is allowed.

Place output staging on the same destination volume as the final directory. Conversion work may be in that private staging/work location so large intermediate data is accounted for. Use explicit ownership markers and a manifest of created paths. Cleanup may touch only the current run's known children. Reject unexpected symlinks/junctions/reparse traversal; do not follow paths from untrusted documents or user-edited manifests into deletion targets. Atomic publication reduces partial exposure but is not a crash-proof transaction on every filesystem: claim only what Windows tests demonstrate.

During normal failure/cancel, stop only owned child processes, wait for shutdown, close handles, then remove owned partial output. Preserve a bounded diagnostic record with `failed` or `cancelled`, in an unmistakably separate diagnostic location; never leave final-looking chapter files labeled success. Hard termination/power loss can leave a marked staging folder; never auto-delete arbitrary old folders. Document safe manual recovery. A failure writing the completion manifest or promoting the final folder is a failure, not success with missing evidence.

## No destructive options

No overwrite switch, destructive cleanup switch, or implicit resume in v1.0.0. Existing neighboring PDFs and all earlier run folders remain unchanged. Copy-to-new behavior is intentional even when inputs share basenames. Validate each written PDF's expected nonzero page count and plan membership; test source page identities and rendered fidelity separately.

## Filenames

Use `width = max(2, len(str(section_count)))`, then `<sequence> - <safe-title>.pdf`. Clean Windows-invalid/control characters, whitespace and terminal dots; keep Unicode text. Fall back to a stable title such as `Section 3` for an empty result. Handle reserved names defensively without claiming the prefixed baseline necessarily hit them. Truncate titles based on the **full destination path budget**, preserving sequence/extension. Detect case-insensitive collisions; never use a title as a path. Too-long/unwritable paths fail with actionable errors, not registry/system-policy edits. Test 100+ outputs.

M2-T02 implements this policy with destination-bound preparation; all modes
freeze exact filenames before preview and reject casefold collisions/reserved
components. Empty/punctuation-only titles receive stable fallbacks; usable Unicode
and emoji are retained. Bound preparation may shorten titles/run stem; unbound
plans keep their names or reject. Limits are conservative UTF-16 budgets:
files 259, created directories 247, components 255, including stage/final and ownership/
manifest/failure records. Changed bound bases reject before allocation. These
limits do not promise arbitrary long paths or UNC and require no settings changes.
See `../evidence/M2-T02-paths.md`; conversion requirements below remain separate.

## Conversion

Resolve an explicit `-CalibrePath`, then trusted PATH discovery and known installation locations. Print/log the resolved executable and version. PDF-only runs do not probe/require Calibre. Do not look for executables in the document directory or download them silently. Use an executable plus correctly marshalled arguments, not shell-evaluated text.

Keep original ebook identity distinct from the generated PDF identity throughout logs/plan. The generated full PDF lives in an owned workspace; never `ChangeExtension` into the source directory. A zero exit code plus an existing filename is insufficient: check a newly created nonempty readable PDF with at least one page. Conversion failure, DRM rejection, missing output, zero-page output, cancellation and unwritable storage all propagate as failures.

Retain the current tested compatibility options initially (including evaluating the existing tablet output-profile behavior), record exact arguments, and verify them for EPUB -> PDF and AZW3 -> PDF with the chosen Calibre version. Do not assert a profile universally improves PDF layout. A small tested layout preset is permissible but not required. See source [S4].

`-KeepConvertedPdf` copies/retains the generated full PDF inside the successful run folder, never beside the input. Without it, remove the intermediate after success. Interactive manual selection may open the converted PDF only at the user's request; its preview location must remain valid until input is finished. Plan-only/preview runs may create temporary conversion data but never chapter outputs; document this distinction.

Keep processing offline after dependencies are installed. Do not promise a sandbox: pypdf and Calibre parse untrusted documents with the user's privileges. Avoid fetching linked external ebook resources; test and document actual converter behavior. Never add DRM-removal plugins.
