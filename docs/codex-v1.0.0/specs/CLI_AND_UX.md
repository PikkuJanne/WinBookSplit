# Interface contract

Keep the current console style, colors optional, and drag-and-drop launcher. No UI framework replacement. Use exact, validated menu entries: Level 1, Level 2, Manual, and explicit cancel. Invalid choices reprompt; default choice must be shown and cannot be inferred from a regex matching any 'm' or 'y' inside arbitrary text.

A no-argument interactive launch offers a file selector or a clearly prompted literal file path. If a Windows dialog is used, test/correct its apartment requirements for supported hosts. Cancelling selection returns cancel, not invalid-file success. The batch launcher describes PDF/EPUB/AZW3 and explicitly rejects more than one file. Do not silently process just the first dropped file. Direct CLI tests and real Explorer drag/drop are both required.

## Small parameter surface

Required proposed parameters: `-InputFile`, `-OutputDirectory`, `-Mode Auto|Manual`, `-BookmarkLevel 1|2`, `-StartPages`, `-Preview`, `-NonInteractive`, `-NoPause`, `-Version`. Supporting explicit runtime paths and conversion retention: `-PythonPath`, `-CalibrePath`, `-KeepConvertedPdf`.

These are **target** parameters; the baseline has only `-InputFile`. Implement binding/parameter-set rules and help before documenting examples as working. Auto with StartPages, Manual with BookmarkLevel, Preview with contradictory options, or absent required noninteractive values should fail early. `-Version` exits independently of input, runtime discovery and pauses. No unsupported `-Force`/overwrite or ranges parameter.

Suggested final behavior examples, to be executable-tested before documentation:

```powershell
.\WinBookSplit.ps1 -InputFile 'C:\Books\Manual.pdf'
.\WinBookSplit.ps1 -InputFile 'C:\Books\Manual.pdf' -Mode Auto -BookmarkLevel 2 -Preview -NonInteractive
.\WinBookSplit.ps1 -InputFile 'C:\Books\Manual.pdf' -Mode Manual -StartPages '1,4,7' -NonInteractive
.\WinBookSplit.ps1 -InputFile 'C:\Books\Book.epub' -Mode Auto -BookmarkLevel 1 -KeepConvertedPdf -NonInteractive
.\WinBookSplit.ps1 -Version
```

## Preview

Show original source, generated PDF where applicable, PDF physical page count, selected mode/level, numbered title/filename, physical inclusive start-end, section count, output base, warnings and exact page coverage. Confirm before interactive writes. `-Preview` prints a plan only and exits without chapter outputs; temporary conversion for ebook analysis is disclosed and cleaned safely. `-NonInteractive` never prompts, opens a file/Explorer, clears useful redirected output, or pauses. Missing selected bookmarks fails with an actionable code; no silent fallback. Manual input labels explicitly say physical PDF pages, not printed page numbers.

## Exit codes

Use a single documented mapping across Python, PowerShell and batch; proposed mapping: 0 = successful execution/explicit preview/version; 2 = input/arguments/path validation; 3 = dependency missing/incompatible; 4 = conversion failure; 5 = no usable plan at requested bookmark level; 6 = PDF/read/write/output-validation failure; 7 = unsupported document features; 130 = cancelled/timeout (distinguish reason in structured result). Existing internal code 55 may be mapped, not leaked inconsistently. Pin this contract in tests; precise alternatives require a decision update, not ad hoc codes.

M3-T01 implements the split-operation mapping in the shipped
`engine/WinBookSplit.Outcomes.json`, consumed by Python and PowerShell. `55` is
historical only. The final `[OUTCOME]` JSON record describes the whole console
operation and retains the validated final attempted engine result. Explicit
fallback cancellation returns 130; timeout has a distinct `timeout` status/code.
No missing, conflicting or zero-output engine result can become success.
Preview/version and full parameter binding remain M3-T02; reserved code 7's
feature rejection remains the PDF support work.

Console log flush/close precedes the success announcement. If chapter publication
has already succeeded but handle/log finalization fails, the operation reports
`incomplete` and exit 6, names the retained completed folder, and preserves its
validated execution evidence. A secondary finalization error preserves an
existing failure/cancellation code and appends its cause. The launcher explicitly
returns the saved PowerShell exit; a pause cannot replace the operation outcome.

## Completion, logs and accessibility

Completion names the actual final folder, section count and page coverage; zero-output execution fails. Offer to open the completed folder only interactively. A converted input keeps its original ebook name in diagnostics. Use readable text without requiring color/ANSI. Show progress by stage/section rather than an invented ETA. Logs and UTF-8 JSON run manifests include version, resolved dependency versions, run ID, mode, original/generated source identity, page count, ranges, output names, warnings and final outcome. Save paths locally; a redacted diagnostic export strips user/profile paths, document titles/content, credentials and command environment. No telemetry or automatic upload.

## M1-T06 diagnostic checkpoint

The shipped engine now ends each ordinary split invocation with exactly one
JSON object using `protocol: winbooksplit.result`, `version: 1`. Its status,
code, selected mode, native exit code, warnings, permitted fallback modes and
positive successful writer count are explicit. The PowerShell handler validates
that result against the actual native exit and fails closed on missing,
malformed, conflicting or zero-output success records.

For this checkpoint, existing native compatibility codes remain **0** for a
positive split, **55** for a valid document without a usable requested bookmark
plan, and **1** for other failures. The proposed final CLI exit map is a
later M3 contract; it is not implemented by this diagnostic task.

No outline and no usable Level 1 destinations permit an explicit manual retry.
Missing usable direct Level 2 bookmarks after valid Level 1 planning permits
an explicit Level 1 or manual retry. Invalid/zero-page, unreadable, malformed
outline, invalid plan and output failures offer no fallback. A decision is
captured before a retry; it never changes modes silently. Cancel keeps the
failed result and nonzero code. A failed retry remains a failure through PS/BAT.
This does not implement the later preview, noninteractive, full menu or final
parameter surface.
