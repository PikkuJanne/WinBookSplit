# WinBookSplit

Split a local PDF into chapter PDFs on Windows, using existing bookmarks or
physical page starts. EPUB and AZW3 inputs are converted to PDF with separately
installed Calibre, then use the same splitter. The source file is unchanged.

This is the **1.0.0 release candidate**, not a published release or a tested
release ZIP. [VERSION](VERSION) is the application version source. The older
`2.1` banner was development naming, not an earlier public semantic version.

## Requirements and formats

| Component | Supported and tested scope |
| --- | --- |
| Windows | Windows 11 x64 with ordinary local filesystem paths |
| PowerShell | Windows PowerShell 5.1; PowerShell 7.6.6 on the local Windows 11 setup |
| Python | Regular GIL CPython **3.14.8 x64**, with plain **pypdf 6.20.0** |
| Calibre | Windows x64 **9.15.0**, needed only for EPUB/AZW3 |

Automated CI also exercises Windows Server 2025 with PowerShell 5.1 and 7.6.6;
that is separate from the Windows 11 document and launcher checks. Windows 10,
ARM, 32-bit/free-threaded Python, other Python/pypdf/Calibre versions, Linux,
macOS, network paths and arbitrary long paths are outside the support claim.
No renderer or developer test package is required to use the application.

| Input | Behavior and output |
| --- | --- |
| Ordinary unencrypted PDF | Copies selected page content/resources to chapter PDFs; image-only scans work with manual starts |
| EPUB | Calibre converts it to a working PDF; chapters use that PDF's physical pages and outline |
| Genuine, unencrypted AZW3/KF8 | Same conversion flow; layout and page count can differ from EPUB |

Automatic modes require actual outline destinations. There is no OCR, inferred
chapter detection, native ebook output, password interface or DRM removal.
Encrypted PDFs are rejected even when an empty password would open them.
Forms/XFA, signed PDFs and unsupported active document features are rejected or
handled by the narrow documented exclusions. See [PDF policy](docs/PDF_POLICY.md),
[fidelity and navigation](docs/PDF_FIDELITY.md) and [ebook conversion](docs/EBOOK_SUPPORT.md).
These checks do not make the application a document sanitizer or security sandbox.

## Set up once

Keep the entry points and `engine` folder together. Follow [setup](docs/SETUP.md)
to create a fresh application `.venv` using the exact Python/pypdf pins. Install
Calibre separately for ebooks. Processing never installs dependencies or changes
global execution policy. The application and scripts are unsigned; setup explains
checksum verification and what it can establish.

## Use the console

Drop **one** supported file onto `WinBookSplit.bat`, or double-click the BAT and
enter a literal source path. The BAT uses Windows PowerShell 5.1. Multiple dropped
files are rejected. The PowerShell entry point can also run in either tested shell.

Choose `1` for Level 1 bookmarks, `2` for Level 2 bookmarks or `M` for manual
starts. `C` or closed input cancels. Invalid or blank method choices are explained
and requested again; blank source-path or plan-confirmation input cancels. If an
automatic mode has no usable plan, explicitly choose another mode or cancel.

Manual starts are **one-based physical PDF page numbers**, such as `1,4,7`;
they are not printed page labels or page ranges. Every comma-separated token must
be an ASCII decimal integer in the document. Starts are sorted and deduplicated
with a notice; page 1 is included. Every physical page is covered exactly once.
Level 2 sections stay inside their Level 1 parent; front matter and parent opening
pages are retained.

For an ebook, conversion finishes before manual page entry. You may request the
working PDF with `O`, continue with `N`, or cancel with `C`. Use its physical pages,
including any generated table-of-contents page. Review the complete plan, then
type `Y` to create chapters or `C` to cancel. After success, `O` requests opening
the output folder and `N` finishes. Opening uses the current Windows association;
WinBookSplit does not supply a PDF viewer.

## Scripted examples

Run these examples from the application directory in PowerShell. Replace the
example absolute paths with your own ordinary local files and an **already-existing**
output base (`C:\LocalRuns`). `Manual.pdf` needs at least seven pages and usable
Level 2 bookmarks for the preview example; the EPUB needs usable Level 1 bookmarks.
The AZW3 example needs at least three converted physical pages.

```powershell
.\WinBookSplit.ps1 -Version
```

```powershell
Get-Help .\WinBookSplit.ps1 -Full
```

Interactive PDF selection and confirmation:

```powershell
.\WinBookSplit.ps1 -InputFile 'C:\Books\Manual.pdf' -OutputDirectory 'C:\LocalRuns'
```

Create manual chapters without prompts:

```powershell
.\WinBookSplit.ps1 -InputFile 'C:\Books\Manual.pdf' -OutputDirectory 'C:\LocalRuns' -Mode Manual -StartPages '1,4,7' -NonInteractive
```

Inspect a Level 2 plan without publishing chapter PDFs or disk diagnostics:

```powershell
.\WinBookSplit.ps1 -InputFile 'C:\Books\Manual.pdf' -OutputDirectory 'C:\LocalRuns' -Mode Auto -BookmarkLevel 2 -Preview -NonInteractive
```

Convert an EPUB, split Level 1 sections and retain the full converted PDF:

```powershell
.\WinBookSplit.ps1 -InputFile 'C:\Books\Book.epub' -OutputDirectory 'C:\LocalRuns' -Mode Auto -BookmarkLevel 1 -KeepConvertedPdf -NonInteractive
```

Create AZW3 chapters from converted physical start pages:

```powershell
.\WinBookSplit.ps1 -InputFile 'C:\Books\Book.azw3' -OutputDirectory 'C:\LocalRuns' -Mode Manual -StartPages '1,2,3' -NonInteractive
```

`-Mode Auto` requires `-BookmarkLevel 1` or `2`; `-Mode Manual` requires
`-StartPages`. `-NonInteractive` requires complete choices and never prompts,
pauses, opens files/folders, clears the console or retries another mode.
`-NoPause` suppresses only the final pause; interactive decisions still apply.
`-Preview` also requires complete choices. Ebook preview performs temporary
conversion and cleanup; it cannot be combined with `-KeepConvertedPdf`.

Use `-OutputDirectory` for an **existing output base**, `-PythonPath` for an
explicit trusted absolute Python executable, or `-CalibrePath` for the trusted
absolute `ebook-convert.exe`. An invalid explicit path or broken application
`.venv` fails instead of silently selecting another runtime. `-ConversionTimeout`
is 1–86400 seconds (default 1800); `-ProcessTimeout` is 1–172800 seconds
(default the larger of 3600 and conversion timeout plus 1800). Both process
streams are bounded and drained; cancellation/deadlines preserve a nonzero result.
`-Version` needs no document or installed runtime and cannot be mixed with processing
parameters. Full help lists every parameter.

## Outputs and troubleshooting

Execution creates a unique new run directory under Documents, or the selected
output base. Chapter names have ordered numeric prefixes and safe Windows title
labels. Existing output, neighboring files and the source are not overwritten.
Chapters are reopened and counted before publication. A successful run contains
`WinBookSplit_Manifest.json`; ebook retention adds `WinBookSplit_Converted.pdf`
separately from the chapter count. Complete PDF metadata, outlines, signatures,
annotations and cross-chapter links are not preserved generally; review the
documented fidelity limits and warnings.

After setup, a separate `.WinBookSplit-console-<run-id>` directory under the
output base holds `console.log` and, when finalized, `WinBookSplit_Run.json`.
Preview and early validation/dependency failures create no disk diagnostics.
Local records can contain paths, titles, messages and source identities. Keep
them private; use the [explicit redacted support export](docs/SUPPORT_DIAGNOSTICS.md)
when reporting a problem. No upload or export is automatic.

| Application exit | Meaning |
| --- | --- |
| 0 | Successful execution, validated preview or version query |
| 2 | Invalid input, arguments or path |
| 3 | Dependency setup failure |
| 4 | Ebook conversion failure |
| 5 | No usable plan at the requested bookmark level |
| 6 | PDF/output failure, including incomplete diagnostic finalization |
| 7 | Unsupported document feature |
| 130 | Cancellation or deadline |

PowerShell errors before the script runs use the host's own status. Missing or
malformed `VERSION` returns 6 before input or dependency processing; restore the
complete trusted application rather than inventing a replacement version. A success
followed by failed diagnostic finalization returns 6, retains completed chapters
and omits `Done.` Inspect the outcome rather than assuming any files mean success.
For a no-bookmark PDF, use valid manual starts. For unreadable or unsupported input,
retain the original and inspect the stated reason. For missing dependencies, use
[setup](docs/SETUP.md); do not disable policy, install globally or accept an
untrusted runtime just to suppress an error.

WinBookSplit runs with your account's privileges. It does not upload documents,
add telemetry or auto-install packages. Ebook resources and Calibre plugins are
not isolated by an OS network sandbox; prefer self-contained ebooks. See
[security reporting](SECURITY.md), [contributing](CONTRIBUTING.md),
[changes](CHANGELOG.md) and [third-party notices](THIRD_PARTY_NOTICES.md).
WinBookSplit is provided under the [MIT license](LICENSE), without warranty.
