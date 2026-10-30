# WinBookSplit

Split a local PDF into chapter PDFs on Windows, using existing bookmarks or
physical page starts. EPUB and AZW3 inputs are converted to PDF with separately
installed Calibre, then use the same splitter. The source file is unchanged.

This branch contains the **unpublished 30 October functional repair**: interactive
setup with consent and continuation, empty-password PDF access and narrowly
validated local opening views. The application version remains **1.0.3**; the
published v1.0.3 tag and assets contain earlier bytes and are unchanged. This
verification checkpoint publishes no release. See the
[repair evidence](https://github.com/PikkuJanne/WinBookSplit/tree/codex/functional-repair/docs/functional-repair).

The repaired engine and explicit healthy Windows launches have passed authored
PDF checks. Ordinary Store first-launch setup/resume, its repeat/network check,
Windows 10 and the private supplied book remain **NOT_RUN**. Human testing is out
of scope under the owner's latest order; no owner test, sign-off or report is
required. These limits are separate from historical v1.0.3 acceptance. Scripts
are unsigned; enforced organization controls still apply. Setup uses reviewed pypdf **6.21.3** in this
copy's own environment; no global pypdf installation is required.

The first public release,
[v1.0.0](https://github.com/PikkuJanne/WinBookSplit/releases/tag/v1.0.0), was
published on 11 October 2026. [VERSION](VERSION) is the application version
source. The older `2.1` banner was development naming, not an earlier public
semantic version.

## Requirements and formats

| Component | Compatibility scope |
| --- | --- |
| Windows target | Windows 10 **22H2 x64** and Windows 11 x64; ordinary local filesystem paths |
| PowerShell | Inbox Windows PowerShell **5.1**; **7.6.6** is the tested PowerShell 7 host |
| Python | Regular GIL Windows x64 CPython **3.15.0 or later stable 3.15 patch**, **3.14.8 or later stable 3.14 patch**, or **3.13.16 or later stable 3.13 patch** |
| pypdf | Plain **6.21.3**, controlled by its separate package version and SHA-256 lock |
| Calibre | Windows x64 **9.15.0**, needed only for EPUB/AZW3 |

The exact source-test interpreters are **3.13.16, 3.14.8 and 3.15.0**. Later stable
patches in those three branches are accepted numerically by the shipped policy; acceptance is
not a claim that each patch has been tested. Windows 11 source checks and
Windows Server 2025 CI are distinct from final-package Windows 10 acceptance.
The [v1.0.3 evidence](https://github.com/PikkuJanne/WinBookSplit/tree/codex/v1.0.3-store-discovery/docs/codex-v1.0.3/evidence)
records actual source/package checks and owner-waived manual scenarios separately.
The [public releases](https://github.com/PikkuJanne/WinBookSplit/releases) retain
their original named ZIPs, manifests and checksums; published v1.0.3 does not
include this repair. Local repair ZIPs are unpublished test artifacts bound to
their recorded source commits. A complete source-code ZIP with the entry points,
`VERSION`, requirements and `engine` files uses the same first-run path. A source
archive's format alone does not explain a setup failure, and source tests alone
do not certify an exact distribution ZIP.

The [shipped policy](engine/WinBookSplit.Compatibility.json) rejects prereleases,
free-threaded Python, PyPy, 32-bit/ARM processes and unlisted minor branches,
such as 3.16. An unknown minor is not yet validated for this release; it is
not labeled insecure. The reviewed security floors include the bundled OpenSSL fix
used by explicit HTTPS setup; earlier 3.13/3.14 patches, including 3.13.15 and
3.14.7, are excluded. Python 3.11 and 3.12 were evaluated and excluded because the
applicable Windows security fix is available only in their source releases,
without a patched official Windows binary. See [setup](docs/SETUP.md#python-policy).
Linux, macOS, network paths and arbitrary long paths remain outside the scope.
No renderer or developer test package is required to use the application.

Microsoft's standard Windows 10 support ended on **14 October 2025**.
Application compatibility does not extend operating-system servicing; see
[Microsoft's lifecycle notice](https://learn.microsoft.com/en-us/lifecycle/announcements/windows-10-end-of-support).
WinBookSplit does not require an ESU purchase or change Windows servicing settings.

| Input | Behavior and output |
| --- | --- |
| Ordinary or empty-password PDF | Copies selected page content/resources to chapter PDFs; image-only scans work with manual starts |
| EPUB | Calibre converts it to a working PDF; chapters use that PDF's physical pages and outline |
| Genuine, unencrypted AZW3/KF8 | Same conversion flow; layout and page count can differ from EPUB |

Automatic modes require actual outline destinations. There is no OCR, inferred
chapter detection, native ebook output, password interface or DRM removal.
The library attempts an empty password once. Readable empty-password PDFs are
supported; files needing a nonempty password return `password_required` (exit 7).
There is no password prompt or guessing. Unsupported encryption and unavailable
backends have separate errors; AES interoperability is not claimed. A valid
bounded local opening destination is ignored with one warning and is not copied
into chapters. Other active actions remain restricted.
Forms/XFA, signed PDFs and unsupported active document features are rejected or
handled by the narrow documented exclusions. See [PDF policy](docs/PDF_POLICY.md),
[fidelity and navigation](docs/PDF_FIDELITY.md) and [ebook conversion](docs/EBOOK_SUPPORT.md).
These checks do not make the application a document sanitizer or security sandbox.

## First launch and optional advance setup

Keep the entry points, `VERSION`, requirements and `engine` folder together in
a writable per-user folder. Use an existing accepted Python installation,
including regular x64 Python 3.15; no precisely matching patch or downgrade is
required. If Python is absent, obtain it separately from
[Python.org](https://www.python.org/downloads/windows/).

Drop a supported file onto `WinBookSplit.bat` or open the BAT and enter its path.
When a verified supported Python needs the locked PDF package, the interactive
launch offers setup once. Answer `Y` or `Yes` to use the existing application-local
setup, validate the new `.venv`, and resume the original source and options without
redropping the file. Decline, blank input, EOF or cancellation stops without
installation. An unhealthy application `.venv` requires explicit consent to
backup-preserving repair; setup failure stops the request without a retry loop.
Healthy launches do not reinstall packages.

To prepare dependencies before opening a document, run from the application
directory in PowerShell:

```powershell
.\Setup-WinBookSplit.bat -NonInteractive
```

The setup BAT resolves inbox Windows PowerShell from the Windows system
location and uses Bypass only for its child process. It works with Internet-marked
unsigned scripts where organization policy permits. It does not change persistent
policy or remove Internet marks. A complete PowerShell-host alternative is:

```powershell
& ([System.IO.Path]::Combine([Environment]::SystemDirectory, 'WindowsPowerShell\v1.0\powershell.exe')) -NoProfile -ExecutionPolicy Bypass -File '.\Setup-WinBookSplit.ps1' -NonInteractive
```

Enforced organization controls still apply; follow them if execution is denied.

Setup creates an application-local `.venv` and downloads only the hash-locked
pypdf package from PyPI. Use `-PythonPath` for a concrete trusted base interpreter,
`-Check` for read-only diagnosis, `-Wheelhouse` for the reviewed offline wheel,
and explicit `-Repair` for an unhealthy environment. See [setup](docs/SETUP.md)
for the full commands and preservation behavior. Install Calibre separately for
ebooks. Preview, NonInteractive processing and read-only `-Check` never prompt
or install. Missing or invalid required packages report `setup_required` / exit 3
with a setup action; a missing or rejected interpreter remains a runtime failure.
No launch changes global execution policy. The application and scripts are unsigned.

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
& ([System.IO.Path]::Combine([Environment]::SystemDirectory, 'WindowsPowerShell\v1.0\powershell.exe')) -NoProfile -ExecutionPolicy Bypass -File '.\WinBookSplit.ps1' -Version
```

The parameter comments at the start of `WinBookSplit.ps1` describe the
supported options; the examples below show their invocation.

Interactive PDF selection and confirmation:

```powershell
& ([System.IO.Path]::Combine([Environment]::SystemDirectory, 'WindowsPowerShell\v1.0\powershell.exe')) -NoProfile -ExecutionPolicy Bypass -File '.\WinBookSplit.ps1' -InputFile 'C:\Books\Manual.pdf' -OutputDirectory 'C:\LocalRuns'
```

Create manual chapters without prompts:

```powershell
& ([System.IO.Path]::Combine([Environment]::SystemDirectory, 'WindowsPowerShell\v1.0\powershell.exe')) -NoProfile -ExecutionPolicy Bypass -File '.\WinBookSplit.ps1' -InputFile 'C:\Books\Manual.pdf' -OutputDirectory 'C:\LocalRuns' -Mode Manual -StartPages '1,4,7' -NonInteractive
```

Inspect a Level 2 plan without publishing chapter PDFs or disk diagnostics:

```powershell
& ([System.IO.Path]::Combine([Environment]::SystemDirectory, 'WindowsPowerShell\v1.0\powershell.exe')) -NoProfile -ExecutionPolicy Bypass -File '.\WinBookSplit.ps1' -InputFile 'C:\Books\Manual.pdf' -OutputDirectory 'C:\LocalRuns' -Mode Auto -BookmarkLevel 2 -Preview -NonInteractive
```

Convert an EPUB, split Level 1 sections and retain the full converted PDF:

```powershell
& ([System.IO.Path]::Combine([Environment]::SystemDirectory, 'WindowsPowerShell\v1.0\powershell.exe')) -NoProfile -ExecutionPolicy Bypass -File '.\WinBookSplit.ps1' -InputFile 'C:\Books\Book.epub' -OutputDirectory 'C:\LocalRuns' -Mode Auto -BookmarkLevel 1 -KeepConvertedPdf -NonInteractive
```

Create AZW3 chapters from converted physical start pages:

```powershell
& ([System.IO.Path]::Combine([Environment]::SystemDirectory, 'WindowsPowerShell\v1.0\powershell.exe')) -NoProfile -ExecutionPolicy Bypass -File '.\WinBookSplit.ps1' -InputFile 'C:\Books\Book.azw3' -OutputDirectory 'C:\LocalRuns' -Mode Manual -StartPages '1,2,3' -NonInteractive
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
`.venv` fails instead of silently selecting another runtime; interactive repair
requires the explicit consent described above. `-ConversionTimeout`
is 1–86400 seconds (default 1800); `-ProcessTimeout` is 1–172800 seconds
(default the larger of 3600 and conversion timeout plus 1800). Both process
streams are bounded and drained; cancellation/deadlines preserve a nonzero result.
`-Version` needs no document or installed runtime and cannot be mixed with processing
parameters. The script's parameter comments describe every supported option.

## Outputs and troubleshooting

Execution creates a unique new run directory under Documents, or the selected
output base. Chapter names have ordered numeric prefixes and safe Windows title
labels. Existing output, neighboring files and the source are not overwritten.
Chapters are reopened and counted before publication. A successful run contains
`WinBookSplit_Manifest.json`; ebook retention adds `WinBookSplit_Converted.pdf`
separately from the chapter count. Complete PDF metadata, outlines, signatures,
annotations and cross-chapter links are not preserved generally. New chapter PDFs
are unencrypted and omit the ignored opening view; the source remains unchanged.
Encryption access is separate from permissions; this repair adds no per-operation
permission enforcement. Review the
documented fidelity limits and warnings.

Ordinary console runs use a separate `.WinBookSplit-console-<run-id>` directory
under the output base for `console.log` and, when finalized, `WinBookSplit_Run.json`.
Interactive dependency preflight can offer the single consent/setup cycle above.
Dependency failures show a brief reason and one setup action; the local log
retains technical details. If diagnostics cannot be reserved safely, the failure
frame stays on the console with the original nonzero status. Preview and early
NonInteractive dependency failures create no disk diagnostics and retain the
machine-readable console frames. Local records can contain paths, titles,
messages and source identities. Keep them private; use the
[explicit redacted support export](docs/SUPPORT_DIAGNOSTICS.md) when reporting
a problem. No upload or export is automatic.

| Application exit | Meaning |
| --- | --- |
| 0 | Successful execution, validated preview or version query |
| 2 | Invalid input, arguments or path |
| 3 | Runtime/package prerequisite, setup required or missing encryption backend |
| 4 | Ebook conversion failure |
| 5 | No usable plan at the requested bookmark level |
| 6 | PDF/output or setup/rollback failure, including incomplete diagnostic finalization |
| 7 | Unsupported document/encryption or nonempty password required |
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
add telemetry or install packages without explicit setup authorization or
interactive consent. Ebook resources and Calibre plugins are
not isolated by an OS network sandbox; prefer self-contained ebooks. See
[security reporting](SECURITY.md), [contributing](CONTRIBUTING.md),
[changes](CHANGELOG.md) and [third-party notices](THIRD_PARTY_NOTICES.md).
WinBookSplit is provided under the [MIT license](LICENSE), without warranty.
