# PowerShell/Python boundary and dependency reliability

## Minimal extraction

Move the existing Python implementation into `engine/winbooksplit_engine.py` with a narrow callable test seam and guarded CLI. Keep entry points and behavior characterization first; fix bugs in subsequent targeted commits. PowerShell resolves the engine relative to `$PSScriptRoot`, never the current directory or a shared temporary filename. Avoid adding a web framework, job server, heavy GUI, new PDF engine, or generic plugin architecture.

## Runtime discovery

A version command must work without document/dependency discovery. Before any conversion/writes, find a usable Python interpreter and successfully import the pinned/tested pypdf from that interpreter. Prefer an explicit `-PythonPath`, the documented application-local venv, then validated launcher/PATH candidates. Never run an arbitrary `python.exe` next to an untrusted book. Use `python -m pip` with that exact interpreter in setup instructions, not ambiguous `pip`. Offline processing must not auto-install or alter global packages.

Choose/test minimum and selected Python versions during M0; helper Python 3.10+ is not automatically the application support matrix. Keep runtime and developer dependency files separate. Pin a reviewed working dependency set; document how it was chosen and how to update it safely. Check current package/runtime advisories before release without labeling an unperformed scan 'passed'. User installation/network access is an explicit setup action. Source [S3] documents current pypdf installation requirements; do not assume 'all Python 3'.

M0-T03 freezes regular x64 CPython 3.14.8 only (minimum and selected are the same), plain pypdf 6.19.0 and the Calibre 9.15.0 conversion test target. See `../SUPPORT_AND_SETUP.md` and the root hashed runtime/dev requirements. These isolated dependency checks do not prove the original launcher's discovery or later Windows/conversion acceptance; M2-T04 must integrate and enforce the recorded selection.

## Process supervision

Start child executables directly (`UseShellExecute = false` where appropriate). Correctly marshal arguments for Windows PowerShell 5.1/.NET Framework **and** PowerShell 7; do not blindly use APIs available only in newer .NET. Prefer a reviewed, small compatibility helper over repeated quoted strings. Untrusted titles belong in a local UTF-8 JSON plan, not an executable command. Never use `Invoke-Expression`, `cmd /c` with constructed text, or text interpolation into Python code.

Drain stdout and stderr concurrently from process start to EOF, including fast exit, blank lines, output without terminal newline, large bursts, and final tail. Bound buffers/log sizes. If using .NET callbacks from PowerShell, avoid callbacks that execute arbitrary PowerShell without an available runspace; test the actual host. Do not repair the current deadlock by reading one full pipe then the other. Source [S5] explains redirected-stream dependency hazards.

Use an explicit UTF-8 child environment/stdout/stderr configuration and matching decoding. Handle diagnostics separately from machine-readable results. A small result JSON file or an explicit framed protocol is acceptable; don't scrape colored human text to infer success. Close readers/processes with try/finally. Clear, distinct errors for startup failure, no bookmarks, invalid input, conversion, processing/output validation, timeout/cancel. Document a test timeout and conservative configurable conversion timeout; avoid arbitrary short limits for large books. Terminate only the run's tracked process tree and only after cancellation/timeout, not `taskkill /IM python.exe` or Calibre-wide killing.

PowerShell captures the final attempted result, including fallback. The .bat launcher preserves it with `exit /b` and handles no/multiple inputs explicitly. Do not overwrite it with `pause`, logging, or cleanup exit status. Zero files on an execution request is never success. Preview success means 'plan only', not 'files created'.

## Paths and permissions

Use literal filesystem paths and file-only checks. Resolve a full provider path, ensure correct input extension and actual file readability, then let the parser verify contents. Reject directories with .pdf extensions. Test brackets, ampersands, apostrophes, parentheses, percent/exclamation marks, spaces, unicode, long paths, read-only sources, full/unwritable destinations, and paths from different current directories. Do not promise arbitrary UNC/long-path support not demonstrated in Windows tests.

No elevation, no global execution-policy change, no disabling security tools. A process-scoped launcher policy option may preserve existing usability but is not a trust guarantee and must respect enterprise controls. Record its behavior; support rejection rather than advising a user to bypass organization policy.

## M2-T02 implementation checkpoint

The current PS entrypoint uses `Resolve-Path -LiteralPath`, FileSystem provider and
FileInfo/readability checks, then literal metadata; directories/provider objects
fail before processing with exit 1. `-OutputDirectory` selects an existing literal
base; default remains actual Documents. Observed reparse ancestors reject. Native
exclusive creation reserves console records, and a complete setup catch preserves
errors/nonzero exit and names any retained directory. Both actual supported hosts
and BAT preserve the tested literal/read-only source. Controlled PATH/stdin/UTF-8
wrappers are recorded; no Explorer or ordinary discovery/all-entrypoint claim.
No/multiple-input and final noninteractive/CLI behavior still belongs to M3;
conversion discovery/streams remain scheduled M2 work. See
`../evidence/M2-T02-paths.md` and `../../../tests/paths/README.md`.
