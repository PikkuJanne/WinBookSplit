# PowerShell/Python boundary and dependency reliability

## Minimal extraction

Move the existing Python implementation into `engine/winbooksplit_engine.py` with a narrow callable test seam and guarded CLI. Keep entry points and behavior characterization first; fix bugs in subsequent targeted commits. PowerShell resolves the engine relative to `$PSScriptRoot`, never the current directory or a shared temporary filename. Avoid adding a web framework, job server, heavy GUI, new PDF engine, or generic plugin architecture.

## Runtime discovery

A version command must work without document/dependency discovery. Before any conversion/writes, find a usable Python interpreter and successfully import the pinned/tested pypdf from that interpreter. Prefer an explicit `-PythonPath`, the documented application-local venv, then validated launcher/PATH candidates. Never run an arbitrary `python.exe` next to an untrusted book. Use `python -m pip` with that exact interpreter in setup instructions, not ambiguous `pip`. Offline processing must not auto-install or alter global packages.

Choose/test minimum and selected Python versions during M0; helper Python 3.10+ is not automatically the application support matrix. Keep runtime and developer dependency files separate. Pin a reviewed working dependency set; document how it was chosen and how to update it safely. Check current package/runtime advisories before release without labeling an unperformed scan 'passed'. User installation/network access is an explicit setup action. Source [S3] documents current pypdf installation requirements; do not assume 'all Python 3'.

M0-T03 freezes regular x64 CPython 3.14.8 only (minimum and selected are the same), plain pypdf 6.19.0 and Calibre 9.15.0. See `../SUPPORT_AND_SETUP.md` and the root hashed runtime/dev requirements. Its isolated setup observations are historical; the M2-T04 checkpoint below implements and tests selection in actual launchers.

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

## M2-T04 implementation checkpoint

`engine/WinBookSplit.Runtime.ps1` exposes `Resolve-WinBookSplitRuntime` with
ApplicationRoot, PythonPath, DocumentPath and TimeoutSeconds (default 10, range
0.05..60). It returns selected Path/Version/Source, PypdfVersion/PypdfPath,
strict probe Details, native Probe, Arguments and Attempts. The PS entrypoint
appends PythonPath after earlier parameters and runs preflight before console
reservation, conversion or writing. The same absolute interpreter executes the
existing engine with `-I -B`. Standalone version/CLI changes remain M3.

Precedence is explicit PythonPath, application `.venv\Scripts\python.exe`, ordinary
validated `py.exe -0p` read-only listing, then safe PATH python.exe. Invalid explicit
selection or an existing broken/reparse .venv fails without substitution. Empty/
relative PATH entries and book/unrelated-CWD descendants are excluded. Ordinary
absolute long paths and ancestors are checked; reparse/short aliases reject, and
automatic executable candidates require one hard link. Explicit trusted paths
may have hard links but still reject short/reparse paths. Canonical checks do not
authenticate executables or sandbox another process running as the same user.

The fixed Python probe uses `-I -B -c` and strictly validates one object frame:
regular GIL Windows AMD64 64-bit CPython 3.14.8, pypdf 6.19.0 distribution metadata,
site-packages import/all loaded pypdf module origins, isolation and no bytecode.
No input is interpolated into code. Launcher listing disables automatic installs;
processing never installs/registers runtimes. Guidance uses the exact interpreter's
`-I -m pip` and hashed requirements. Selected paths/versions/imports are printed
and logged; structured `[DEPENDENCY-ERROR]` attempts preserve rejection reasons.
`[ENGINE]` is launch intent; actual result frames and reopened PDFs prove execution.

`Resolve-WinBookSplitConverter` takes ApplicationRoot, CalibrePath, DocumentPath
and the same probe timeout. Only ebooks call it: explicit converter, safe PATH,
then known locations. It validates native zero and the exact 9.15.0 banner with
the optional exact real creator line. PDF skips even an invalid CalibrePath.
Real portable explicit/PATH selection and unchanged BAT PATH processing pass.

The narrow native compatibility probe launches without a shell, uses literal
Windows argv, explicit UTF-8, concurrent background byte drains and 64 KiB tails
per stream with total/truncation observations. Deadline covers parent exit and
pipe EOF. Timeout kills/waits only the exact launched parent with bounded grace;
DescendantsStopped is null, never an invented proof. Inherited-pipe reader threads
may remain until EOF; no filesystem cleanup is authorized from that probe. The
existing owned-job Calibre conversion remains separate. Remaining engine bounds,
deadline/tree/cancel/UTF-8 work and cumulative M2 review belong to M2-T05.

See `../evidence/M2-T04-runtime.md/json` and `../../../tests/runtime/README.md`.
Actual host/process, complete physical-page/content, source/neighbor/decoy,
manifest/policy/known cleanup acceptance passes; broader compatibility, human
Explorer, CI, package and release gates remain open.

## M2-T05 implementation checkpoint

`WinBookSplit.Process.ps1` supplies the narrow Windows native transport for both
engine attempts and dependency probes. Existing literal argument quoting feeds
direct `CreateProcessW` with restricted inherited standard handles. The child
starts suspended, joins this run's non-breakaway kill-on-close job, then resumes.
Background .NET readers concurrently drain both streams to EOF without PowerShell
callbacks. Strict UTF-8 validation, 64 KiB rings, actual byte totals and truncation
preserve bounded diagnostics; a separate 8 MiB result budget preserves framed
JSON independently of human tails and refuses oversize/ambiguous/missing results.
The fixed Python probe and engine use `-I -B -X utf8`; isolation ignores Python
encoding environment hooks, so the explicit interpreter flag matters.

`-ProcessTimeout` follows existing parameters, accepts 1..172800 seconds and
defaults to max(3600, ConversionTimeout + 1800). Native test seam deadlines accept
0.05 seconds. Deadlines cover parent, job membership and both EOFs. Cancellation
tokens and the native Console.CancelKeyPress handler terminate only owned job
handles, with bounded stop/drain grace and explicit proof/errors. Timeout or
cancellation preserves failure through the existing launcher. The console lacks
the engine file ledger and retains interrupted stages; process shutdown never
authorizes broad deletion. Human Ctrl+C behavior remains a separate actual gate.

Preflight retains its exact selection/compatibility/banners and 10-second default,
0.05..60 range, but now shares this transport and requires stopped descendants.
The earlier T04 parent-only lifecycle above is historical. Source parsing still
runs with the user's privileges; this is process ownership, not a document sandbox.

## M3-T01 outcome checkpoint

Dependency cancellation/timeout now stops discovery immediately with 130 and
the actual native shutdown proof; another runtime/converter is never substituted
after interruption. Engine transport failures, validated engine results and
explicit fallback cancellation use the shipped shared outcome map. The console
captures each attempt before logging it, preserving the final attempt and primary
failure even if a log write fails. Final console success is emitted only after
flush/close; the preceding log `[OPERATION-OUTCOME]` records the processing stage,
while stdout `[OUTCOME]` is authoritative for finalization. The BAT launcher saves
and explicitly returns the child status. No process ownership or filesystem
cleanup policy is broadened; outer interruption still retains marked engine
staging when it lacks the engine ledger.

The application sets its own console output encoding to UTF-8 before writing
headers or machine outcomes, including through BAT. A redirected PowerShell 7
console skips cosmetic screen clearing while retaining the printed header.
Neither change writes machine-wide settings.

An actual whole-script OS Ctrl+C reproduction showed that the earlier managed
Console.CancelKeyPress callback could stop the PowerShell script pipeline before
it emitted an outcome. During supervised processing, a rooted native console
handler now consumes only Ctrl+C/Break, records the cancellation flag, and leaves
the owned job shutdown/drains to the ordinary supervisor. Registration/removal
failures are explicit; the handler remains alive through shutdown and is removed
afterward. This affects only the calling process during supervision. Windows
documents the last-registered handler order and TRUE return semantics in
[SetConsoleCtrlHandler](https://learn.microsoft.com/en-us/windows/console/setconsolectrlhandler).
Automated signals in hidden private consoles are separate from human keypress,
menu, Explorer, close-window and power-loss checks.
