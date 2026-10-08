# Frozen support and dependency plan

Decision date: 8 October 2026 (Europe/Berlin), M0-T03, AC-007/AC-008.
This is the v1.0.0 implementation/test target. The original application still
contains the reproduced defects; installing these dependencies does not fix it.
The historical README and source headers are unchanged until the documentation
task. Release support requires the later Windows, conversion and exact-ZIP gates.

## Matrix and evidence boundary

| Component | Frozen target | Actual M0-T03 evidence |
| --- | --- | --- |
| OS | Windows 11 x64, ordinary local filesystem paths | Current workstation: Windows 11 Pro 26H2, build 26300.9457. This is an isolated workstation setup, not a clean OS installation or a test of every Windows 11 build. |
| Shell | Windows PowerShell 5.1 and PowerShell 7.6.5 x64 | Probed 5.1.26100.9444 (.NET Framework CLR 4.0.30319.42000) and 7.6.5. Only the original 5.1 early rejection path and batch no-input path ran; successful splitting under both hosts remains required. |
| Application Python | Regular GIL CPython **3.14.8 x64 only** | Fresh official runtime extracted without registration; isolated runtime/dev venvs tested. Minimum and selected version are both 3.14.8. No wider minor/patch range is claimed. |
| PDF engine | Plain **pypdf 6.19.0** | Hash-verified install/import, dependency check and unchanged embedded-engine characterization on 3.14.8. No optional extras. |
| Ebook converter | Windows x64 **Calibre 9.15.0**, only needed for EPUB/AZW3 | Absent from PATH and all three original discovery locations, plus the inspected D: location. Version is selected from current primary sources; NOT installed or conversion-tested. |
| Python dev packages | **ReportLab 5.0.1**, **Pillow 12.3.0**, **charset-normalizer 3.5.2** | Fresh hash-verified install, pip check, deterministic fixture regeneration and original characterization. These are absent from the runtime-only venv. |
| Python test runner | CPython 3.14.8 standard-library **unittest** | Chosen for M0-T04 scaffolding; no pytest dependency is needed. No new unit suite is claimed here. |
| Shell dev tools | **Pester 6.2.0**, **PSScriptAnalyzer 1.25.0** | Selected from official releases/Gallery; NOT installed/imported/tested. Both hosts currently expose inbox Pester 3.4.0 only, with no analyzer. M0-T04 must use explicit versioned manifests. |
| Render QA tool | **Poppler 26.10.0** target for later fidelity tests | Existing 26.07.0 was probed; its M0-T02 render evidence is historical. Newer 26.10.0 and Windows binary provenance are NOT tested here; re-review before M4 rendering. No renderer is a runtime dependency. |
| Setup tooling | Python install manager **26.3** and venv bootstrap **pip 26.2.1** observed | Manager's `--target` extraction left the registered 3.14.7 installation unchanged. No global package installation or policy change. These are setup tools, not shipped runtime dependencies. |

Windows 10, ARM, 32-bit Python, free-threaded Python, other Python versions,
Linux/macOS application execution, live UNC paths and arbitrary long paths are
unverified and outside the v1.0.0 support claim. Python/pypdf/Calibre upstream
platform minima do not broaden this matrix. Windows registry `ProductName`
still says Windows 10 Pro on this workstation; CIM caption and Python platform
identify Windows 11, so that stale registry string is not a Windows 10 test.

## Features and release policy

PDF output only. Ordinary unencrypted PDFs, including image-only scanned pages,
use manual starts or existing bookmarks; no OCR or inferred chapters. EPUB and
AZW3 to PDF are mandatory release features using real Calibre inputs and output
inspection. Calibre is optional for PDF-only users, not an optional release gate.
Preserve the existing conversion profile first; select verified settings in
M2-T03/M4-T04. No ebook-native output or DRM removal.

Detect and reject every encrypted PDF before extraction or output creation,
including empty-open/owner-password files; do not decrypt to test the policy.
No crypto extra or password interface. Forms/XFA, signatures and unsupported
active catalog features follow the explicit rejection/warning policy in
`specs/PDF_SUPPORT.md`; these protections are requirements, not implemented
passes at M0. Keep source files immutable and all processing local.

The release may be unsigned. Release notes must say so and include SHA-256
checksums; no signing purchase or SmartScreen guarantee. Keep project MIT.
Python, Calibre and developer binaries are installed separately, not bundled.
If any dependency bytes are later redistributed, review their actual licenses
and notices in M5 before packaging; requirements files alone do not redistribute
the packages.

## Isolated setup (explicit user/developer action)

Obtain regular x64 Python 3.14.8 from [its official release page][S10]. Use an
explicit interpreter path. With the already installed Python manager, a local
extraction is possible without changing registered runtimes:

```powershell
# Choose a new directory; do not extract over an existing runtime.
$PythonDirectory = 'C:\Tools\WinBookSplit-Python-3.14.8'
if (Test-Path -LiteralPath $PythonDirectory) { throw 'Choose a new runtime directory' }
py install --target=$PythonDirectory 3.14.8
if ($LASTEXITCODE -ne 0) { throw 'Python extraction failed' }
$SetupPython = Join-Path $PythonDirectory 'python.exe'
```

An existing official 3.14.8 installation can instead supply `$SetupPython`.
Run the following from either supported shell after setting that absolute path
and the trusted application directory. Refuse to reuse an existing venv so an
old environment cannot silently supply undeclared packages. No activation or
execution-policy change is needed.

```powershell
$ApplicationRoot = (Resolve-Path -LiteralPath 'C:\Tools\WinBookSplit').ProviderPath
& $SetupPython -I -c "import sys,platform,sysconfig; assert sys.version_info[:3] == (3,14,8); assert sys.platform == 'win32' and platform.machine() == 'AMD64'; assert not sysconfig.get_config_var('Py_GIL_DISABLED')"
if ($LASTEXITCODE -ne 0) { throw 'Expected regular Windows x64 CPython 3.14.8' }
$RuntimeVenv = Join-Path $ApplicationRoot '.venv'
if (Test-Path -LiteralPath $RuntimeVenv) { throw 'Choose a fresh application/venv directory' }
& $SetupPython -I -m venv $RuntimeVenv
if ($LASTEXITCODE -ne 0) { throw 'venv creation failed' }
$RuntimePython = Join-Path $RuntimeVenv 'Scripts\python.exe'
& $RuntimePython -I -m pip install --require-hashes --only-binary=:all: --index-url https://pypi.org/simple -r (Join-Path $ApplicationRoot 'requirements.txt')
if ($LASTEXITCODE -ne 0) { throw 'Runtime dependency installation failed' }
& $RuntimePython -I -m pip check
if ($LASTEXITCODE -ne 0) { throw 'Runtime dependency check failed' }
& $RuntimePython -I -c "import sys,pypdf; assert sys.prefix != sys.base_prefix; assert pypdf.__version__ == '6.19.0'; print(sys.executable); print(pypdf.__file__)"
if ($LASTEXITCODE -ne 0) { throw 'Runtime import check failed' }
```

Inspect the printed paths: pypdf must come from this venv, never from a book
directory or unrelated interpreter. Network access occurs only for explicit
setup. A reviewed wheelhouse can be prepared with the same requirements and
`pip download --require-hashes --only-binary=:all:` and then installed using
`--no-index --find-links <absolute-wheelhouse>` in place of `--index-url`.
Normal document processing must remain offline and never auto-install.

M2-T04 will implement explicit `-PythonPath`, application-local `.venv`, and
validated launcher/PATH discovery in that order. The current launcher has no
`-PythonPath` option and does not automatically select this venv. These setup
commands prepare dependencies; they are not a claim that original launcher
discovery or splitting is ready.

For development, create a separate fresh `.venv-dev` with the same interpreter,
then use its absolute `Scripts\python.exe` with the same pip flags and
`requirements-dev.txt`, followed by `pip check`. That file includes the runtime
pin and all mandatory developer transitive packages. Its native wheel hashes
target regular CPython 3.14 Windows x64; no cross-platform install is promised.
Use `-I` for direct scripts and explicit import origins; M0-T04's test runner
must deliberately arrange its trusted test imports without a book directory.

For the shell test tools, M0-T04 should `Save-Module -Name Pester
-RequiredVersion 6.2.0 -Path <new-tool-directory>` and similarly save
PSScriptAnalyzer 1.25.0, then import each absolute versioned `.psd1` in fresh
`powershell -NoProfile` and `pwsh -NoProfile` children. Check command errors and
the imported version; do not fall back to inbox Pester. This is an unexecuted
setup plan, not permission to change global module locations.

Calibre 9.15.0 must be obtained separately from its official Windows/portable
distribution and `ebook-convert.exe --version` checked at a trusted absolute
path. M2-T04 will support a trusted nonstandard converter path and resolve it
only for ebooks. It must not find an executable next to an untrusted input.
No Calibre installer, converter or library is shipped with WinBookSplit.

## Pin selection and maintenance

The checked primary sources are indexed in `SOURCES.md` (S10-S18). Python
3.14.8 is the 30 September expedited security release. pypdf 6.19.0 includes
published fixes newer than historical 6.10.0; plain pypdf on 3.14.8 has no
mandatory transitive package. ReportLab 4.4.9 is historical only: vendor notes
identify a later `rl_safe_eval` security fix and 5.0's safer image-fetch defaults.
The synthetic generator does not fetch URLs or evaluate untrusted layouts.

On 8 October, the public pypdf/Calibre advisory APIs returned 54/21 records with
no next page. A manual published-range review found no declared affected range
including 6.19.0/9.15.0; 31 pypdf ranges included historical 6.10.0. Pillow's
12.3.0 release fixes the inspected decompression and Windows-viewer advisories.
Vendor notes were also checked because empty PyPI vulnerability arrays alone
missed ReportLab's known fix. This is not an automated scanner pass, a full
third-party audit, or proof that unpublished defects are absent.

Before changing any pin, consult current vendor requirements/release notes and
advisories, resolve only the necessary wheels in a fresh isolated environment,
verify their PyPI hashes, update both requirements/evidence as applicable, and
rerun affected characterization/regressions. Expanding the Python/OS matrix
requires actual tests and a recorded decision. M5-T04 and M6 must recheck
advisories and support evidence before the release, without waiving real
Calibre, Explorer, both-shell or exact-asset checks.

[S10]: https://www.python.org/downloads/release/python-3148/
