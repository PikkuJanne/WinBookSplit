# Frozen support and dependency plan

**Historical planning and evidence record.** For current end-user setup and the
implemented support scope, use [docs/SETUP.md](../SETUP.md) and
[README](../../README.md). The dated observations and old example recipes below
are retained as history; they are not new acceptance results or the current
end-user command inventory.

Decision date: 8 October 2026 (Europe/Berlin), M0-T03, AC-007/AC-008.
This is the v1.0.0 implementation/test target. At M0 the original application
still contained the reproduced defects; dependency installation alone was no fix.
M1-T01 updated the PowerShell header for the shipped engine. M2-T04 now updates
README setup/runtime guidance and implements preflight as recorded below; full
documentation and release support still require the later Windows and exact-ZIP
gates. Historical matrix observations are preserved explicitly.

M0-T04 subsequently verified the exact isolated Pester/PSScriptAnalyzer pins
under both actual hosts through the local scaffold; see
[`evidence/M0-T04-scaffolding.md`](evidence/M0-T04-scaffolding.md) and
`tests/powershell/README.md`. The table's last column retains M0-T03's historical
observations. M0-T04 does not certify application success paths or conversion.

Subsequent M1-T01 observation (8 October 2026): the user requested Calibre
installation. The official portable **9.15.0** build was installed per-user at
`$UserProfile\Apps\Calibre915\Calibre Portable`, without elevation or global
PATH/policy changes. Its download matched the official SHA-512 and GitHub
asset SHA-256, with a valid Kovid Goyal Windows signature; the installed
`Calibre\ebook-convert.exe --version` returned 9.15.0/exit 0. See
[`evidence/M1-T01-extraction.md`](evidence/M1-T01-extraction.md) and its machine
record. This trusted nonstandard path is outside the unchanged launcher's three
search locations. Use the recorded absolute converter path for later tests;
application discovery work remains M2-T04. No actual EPUB/AZW3 conversion or
GUI interaction is claimed. The matrix below retains its historical M0-T03
observations rather than retroactively changing them.

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

## Historical isolated-setup recipe (explicit user/developer action)

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
directory or unrelated interpreter. Dependency downloads occur during explicit
setup. A reviewed wheelhouse can be prepared with the same requirements and
`pip download --require-hashes --only-binary=:all:` and then installed using
`--no-index --find-links <absolute-wheelhouse>` in place of `--index-url`.
Normal document processing never auto-installs dependencies or uploads documents
through WinBookSplit. This is not OS network isolation for Calibre/plugins or
external ebook resources; prefer self-contained ebooks. The later bounded
loopback observation is documented in [ebook support](../EBOOK_SUPPORT.md).

M2-T04 implements explicit `-PythonPath`, application-local `.venv`, validated
read-only `py.exe -0p` listing, then trusted PATH discovery. The chosen interpreter
must pass exact regular Windows x64 CPython 3.14.8/pypdf 6.19.0 isolated import
checks before console/conversion/output creation, and executes the engine with
`-I -B`. Invalid explicit or existing broken .venv selections fail without
substitution. Book/current-directory modules and automatic executable decoys are
excluded. Printed/logged selected paths and versions identify the setup target.
See `evidence/M2-T04-runtime.md/json` for actual host acceptance and limits.
Setup commands remain explicit user actions; normal processing never installs.

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
path. M2-T04 supports trusted nonstandard `-CalibrePath`, then safe PATH and known
locations, and resolves/version-checks it only for ebooks. Real portable conversion
passes under both actual hosts and unchanged BAT. PDF skips Calibre entirely.
Automatic discovery excludes book/unrelated-CWD descendants, empty/relative PATH,
reparse/short aliases and hard-linked executables. An explicit trusted path may
have hard links but must remain an ordinary long path. These compatibility checks
are not authentication or a document sandbox; no wider support matrix is implied.
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
