# Set up WinBookSplit

The application targets Windows 11 x64, Windows PowerShell 5.1 and the tested
PowerShell 7.6.5 host. Use regular GIL CPython **3.14.8 x64** and plain
**pypdf 6.19.0**. EPUB/AZW3 inputs also require Windows x64 **Calibre 9.15.0**.
Other versions and platforms are not part of this support claim. Developer
renderers, Pester and fixture packages are not application requirements.

Obtain Python from [the official 3.14.8 release](https://www.python.org/downloads/release/python-3148/)
and Calibre, if needed, from [its official download archive](https://download.calibre-ebook.com/9.15.0/).
Choose trusted local installations. WinBookSplit includes neither runtime nor
converter binaries. Keep its entry points, `VERSION`, requirements and `engine`
folder together. This source/candidate setup is not certification of a future
release ZIP or a clean operating system.

## Create a fresh application environment

Replace the two example absolute paths below with the trusted application folder
and an existing official Python 3.14.8 x64 interpreter. Run the entire block in
either supported PowerShell host. It refuses an existing `.venv`; use a fresh
application folder instead of deleting or repairing someone else's environment.
No activation, administrator session or execution-policy change is required.
The package download is an explicit setup action.

```powershell
$ApplicationRoot = (Resolve-Path -LiteralPath 'C:\Tools\WinBookSplit').ProviderPath
$SetupPython = 'C:\Tools\Python3148\python.exe'
& $SetupPython -I -c "import sys,platform,sysconfig; assert sys.version_info[:3] == (3,14,8); assert sys.platform == 'win32' and platform.machine() == 'AMD64'; assert not sysconfig.get_config_var('Py_GIL_DISABLED')"
if ($LASTEXITCODE -ne 0) { throw 'Expected regular Windows x64 CPython 3.14.8' }
$RuntimeVenv = Join-Path $ApplicationRoot '.venv'
if (Test-Path -LiteralPath $RuntimeVenv) { throw 'Choose a fresh application folder; .venv already exists' }
& $SetupPython -I -m venv $RuntimeVenv
if ($LASTEXITCODE -ne 0) { throw 'venv creation failed' }
$RuntimePython = Join-Path $RuntimeVenv 'Scripts\python.exe'
& $RuntimePython -I -m pip --isolated --disable-pip-version-check install --require-hashes --only-binary=:all: --index-url https://pypi.org/simple -r (Join-Path $ApplicationRoot 'requirements.txt')
if ($LASTEXITCODE -ne 0) { throw 'Runtime dependency installation failed' }
& $RuntimePython -I -m pip --isolated --disable-pip-version-check check
if ($LASTEXITCODE -ne 0) { throw 'Runtime dependency check failed' }
& $RuntimePython -I -c "import sys,pypdf; assert sys.prefix != sys.base_prefix; assert pypdf.__version__ == '6.19.0'; print(sys.executable); print(pypdf.__file__)"
if ($LASTEXITCODE -ne 0) { throw 'Runtime import check failed' }
```

Inspect the printed paths: the interpreter and pypdf must come from this new
environment. The application checks them again and invokes the engine with
isolated Python imports. A selected invalid Python or existing broken `.venv`
fails; it does not silently fall back. You can select a different trusted absolute
interpreter with `-PythonPath`, as described in full application help.

For ebooks, install Calibre separately and use `-CalibrePath` when its trusted
`ebook-convert.exe` is outside safe automatic discovery locations. The application
checks version 9.15.0. PDF-only processing does not inspect or require Calibre.
WinBookSplit never downloads dependencies while processing, changes machine PATH
or modifies global execution policy. It runs with your account's privileges;
neither dependency validation nor the documented loopback observation is a
network isolation guarantee for Calibre/plugins or external ebook resources.

If Windows blocks an unsigned script, follow your organization's script policy.
Do not disable that policy, elevate, or apply a global execution-policy change.
The BAT's process-scoped invocation does not change machine/user policy or
override enforced organizational policy.

## Check an archive's checksum

The application scripts and candidate are unsigned. No public v1.0.0 release
asset is claimed here. For an archive you have obtained, replace the example
path and compute the hash before extraction:

```powershell
$ArchivePath = 'C:\Downloads\WinBookSplit-candidate.zip'
Get-FileHash -LiteralPath $ArchivePath -Algorithm SHA256
```

Compare all 64 hexadecimal digits with a checksum obtained from the intended
publisher through a trusted source. Stop on a mismatch. A matching SHA-256
checks bytes; it does not establish who published them, sign the scripts or
guarantee SmartScreen acceptance. A future public release must provide its
actual tested ZIP, manifest and checksums before this becomes release verification.

Return to [usage examples](../README.md#scripted-examples). The older dependency
decisions and observations remain in [the historical setup record](codex-v1.0.0/SUPPORT_AND_SETUP.md).
Development setup and tests are documented separately in [tests/README.md](../tests/README.md).
