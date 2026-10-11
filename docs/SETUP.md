# Set up WinBookSplit

Use Windows 10 **22H2 x64** or Windows 11 x64 with inbox Windows PowerShell 5.1.
PowerShell 7.6.6 is also a tested host. Install a regular GIL Windows x64 CPython
from one of the accepted branches below. The application uses plain **pypdf
6.20.0**; EPUB/AZW3 also require separately installed Windows x64 **Calibre 9.15.0**.
Developer renderers, Pester and fixture packages are not application requirements.
Final-release platform evidence belongs to the exact release archive; these
setup instructions do not certify an untested ZIP or a clean operating system.

## Set up once

Extract the complete application to a short, writable per-user folder, such as
`C:\Users\YourName\Tools\WinBookSplit`. Keep the entry points, `VERSION`,
requirements and `engine` folder together. Install compatible Python separately
from [Python.org](https://www.python.org/downloads/windows/). The exact tested
choices are [3.13.16](https://www.python.org/downloads/release/python-31316/) and
[3.14.8](https://www.python.org/downloads/release/python-3148/); later accepted
stable patches do not need to match either test version.

Open PowerShell in the extracted application directory and run one command:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Setup-WinBookSplit.ps1 -NonInteractive
```

This execution-policy setting applies only to the new PowerShell process; it
does not change global execution policy. Enforced organization policy still
applies. For direct script commands below, use the same
`powershell.exe -NoProfile -ExecutionPolicy Bypass -File` prefix if your local
default policy blocks unsigned scripts.

This explicit setup action discovers a compatible base interpreter, creates
`.venv` at its final application-local path, downloads the reviewed pypdf wheel
from `https://pypi.org/simple`, and verifies package consistency and import origin.
Without `-NonInteractive`, setup explains the interpreter, destination and package
source once and asks before making changes. A healthy existing application
`.venv` is a validated no-op, with no package download or environment change.

If automatic discovery cannot use your installation, supply its ordinary
absolute executable path. Replace this example with your trusted installation:

```powershell
.\Setup-WinBookSplit.ps1 -PythonPath 'C:\Tools\Python314\python.exe' -NonInteractive
```

No environment activation, administrator session, global package installation
or PATH change is needed. Setup installs neither Python nor Calibre. Processing
through the BAT or PowerShell entry point never installs dependencies.

## Check without making changes

```powershell
.\Setup-WinBookSplit.ps1 -Check -NonInteractive
```

`-Check` probes readiness without creating an environment, setup lock or marker,
or installing/downloading packages. It can report ready for an existing trusted
external processing runtime even when the application has no `.venv`. A compatible
base interpreter without the locked package produces a setup-required result.
Use the normal setup command above to create the application environment.

The setup command prints a `[SETUP]` JSON record. Exit 0 means ready or completed;
3 means a prerequisite, package-source or explicit repair action is needed;
6 means setup/rollback could not complete; 130 means cancellation or timeout.
Read the stated reason before retrying. The raw record may contain local paths;
keep it private.

## Repair an unhealthy application environment

An existing broken or incompatible `.venv` is preserved by default. Processing
does not silently fall back around it, and ordinary setup does not replace it.
To authorize repair, select a compatible base outside the application `.venv`:

```powershell
.\Setup-WinBookSplit.ps1 -PythonPath 'C:\Tools\Python314\python.exe' -Repair -NonInteractive
```

Repair moves the old ordinary directory to a unique sibling
`.WinBookSplit-venv-backup-<id>`, then creates the new environment at the original
`.venv` path. A successful repair preserves the backup and reports its location.
When a failed repair can be rolled back safely, the incomplete new environment
is preserved under `.WinBookSplit-incomplete-<id>` and the old one is restored to
its original path. If process shutdown or path ownership cannot be proved, setup
preserves remaining files and reports the failed rollback for inspection.
Backups and incomplete environments are retained artifacts, not usable runtime
locations. Setup does not recursively delete them.

Setup serializes changes with an exclusive lock and refuses unsafe paths,
reparse destinations, concurrent environment replacement or unavailable file
access. If the folder is unwritable, extract a fresh copy to a writable per-user
location. If another setup is active, let it finish. Moving an application with
its `.venv` later requires runtime revalidation or recreation; virtual environments
are not promised to be portable.

## Offline package setup

Place the exact reviewed file `pypdf-6.20.0-py3-none-any.whl` in an ordinary local
wheelhouse directory. Its SHA-256 must be:

```text
f003fc2014814d264fe7dd3f9d435c158e23e1a85a2233f87a0a2d6d21c914ad
```

Then run:

```powershell
.\Setup-WinBookSplit.ps1 -PythonPath 'C:\Tools\Python314\python.exe' -Wheelhouse 'C:\Tools\WinBookSplit-Wheels' -NonInteractive
```

Setup validates and holds that exact wheel file while installing. The offline
path uses no package index and has no network fallback. A missing wheel or hash
mismatch fails before environment changes. The runtime requirements and policy
must agree. Pip configuration and `PIP_*` variables are isolated in the setup
child; your account's configuration is not changed. The application distribution
does not bundle this wheel.

## Python policy

The [shipped policy](../engine/WinBookSplit.Compatibility.json) accepts regular
stable GIL CPython for Windows AMD64, 64-bit:

| Minor branch | Security floor | Exact source-test target |
| --- | --- | --- |
| 3.13 | 3.13.16 | 3.13.16 |
| 3.14 | 3.14.8 | 3.14.8 |

Later stable patches in an accepted minor are compared numerically and remain
accepted without a network lookup. The test targets are reproducible test choices,
not patch-equality gates or a promise that every accepted patch was executed.
The effective floor incorporates a reviewed Windows import-search-path fix
([CVE-2026-12003 / gh-151544](https://github.com/python/cpython/issues/151544));
see the [3.13.15 changelog](https://docs.python.org/3.13/whatsnew/changelog.html#python-3-13-15-final)
and [3.14.7 changelog](https://docs.python.org/3.14/whatsnew/changelog.html#python-3-14-7-final).
Policy version 2 also requires the bundled OpenSSL 3.5.9 update in
[Python 3.13.16](https://www.python.org/downloads/release/python-31316/) and
[Python 3.14.8](https://www.python.org/downloads/release/python-3148/). The
[29 September 2026 OpenSSL advisory](https://openssl-library.org/news/secadv/20260929.txt)
describes excessive memory use from a crafted peer certificate during an
ordinary TLS handshake (CVE-2026-35189). Explicit online setup runs pip over
HTTPS with the selected interpreter, so this application policy excludes
the older bundled TLS versions, including the previously reviewed floors
3.13.15 and 3.14.7. This policy does not declare every older interpreter
secure; Python 3.14.2 also remains below the reviewed security floor.

Python 3.11 and 3.12 were evaluated. Their applicable fixes are in source-only
[3.11.16](https://www.python.org/downloads/release/python-31116/) and
[3.12.14](https://www.python.org/downloads/release/python-31214/); patched official
Windows binaries are unavailable in the reviewed catalog. End users are not
required to compile Python, so these branches are excluded. Prereleases,
free-threaded builds, PyPy, 32-bit/ARM processes and unlisted minors such as 3.15
are also excluded. The exact pypdf 6.20.0 lock and Calibre 9.15.0 check remain
separate from Python's multi-minor policy.

For processing, an explicit `-PythonPath` is authoritative, followed by an existing
application `.venv`. Otherwise discovery probes trusted concrete PATH candidates
and bounded installed-runtime listings. A failed optional launcher listing does
not discard a valid PATH interpreter. Among accepted automatic candidates, the
policy prefers 3.14, then 3.13, and the highest numeric patch in that branch.
It never invokes a launcher install or exec operation. If an alias or short path
cannot pass trust checks, use the ordinary long path to the installed executable;
do not remove the guards.

## Calibre, unsigned scripts and Windows servicing

For ebooks, install [Calibre 9.15.0](https://download.calibre-ebook.com/9.15.0/)
separately and use the application's `-CalibrePath` when its trusted
`ebook-convert.exe` is outside automatic discovery. PDF-only processing does not
inspect or require Calibre. WinBookSplit runs with your account's privileges;
dependency validation is not an OS network sandbox for Calibre/plugins or
external ebook resources.

The scripts are unsigned. If Windows or organization policy blocks execution,
follow that policy. Do not elevate or apply a global execution-policy change.
The BAT uses a process-scoped invocation and does not override enforced policy.

Microsoft's standard Windows 10 servicing ended on **14 October 2025**.
WinBookSplit compatibility does not extend Windows support or require an ESU
purchase. See [Microsoft's lifecycle notice](https://learn.microsoft.com/en-us/lifecycle/announcements/windows-10-end-of-support)
for operating-system servicing options.

## Verify a downloaded application archive

Public release assets are listed on the
[GitHub releases page](https://github.com/PikkuJanne/WinBookSplit/releases).
Use the ZIP, `release-manifest.json` and `SHA256SUMS.txt` from the same release.
Exact-package native results are recorded in the
[release evidence](https://github.com/PikkuJanne/WinBookSplit/tree/main/docs/codex-v1.0.1/evidence).
An unpublished candidate or source-test report is not release certification.

Compute the archive's hash before extraction, replacing this example path:

```powershell
Get-FileHash -LiteralPath 'C:\Downloads\WinBookSplit-v1.0.1.zip' -Algorithm SHA256
```

Compare all 64 hexadecimal digits with the intended publisher's checksum. Stop
on a mismatch. A matching SHA-256 checks bytes; it does not establish publisher
identity, sign the scripts or guarantee SmartScreen acceptance.

Return to [usage examples](../README.md#scripted-examples). Developer dependencies
and exact test pins are documented separately in the
[source testing guide](https://github.com/PikkuJanne/WinBookSplit/blob/main/tests/README.md).
Historical v1.0.0 observations remain in the
[historical setup record](https://github.com/PikkuJanne/WinBookSplit/blob/main/docs/codex-v1.0.0/SUPPORT_AND_SETUP.md).
