# Set up WinBookSplit

The [30 October functional repair](functional-repair/STATUS_NEXT.md) requires
interactive first launch to offer the existing application-local setup with
explicit consent and resume the original operation. That behavior is not yet
implemented at checkpoint 1. The v1.0.3 statement below that processing never
installs is a baseline description, not a prohibition on this requested repair.
Read-only, preview and noninteractive operations must continue without prompts
or installation.

Use Windows 10 **22H2 x64** or Windows 11 x64 with inbox Windows PowerShell 5.1.
PowerShell 7.6.6 is also a tested host. Install a regular GIL Windows x64 CPython
from one of the accepted branches below. The application uses plain **pypdf
6.21.3**; EPUB/AZW3 also require separately installed Windows x64 **Calibre 9.15.0**.
Developer renderers, Pester and fixture packages are not application requirements.
Final-release platform evidence belongs to the exact release archive; these
setup instructions do not certify an untested ZIP or a clean operating system.

## Set up once

Extract the complete application to a short, writable per-user folder, such as
`C:\Users\YourName\Tools\WinBookSplit`. Keep the entry points, `VERSION`,
requirements and `engine` folder together. Use a compatible existing Python, or
obtain one separately
from [Python.org](https://www.python.org/downloads/windows/) only if no accepted
installation is already available. Regular x64 Python 3.15 is supported; do not
downgrade it or install a precisely matching patch. The exact source-test choices
are [3.15.0](https://www.python.org/downloads/release/python-3150/),
[3.14.8](https://www.python.org/downloads/release/python-3148/) and
[3.13.16](https://www.python.org/downloads/release/python-31316/); later accepted
stable patches do not need to match those test versions.

Open PowerShell in the extracted application directory and run one command:

```powershell
.\Setup-WinBookSplit.bat -NonInteractive
```

The BAT resolves inbox Windows PowerShell from the Windows system location
and uses Bypass only for that child process. Internet marks remain intact and
persistent execution-policy scopes are unchanged. Enforced organization controls
still apply. This complete alternative can be pasted into PowerShell:

```powershell
& ([System.IO.Path]::Combine([Environment]::SystemDirectory, 'WindowsPowerShell\v1.0\powershell.exe')) -NoProfile -ExecutionPolicy Bypass -File '.\Setup-WinBookSplit.ps1' -NonInteractive
```

This explicit setup action discovers a compatible base interpreter, creates
`.venv` at its final application-local path, downloads the reviewed pypdf wheel
from `https://pypi.org/simple`, and verifies package consistency and import origin.
Without `-NonInteractive`, setup explains the interpreter, destination and package
source once and asks before making changes. A healthy existing application
`.venv` is a validated no-op, with no package download or environment change.

If automatic discovery cannot use your installation, supply its ordinary
absolute executable path. Replace this example with your trusted installation:

```powershell
.\Setup-WinBookSplit.bat -PythonPath 'C:\Tools\Python\python.exe' -NonInteractive
```

No environment activation, administrator session, global package installation
or PATH change is needed. Setup installs neither Python nor Calibre. Processing
through the BAT or PowerShell entry point never installs dependencies.

## Check without making changes

```powershell
.\Setup-WinBookSplit.bat -Check -NonInteractive
```

`-Check` probes readiness without creating an environment, setup lock or marker,
or installing/downloading packages. It can report ready for an existing trusted
external processing runtime even when the application has no `.venv`. A compatible
base interpreter without the locked package produces a setup-required result.
Use the normal setup command above to create the application environment.

Discovery reports a reason, stage and source for rejected routes. It distinguishes
a rejected manager identity or entry point, a failed installed-runtime listing,
a recognized manager with no runtimes, an interpreter policy rejection, and a
missing or wrong-origin pypdf package. Failed alternatives remain in the local
diagnostic record. A completed package query alone does not mean a manager or
interpreter was accepted. A Python path/version is reported as observed only
after its independent isolated identity check.

Automatic discovery uses verified current-user Python Install Manager registration
and its installed-only runtime listing. It does not install Python or change
aliases. An existing compatible base without pypdf needs application setup; it
does not need Python reinstallation. Follow the printed **Action (PowerShell)**
verbatim in PowerShell. It quotes literal paths and may name a verified base for
an explicit repair; `-Check` itself never executes that action. Keep the full
structured records local because they can contain workstation paths and probe
details. Human messages escape control characters.

After script entry the setup command prints a `[SETUP]` JSON record with the
actual selected runtime/reason and host executable/version/policy scopes. A host
policy error before script entry has no setup record and is not a Python failure. Exit 0 means ready or completed;
3 means a prerequisite, package-source or explicit repair action is needed;
6 means setup/rollback could not complete; 130 means cancellation or timeout.
Read the stated reason before retrying. The raw record may contain local paths;
keep it private.

## Repair an unhealthy application environment

An existing broken or incompatible `.venv` is preserved by default. Processing
does not silently fall back around it, and ordinary setup does not replace it.
To authorize repair, select a compatible base outside the application `.venv`:

```powershell
.\Setup-WinBookSplit.bat -PythonPath 'C:\Tools\Python\python.exe' -Repair -NonInteractive
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

Place the exact reviewed file `pypdf-6.21.3-py3-none-any.whl` in an ordinary local
wheelhouse directory. Its SHA-256 must be:

```text
c8ca4d9a360dbd5d4b9bb7a8389e2a9ae6a106bee722b496f918040ba2daa17e
```

Then run:

```powershell
.\Setup-WinBookSplit.bat -PythonPath 'C:\Tools\Python\python.exe' -Wheelhouse 'C:\Tools\WinBookSplit-Wheels' -NonInteractive
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
| 3.15 | 3.15.0 | 3.15.0 |

Later stable patches in an accepted minor are compared numerically and remain
accepted without a network lookup. The test targets are reproducible test choices,
not patch-equality gates or a promise that every accepted patch was executed.
The effective floor incorporates a reviewed Windows import-search-path fix
([CVE-2026-12003 / gh-151544](https://github.com/python/cpython/issues/151544));
see the [3.13.15 changelog](https://docs.python.org/3.13/whatsnew/changelog.html#python-3-13-15-final)
and [3.14.7 changelog](https://docs.python.org/3.14/whatsnew/changelog.html#python-3-14-7-final).
Policy version 3 also requires the bundled OpenSSL 3.5.9 update in
[Python 3.13.16](https://www.python.org/downloads/release/python-31316/) and
[Python 3.14.8](https://www.python.org/downloads/release/python-3148/). Stable
[Python 3.15.0](https://www.python.org/downloads/release/python-3150/) includes
the Windows landmark fix in its
[final-tag beta-3 NEWS](https://raw.githubusercontent.com/python/cpython/v3.15.0/Misc/NEWS.d/3.15.0b3.rst)
and OpenSSL 3.5.9 in its
[Windows SBOM](https://www.python.org/ftp/python/3.15.0/python-3.15.0-amd64.exe.spdx.json). The
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
free-threaded builds, PyPy and 32-bit/ARM processes are also excluded. Unlisted
minors such as 3.16 are not yet validated for this release; that reason does not
label an unknown minor insecure. The exact pypdf 6.21.3 lock and Calibre 9.15.0 check remain
separate from Python's multi-minor policy.

For processing, an explicit `-PythonPath` is authoritative, followed by an existing
application `.venv`. Otherwise discovery probes trusted concrete PATH candidates
and bounded installed-runtime listings. A failed optional launcher listing does
not discard a valid PATH interpreter. Among accepted automatic candidates, the
policy prefers 3.15, then 3.14, then 3.13, and the highest numeric patch in that branch.
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
For v1.0.3 installation, use the named `WinBookSplit-v1.0.3.zip`,
`release-manifest.json` and `SHA256SUMS.txt` from the same public release. GitHub's automatic
**Source code (zip)** and tarball are source checkouts, not the accepted runtime
distribution. The named ZIP contains no Python/Calibre runtime, wheel or `.venv`.
Exact-package native results are recorded in the
[v1.0.3 evidence](https://github.com/PikkuJanne/WinBookSplit/tree/codex/v1.0.3-store-discovery/docs/codex-v1.0.3/evidence).
The earlier Windows10 Store result covers its original package. The owner
closed further manual testing for v1.0.3: the final dependency-updated ZIP's
Windows10 Store run and genuine company-enforced script-policy scenario remain
untested. Agent verification and these limits are recorded separately. Enforced
organization controls still apply. Source tests alone are not native acceptance.

Compute the archive's hash before extraction, replacing this example path:

```powershell
Get-FileHash -LiteralPath 'C:\Downloads\WinBookSplit-v1.0.3.zip' -Algorithm SHA256
```

Compare all 64 hexadecimal digits with the intended publisher's checksum. Stop
on a mismatch. A matching SHA-256 checks bytes; it does not establish publisher
identity, sign the scripts or guarantee SmartScreen acceptance.

Return to [usage examples](../README.md#scripted-examples). Developer dependencies
and exact test pins are documented separately in the
[source testing guide](https://github.com/PikkuJanne/WinBookSplit/blob/main/tests/README.md).
Historical v1.0.0 observations remain in the
[historical setup record](https://github.com/PikkuJanne/WinBookSplit/blob/main/docs/codex-v1.0.0/SUPPORT_AND_SETUP.md).
