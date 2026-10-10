# Windows CI checks

The Windows workflow uses the existing `tests/run_tests.py` entry point for
Python and separate explicit Windows PowerShell 5.1 and PowerShell 7 shell
runs. The shell layer gates syntax, selected scaffold analyzer rules and
Pester; historical application analyzer findings remain recorded observations.
Windows Server 2025 hosted results do not replace local Windows 11, real
Calibre, rendered fidelity, Explorer or the closed human evidence.
Checkout fetches full history because an unchanged AC-011 unit reads the
immutable extraction commit with `git cat-file`; a shallow checkout omits that
required historical blob. No historical test is skipped or weakened.

Regular x64 CPython 3.14.8 creates a fresh temporary venv. `requirements-dev.txt`
and `requirements.txt` are installed with required hashes and binary wheels.
`bootstrap.py` downloads exact developer modules from the official PowerShell
Gallery, checks fixed package SHA256 values before bounded ordinary-file
extraction, and leaves global module, trust and execution-policy settings alone.
The shared shell harness imports only the selected versioned manifests.

Reviewed action revisions:

| Official action | Selected release | Full commit |
|---|---|---|
| [checkout](https://github.com/actions/checkout/releases/tag/v7.0.1) | 7.0.1 | `3d3c42e5aac5ba805825da76410c181273ba90b1` |
| [setup-python](https://github.com/actions/setup-python/releases/tag/v7.0.0) | 7.0.0 | `5fda3b95a4ea91299a34e894583c3862153e4b97` |
| [upload-artifact](https://github.com/actions/upload-artifact/releases/tag/v7.0.2) | 7.0.2 | `cf430e030ddbb5b0abf93d22962f4752f3646cd9` |

The exact [Python 3.14.8 Windows x64 distribution](https://github.com/actions/python-versions/releases/tag/3.14.8-36806082737)
is available to the pinned setup action. Developer modules are
[Pester 6.2.0](https://www.powershellgallery.com/packages/Pester/6.2.0) and
[PSScriptAnalyzer 1.25.0](https://www.powershellgallery.com/packages/PSScriptAnalyzer/1.25.0).
Downloads and imports are setup evidence, not executed test passes.

`check_package_inputs.py` checks committed application and support prerequisites;
it creates no package and makes no release/package acceptance claim. Ordinary
push, pull request and manual CI dispatch have only `contents: read`, no secret
references or release writes. Checkout does not persist credentials. Later
ready checks still execute after a test failure, while native failures keep
the job failed.

The artifact allowlist contains only five generated JSON reports: tool
bootstrap, package prerequisites, Python, shell PS5.1 and shell PS7. No PDFs,
Documents directory, environment dump or temporary-workspace wildcard is
uploaded. Reports remain available on failure. A proposed workflow is not
actual hosted CI evidence; record run URL, tested commit, results and limitations
after execution.
