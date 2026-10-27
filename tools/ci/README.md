# Windows CI checks

The Windows workflow uses the existing `tests/run_tests.py` entry point for
Python and separate explicit Windows PowerShell 5.1 and PowerShell 7 shell
runs. The shell layer gates syntax, selected scaffold analyzer rules and
Pester; historical application analyzer findings remain recorded observations.
Windows Server 2025 hosted results do not replace native Windows 10/11,
real Calibre or rendered fidelity evidence. The unchanged Explorer integration
does not require a repeated human programme.
Checkout fetches full history because an unchanged AC-011 unit reads the
immutable extraction commit with `git cat-file`; a shallow checkout omits that
required historical blob. No historical test is skipped or weakened.

Regular x64 CPython 3.14.8 creates a fresh temporary venv. `requirements-dev.txt`
and `requirements.txt` are installed with required hashes and binary wheels.
`bootstrap.py` downloads Pester from the official PowerShell Gallery and
PSScriptAnalyzer from its official PowerShell GitHub release, checking the same
reviewed package SHA256 values before bounded ordinary-file
extraction, and leaves global module, trust and execution-policy settings alone.
The shared shell harness imports only the selected versioned manifests.

Three separate application jobs use the exact test targets in
`tests/compatibility/runtime-lock.json`: 3.13.16, 3.14.8 and 3.15.0. Each uses its own
runtime environment containing only pypdf 6.21.3 and optional pip; a distinct
3.14.8 driver generates authored fixtures with the development dependencies.
The shared application route runs the existing nine stdlib/pypdf unit modules
and eight PDF/native wrapper cases under both shell hosts. The 3.15 runtime
contains no fixture or renderer dependencies. This bounded Server route explicitly
omits Calibre; actual EPUB/AZW3 and native OS gates remain separate. Numeric
application acceptance comes from the shipped policy, independently of these
exact reproducible test selections. The workflow checks that every advertised
minor has a matrix target, uses no fail-fast cancellation between branches,
and retains finite30-minute reference/20-minute application job bounds.

Reviewed action revisions:

| Official action | Selected release | Full commit |
|---|---|---|
| [checkout](https://github.com/actions/checkout/releases/tag/v7.0.1) | 7.0.1 | `3d3c42e5aac5ba805825da76410c181273ba90b1` |
| [setup-python](https://github.com/actions/setup-python/releases/tag/v7.0.0) | 7.0.0 | `5fda3b95a4ea91299a34e894583c3862153e4b97` |
| [upload-artifact](https://github.com/actions/upload-artifact/releases/tag/v7.0.2) | 7.0.2 | `cf430e030ddbb5b0abf93d22962f4752f3646cd9` |

The exact [Python 3.14.8 Windows x64 distribution](https://github.com/actions/python-versions/releases/tag/3.14.8-36806082737)
and [Python 3.15.0 Windows x64 distribution](https://github.com/actions/python-versions/releases/tag/3.15.0-38076878863)
are listed as stable regular x64 targets in the official
[Actions Python catalog](https://github.com/actions/python-versions/blob/main/versions-manifest.json).
Availability is not an executed CI pass. Developer modules are
[Pester 6.2.0](https://www.powershellgallery.com/packages/Pester/6.2.0) and
[PSScriptAnalyzer 1.25.0](https://www.powershellgallery.com/packages/PSScriptAnalyzer/1.25.0).
Downloads and imports are setup evidence, not executed test passes.

`check_package_inputs.py` checks committed application and support prerequisites;
it creates no package and makes no release/package acceptance claim. Ordinary
push, pull request and manual CI dispatch have only `contents: read`, no secret
references or release writes. Checkout does not persist credentials. Later
ready checks still execute after a test failure, while native failures keep
the job failed.

The reference artifact allowlist contains six generated JSON reports: actual
CIM OS/source facts, tool bootstrap, package prerequisites, Python, shell PS5.1
and shell PS7. Each application artifact contains OS/source facts, package
prerequisites and its application report. Artifact names include source SHA,
role/target and run attempt. `validate_reports.py` requires every declared
report, matching source, Server ProductType, successful non-skipped suites and
the exact application target from the source runtime lock; malformed targets,
partial artifacts and unrelated older green runs
cannot satisfy the gate. Native commands keep their nonzero exit codes. No PDFs,
Documents directory, environment dump or temporary-workspace wildcard is
uploaded. Reports remain available on failure. A proposed workflow is not
actual hosted CI evidence; record run URL, tested commit, results and limitations
after execution.
