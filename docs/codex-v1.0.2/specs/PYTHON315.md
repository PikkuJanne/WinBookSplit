# Python 3.15 compatibility contract

## Required acceptance, not an optional matrix entry

Target regular stable GIL CPython Windows x64 3.15.0. Retain the existing accepted 3.13/3.14 security floors (currently 3.13.16 and 3.14.8) unless an independently verified new advisory requires a documented change. Test official current patch artifacts; record provenance, executable path, actual `sys.version_info`, `sys.implementation`, architecture, GIL/build status, OpenSSL and package origin. Preserve ordinary per-user installs. Python.org already lists stable 3.15.0 [S1]; do not label it a prerelease.

Add a 3.15 branch to the existing shared policy. Use numeric patch comparisons, not string/float comparisons or `== 3.15.0`. Initial target floor is 3.15.0 subject to actual primary-source review. Increment the policy version and review date with real sources. Tests/CI may record exact reproducible versions while the runtime accepts later compatible stable patches. Keep malformed policy rejection and no online policy updates during processing. Do not copy a planning template into the runtime as proof of review.

Unknown 3.16+ is `unsupported_python`/not yet validated, not `python_security_floor`. Older 3.14.2 remains a security-floor rejection with its observed version/minimum visible. New security restrictions must cite affected versions and relevance, not simply whichever Python happened to be installed on the developer machine. Do not broaden to 3.11/3.12, PyPy, ARM, x86, prereleases or free-threaded builds in this patch.

## Probe and import behavior

Use one shared policy source across setup, preflight, diagnostics, selection and tests. Audit all exact version checks, allowed-minor assumptions, source-test manifests, bootstrap scripts and release allowlists. Change current active expectations, not historical reports.

Preserve isolated execution, bytecode suppression and trustworthy module origins. Recheck 3.15 startup/encoding changes [S7] with malicious book-directory/CWD shadow fixtures. Do not import arbitrary `.pth`/startup material as a workaround or claim isolation is a sandbox. Probe a base Python before requiring pypdf for setup; final readiness must validate the package in the selected interpreter. Do not substitute a helper/test interpreter for the application runtime and still report 3.15 success.

Exercise the real pypdf reader/writer, document-policy guards, ctypes/process helpers, conversion, CLI protocol and setup/venv/pip paths under actual 3.15.0. Keep the reviewed pypdf 6.20.0 lock by default. Its private `_handle_encryption` hook and reader constructor options require regression tests, not a version-string assertion alone. If an actual compatible-library fix is needed, keep it minimal and test the supported branches and the existing unsupported-document boundary.

## Test dependencies are not end-user requirements

A cp314-only dev wheel lock must not force end users off 3.15. Generate deterministic authored fixtures using an existing supported developer runtime when necessary, with recorded provenance, then run PDF/CLI/setup/native acceptance under 3.15 using the runtime-only environment. Use per-minor dev locks only where necessary and reviewed, retain required hashes, and keep renderer/reportlab/Pillow/Pester out of the runtime distribution. Missing optional renderer wheels are not a reason to skip the actual 3.15 PDF split tests.

## Discovery selection and upgrades

Preserve precedence: explicit `-PythonPath`, authoritative existing application `.venv`, then automatic discovery. A healthy accepted venv is a no-op on setup even when PATH changes to 3.15. A broken/stale/rejected venv is preserved and explained; normal setup never silently replaces it, and processing never falls through to another interpreter behind the user's back. An explicit valid base plus `-Repair` preserves a backup and creates a tested new venv. Do not require the user to delete `.venv` manually.

With no venv/override, selection should deterministically prefer the highest tested supported branch (3.15, then 3.14, then 3.13), and highest allowed patch within it. Preserve failed optional listing diagnostics but continue trusted direct discovery. A `py.exe` that forwards `-0p` to Python must not hide a valid regular PATH Python. Use bounded installed-runtime listings only, never a launcher operation that may auto-install. Report the actual chosen interpreter, not the output of a separate `python --version` command.
