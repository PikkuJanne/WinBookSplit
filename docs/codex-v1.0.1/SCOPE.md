# Controlling scope and decisions

## Required result
Public **v1.0.1** of the existing local WinBookSplit application, with Windows 10 22H2 x64 and Windows 11 x64 support backed by real native automation; no exact Python-maintenance-version requirement; tested regular x64 CPython minor branches; robust discovery; explicit application-local dependency setup. Preserve the existing batch shortcut, console, PDF/EPUB/AZW3 workflow, source preservation and output policies.

The user authorized start-to-finish changes and publication and asked to keep human testing minimal. This supersedes conflicting historical Windows-11-only, exact-patch and manual-acceptance gates from v1.0.0. Historical results are not changed to pretend they covered the new release.

## Python acceptance policy
Evaluate **3.11, 3.12, 3.13 and 3.14**. Initial preferred release scope includes all four after verification. At least two genuinely tested minor branches (preferably 3.13 and 3.14) are mandatory; do not call a one-patch or one-minor environment broad compatibility. Exclusions from the candidate list require specific syntax/API/package, available trusted Windows distribution or material security evidence recorded before freeze, not merely inconvenience. Do not expand to 3.15 just because it exists upstream; it is outside this bounded patch target unless separately approved.

Use numeric stable-version comparison with a per-branch security/compatibility floor and finite documented exclusions where justified. Accept later **stable patches** in an advertised minor without an equality gate or live network consultation. Distinguish accepted-by-policy from exact-version tested. Security floors require cited applicability; the latest release-build interpreter is not automatically the minimum. Do not promise 3.14.2 is secure or necessarily accepted before the review. Do not add an unsafe override to bypass material security exclusions.

An end user must not be required to compile Python from source. Source-only maintenance series and missing trusted Windows distributions may affect which branches can honestly be supported. Keep an ordinary compatible installed Python usable. No global replacement of the user's Python.

## Dependencies and setup
Retain pypdf 6.20.0 plus its reviewed hashes as the starting application lock, subject to a new upstream advisory check. Keep exact reproducible build/test versions separate from the runtime policy. Do not silently replace the lock with `pypdf>=...` or `pypdf<7`. Review all private-hook dependencies before a package update. Calibre is still optional for PDF input. Test the current converter on Windows 10; broadening its version range is not required for this Python/Windows patch. If an actual Win10 blocker requires a narrowly verified converter change, record and test it.

Add `Setup-WinBookSplit.ps1`: discover a suitable **base interpreter without requiring pypdf first**, create/validate an app-local .venv, install the hash-locked runtime package and run preflight. Processing remains read-only with respect to dependency installation. Setup may download locked packages only as an explicit action; it must not install Python/Calibre automatically, change global packages/PATH/execution policy or request elevation.

## Deliberate exclusions
No application rewrite, website, server-side processing, new GUI, native ebook output, OCR, AI inference, DRM changes, auto-updater, Python bundling/freezing, broad installer framework, 32-bit/ARM/free-threaded/PyPy claims or expanded UNC/long-path scope. No changes to sound/media tools. Keep existing PDF-policy exclusions. No paid services or code-signing requirement. Scripts may remain unsigned with honest disclosure.

## Testing and completion
Use HUMAN_TESTING.md and TESTING.md instead of rerunning the old owner checklist. Reuse existing automated safety/fidelity fixtures. Preserve real Windows 10/11 final-package execution and actual EPUB/AZW3 conversion; Windows Server results do not substitute for Windows 10. Do not require a clean OS reinstall: a clean application extraction and isolated dependencies on an existing native host/VM suffice, with that limitation documented.

Already-delegated ordinary work includes code, scoped setup, commits, pushes, PRs, passing reviewed merges and v1.0.1 publication. Respect actual tool approval/protection. No force push, reset hard, automatic stash, repository-setting changes, deletion of user work, tag movement or clobbering published assets. The bundle's old baseline is not a reset target.

Completion means publicly verified release bytes, fixed tag/source identity and synchronized main, not only passing local tests. Genuine unavailable infrastructure must be reported as blocked while finishing all independent work. It must not be disguised as a test pass or produce a release with false Windows 10 claims.
