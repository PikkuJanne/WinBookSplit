# Scope and non-negotiable outcomes

Owner authorization: October 12, 2026; implement step by step and publish **v1.0.2**. The owner previously required minimal human testing and local/GitHub synchronization; those constraints remain.

## Fix the user journey

A standard Windows 10/11 x64 user with regular stable Python 3.15.0 must be able to extract the complete release, invoke the new setup BAT, create a local `.venv`, then split through the existing BAT without downgrading Python. The same must work with Internet-marked scripts under a non-GPO RemoteSigned environment. Processing failures must offer commands that actually reach setup, not direct unsigned `.ps1` calls that cannot start.

Preserve numeric maintenance-patch acceptance for existing serviced 3.13/3.14 branches. Add the tested 3.15 branch; future stable 3.15 patches are accepted numerically unless specifically reviewed as bad. Do not promise unknown future Python minors, ARM, x86, free-threaded builds or PyPy. Version 3.15.0 is a required target, not an optional idea to be removed when testing is inconvenient. A real upstream/security incompatibility must be sourced, reproduced and resolved or recorded as a blocking defect; never silently change the finish line.

## Keep and reuse

Retain `WinBookSplit.bat`, `WinBookSplit.ps1`, `Setup-WinBookSplit.ps1`, the shared compatibility policy and evaluator, discovery/setup engines, PDF engine, local logs, existing fixtures and regression tools, package builder and `tools/compatibility` native kit. Add a small setup BAT and shared command-formatting logic as needed. Avoid a parallel installer/runtime manager or a new test framework. Build changes stay proportional to this patch.

Preserve reviewed security floors for older branches, dependency-origin checks, isolated probes, finite processes/streams, local-only processing, safe no-overwrite outputs, parent-aware page coverage, explicit repair with backup, user cancellation and meaningful exit codes. Do not remove encrypted/active-PDF guards or private-reader hooks merely to accept a new interpreter.

Keep pypdf's reviewed exact lock and Calibre scope unless a verified issue on the required path demands a minimal reviewed dependency update. Runtime compatibility is not dependency-update permission. Never loosen to arbitrary pypdf versions. Missing developer-wheel support is not evidence the end-user PDF runtime is incompatible: separate fixture/build tooling from application tests.

## Excluded

No website, rewrite, hosted execution, uploads/telemetry, native EPUB output, OCR, DRM/password feature, global package installs, Python uninstalls, PATH changes, permanent execution-policy changes, automatic Unblock-File sweep, AV/WDAC/AppLocker bypass, elevation, signing purchase, paid runner or unsanctioned remote/self-hosted runner registration. No intermediate public releases. No modifications to existing v1.0.0/v1.0.1 releases.

## Authority

Scoped commits, normal pushes, reviewed passing merges, per-user isolated development/test dependency setup and this exact v1.0.2 publication are delegated. Platform approvals and branch protections remain binding. Do not ask whether to publish again once conditions pass. Do not spend money, create new infrastructure or alter organization security controls. A local dirty worktree is not permission to stash/reset/delete owner work.

## Done means

Reviewed changes + real automation + exact-package Windows 10/11/Python 3.15 evidence + accepted legacy-branch regression + public non-draft non-prerelease v1.0.2 assets anonymously verified + old releases unchanged + final clean local main equal to freshly read remote main. No scope substitution with 'tested on developer Python only'.
