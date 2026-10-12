# Observed baseline and incident

Observed remote main: `8fd38e27925985b86d60ddf0950e15da41189ed6`. Published v1.0.1 reports source `6f2c2e5d4d5fc484bcb1530474f09957e09775ef`. Live repo/release data was read through the GitHub connector. The JSON companion contains API-reported asset identities; refresh and independently verify during V102-01. None is a reset target. [R1–R6 in SOURCES.md]

## Incident facts — keep separate

1. Python 3.14.2 was rejected by the reviewed security floor. The goal is not to force acceptance by disabling that floor.
2. After the user installed Python 3.15.0, direct invocation of `Setup-WinBookSplit.ps1 -NonInteractive` was blocked as unsigned. PowerShell rejected the file before its body ran. This output cannot prove pypdf or Python 3.15 execution failed.
3. Current policy lists only 3.13 and 3.14. The shared evaluator has a distinct unsupported-minor reason, but the console action builder in WinBookSplit.ps1 still constructs a bare `& '<...>Setup-WinBookSplit.ps1'` action.
4. Upstream Python.org identifies 3.15.0 as stable, released October 9, 2026. [S1] Treat a new minor as unreviewed until tested, not as inherently insecure.
5. The user invoked from `C:\Windows\System32`; the application path was `C:\projects\WinBookSplit-main`. Tests must use unrelated CWD as well as an application directory with spaces. Do not embed the user's account path in public fixtures.

## Useful existing integration points

- `engine/WinBookSplit.Compatibility.json` and `.ps1`: shared numeric policy and identity checks.
- `engine/WinBookSplit.Runtime.ps1`: discovery/preflight; keep direct/venv authority and trustworthy origins.
- `Setup-WinBookSplit.ps1` and `engine/WinBookSplit.Setup.ps1`: explicit setup/check/repair/offline workflow.
- `WinBookSplit.ps1`: `Get-WinBookSplitDependencySummary` and setup-action formatting.
- `tests/powershell/Outcome.Tests.ps1`: existing summary tests; extend with executable action regressions.
- `tools/compatibility/{Invoke-NativeAcceptance.ps1,Invoke-FinalZipPs7.ps1,native_smoke.py,final_zip_ps7.py,kit_builder.py,prepare_fixtures.py}`: reuse, do not restart native verification from scratch.
- `.github/workflows/windows-ci.yml`, `tests/run_tests.py`, `tools/release/`: inspect current schemas/allowlists before modifying.
- `AGENTS.md`: current active v1.0.1 header coexists with older routing; reconcile rather than prepend contradictory rules.

No application code is patched by this bundle. File listings/selected content were inspected, not a full security audit or native execution.
