# Sources and evidence boundary

Reviewed October 12, 2026. URLs are supplied for reproducible review; recheck dependency/security facts at implementation and release. Existing repo advisory statements are policy history, not independent certification by this handoff.

## R1 — Repository main baseline

https://api.github.com/repos/PikkuJanne/WinBookSplit/branches/main

GITHUB_CONNECTOR: Observed main commit; refresh before work.

## R2 — Published v1.0.1

https://api.github.com/repos/PikkuJanne/WinBookSplit/releases/tags/v1.0.1

GITHUB_CONNECTOR: Release metadata and API digests, not independently downloaded bytes.

## R3 — Working agreement

https://github.com/PikkuJanne/WinBookSplit/blob/8fd38e27925985b86d60ddf0950e15da41189ed6/AGENTS.md

GITHUB_CONNECTOR: Active and historical routing must be reconciled.

## R4 — Shared evaluator

https://github.com/PikkuJanne/WinBookSplit/blob/8fd38e27925985b86d60ddf0950e15da41189ed6/engine/WinBookSplit.Compatibility.ps1

GITHUB_CONNECTOR: Numeric patch checks and distinct unsupported/security-floor reasons already exist.

## R5 — Setup entry point

https://github.com/PikkuJanne/WinBookSplit/blob/8fd38e27925985b86d60ddf0950e15da41189ed6/Setup-WinBookSplit.ps1

GITHUB_CONNECTOR: Reuse setup parameters and structured native exit handling.

## R6 — Console remedy

https://github.com/PikkuJanne/WinBookSplit/blob/8fd38e27925985b86d60ddf0950e15da41189ed6/WinBookSplit.ps1

GITHUB_CONNECTOR_SEARCH_EXCERPT: setupCommand is built as direct unsigned ps1 invocation.

## S1 — Python 3.15.0 release

https://www.python.org/downloads/release/python-3150/

PRIMARY_OPENED: Stable 3.15.0 released October 9, 2026; official Windows artifacts listed. Application acceptance still required.

## S2 — Python 3.15 Windows documentation

https://docs.python.org/3.15/using/windows.html

PRIMARY_OPENED: Verify manager/launcher listing and Windows installation behavior; do not auto-install during discovery.

## S3 — PowerShell execution policies

https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-7.6

PRIMARY_OPENED: RemoteSigned treats Internet-marked scripts differently; Process settings do not override enforced Group Policy.

## S4 — Windows PowerShell command line

https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_powershell_exe?view=powershell-5.1

PRIMARY_OPENED: Use -File and process-scoped policy with correct shell-specific arguments and explicit exit propagation.

## S5 — pypdf 6.20.0 metadata

https://pypi.org/project/pypdf/6.20.0/

PRIMARY_OPENED: Existing reviewed package candidate; metadata is not a Python 3.15 end-to-end pass.

## S6 — GitHub hosted runners

https://docs.github.com/en/actions/reference/runners/github-hosted-runners

PRIMARY_OPENED: Server runner evidence must not be labeled Windows 10 desktop evidence.

## S7 — Python 3.15 changes

https://docs.python.org/3.15/whatsnew/3.15.html

PRIMARY_OPENED: Review startup/stdlib/encoding changes affecting isolated probes, pip and ctypes. Do not infer failures just from a changelog.
