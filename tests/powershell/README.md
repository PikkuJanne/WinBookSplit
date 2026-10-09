# Isolated PowerShell harness tools

The shell layer uses [Pester 6.2.0](https://www.powershellgallery.com/packages/Pester/6.2.0)
and [PSScriptAnalyzer 1.25.0](https://www.powershellgallery.com/packages/PSScriptAnalyzer/1.25.0).
Save these exact versions to a new directory outside the checkout from an
already available PowerShellGet setup. No global module install or repository
trust-setting change is required:

```powershell
$ShellToolRoot = Join-Path $env:TEMP ('WinBookSplit-shell-tools-' + [guid]::NewGuid().ToString('N'))
if (Test-Path -LiteralPath $ShellToolRoot) { throw 'Choose a new tool directory' }
$null = New-Item -ItemType Directory -Path $ShellToolRoot -ErrorAction Stop
Save-Module -Name Pester -RequiredVersion 6.2.0 -Repository PSGallery -Path $ShellToolRoot -ErrorAction Stop
Save-Module -Name PSScriptAnalyzer -RequiredVersion 1.25.0 -Repository PSGallery -Path $ShellToolRoot -ErrorAction Stop
```

Pass this absolute directory to `tests/run_tests.py --tool-root`. The wrapper
imports only `Pester/6.2.0/Pester.psd1` and
`PSScriptAnalyzer/1.25.0/PSScriptAnalyzer.psd1` below it and verifies both imported
versions and module paths. A missing pin fails; inbox Pester never substitutes.
Modules remain external developer tools and are excluded from the application
and release package. Setup uses the network; local tests do not download tools.
Record SHA-256 identities of the actual saved files with the test evidence.
Review new downloads before reusing a tool directory.

The central runner uses each explicitly supplied shell executable with
`-NoProfile -NonInteractive -ExecutionPolicy RemoteSigned`. This process-scoped
option allows the trusted local harness and downloaded modules; it does not
change stored user/machine policies and cannot override enterprise policy.
The wrapper records the child policy list. A policy or import rejection fails
the layer; it does not count as a host test pass. The central wrapper command
loads only the trusted repository script from its absolute path into a script
block; run paths are passed as environment data. No document or bookmark text
is evaluated as PowerShell code.

`Harness.Tests.ps1` covers sticky failure aggregation, exact tool discovery,
rejected report destinations and owned temporary fixtures. Pester TestDrive
uses the runner's external work directory; Pester and the central runner own
their respective cleanup. `FailureProbe.ps1` is selected only by
`--failure-probe pester` and deliberately fails. Its filename excludes it from
normal `*.Tests.ps1` discovery. A later successful native command and JSON
report cannot turn this probe into success.

`Outcome.Tests.ps1` pins all final outcome codes independently under each actual
host's integer representation. Authored records check cancellation/timeout and
validated retained incomplete output. Trusted AST-loaded application functions
run with authored transport and console writer controls to verify that later log
errors preserve the engine evidence and primary failures. Dependency functions
must propagate cancellation/timeout without trying another candidate. These are
function/structural tests; real PS/BAT/Calibre outcome acceptance remains a
separate route.

Syntax errors in the application and new harness fail the shell layer. New
harness files must also pass the defect/security/PowerShell 5.1 compatibility
rules in `../PSScriptAnalyzerSettings.psd1`. The unchanged application is checked
with the default analyzer rules and its existing findings are recorded in
`legacy_application_analysis`; recording them is an observation, not a static
acceptance pass or a repair. Runtime static/behavior gates remain open for later
tasks. These harness checks do not certify splitting, Explorer interaction,
Calibre conversion or a release package.
