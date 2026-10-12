<#
.SYNOPSIS
Explicitly prepare or check the application-local WinBookSplit Python environment.
.DESCRIPTION
Selects a trusted policy-supported base Python without requiring pypdf first.
Setup creates a fresh .venv and installs only the separately hash-locked runtime
requirement. Processing never runs installation. No activation, elevation,
global packages/PATH change or global execution-policy change is required.
.PARAMETER PythonPath
Trusted absolute base Python executable. Discovery is read-only when omitted.
.PARAMETER Check
Read-only validation of the existing application .venv; never installs or repairs.
.PARAMETER NonInteractive
Explicitly authorize unattended setup with no prompts; missing prerequisites fail.
.PARAMETER Wheelhouse
Existing ordinary local directory containing the exact reviewed pypdf wheel.
Installation is offline and hash-locked; there is no online fallback.
.PARAMETER Repair
Explicitly preserve a broken .venv in a unique backup before creating a fresh one.
Healthy environments are unchanged. The backup is not overwritten or deleted.
.EXAMPLE
.\Setup-WinBookSplit.ps1
.EXAMPLE
.\Setup-WinBookSplit.ps1 -Check
.EXAMPLE
.\Setup-WinBookSplit.ps1 -PythonPath 'C:\Tools\Python\python.exe' -NonInteractive -Wheelhouse 'C:\Tools\Wheels'
.NOTES
Unsigned; follow enforced organizational script policy. Full setup and runtime
compatibility requirements are in docs/SETUP.md. Output includes a structured
[SETUP] result and a meaningful native exit status.
#>

[CmdletBinding()]
param([string]$PythonPath, [switch]$Check, [switch]$NonInteractive, [string]$Wheelhouse, [switch]$Repair)

# Machine outcomes and path diagnostics use UTF-8 through every launcher.
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)

Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
try {
    . (Join-Path $PSScriptRoot 'engine\WinBookSplit.Setup.ps1')
    $result=Invoke-WinBookSplitSetup -ApplicationRoot $PSScriptRoot -PythonPath $PythonPath -Check:$Check -NonInteractive:$NonInteractive -Wheelhouse $Wheelhouse -Repair:$Repair
} catch {
    $result=[pscustomobject]@{protocol='winbooksplit.setup';schema_version=1;status='failed';code='setup_failed';exit_code=6;changed=$false;message=$_.Exception.Message}
}
Write-Host $result.message
Write-Host ('[SETUP] '+($result|ConvertTo-Json -Depth 12 -Compress))
$setupExitCode=[int]$result.exit_code
# Only a zero-argument console launch waits for its result to be read. Explicit
# options, redirected automation and cancellation never acquire another prompt.
if ($PSBoundParameters.Count -eq 0 -and $setupExitCode -ne 130 -and
    -not [Console]::IsInputRedirected -and -not [Console]::IsOutputRedirected) {
    try { $null=Read-Host 'Press Enter to close setup' }
    catch { Write-Verbose 'The console wait was unavailable; preserving the setup exit status.' }
}
exit $setupExitCode
