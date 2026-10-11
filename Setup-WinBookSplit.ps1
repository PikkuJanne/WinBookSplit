[CmdletBinding()]
param([string]$PythonPath, [switch]$Check, [switch]$NonInteractive, [string]$Wheelhouse, [switch]$Repair)

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
exit $result.exit_code
