[CmdletBinding()]
param([AllowEmptyString()][string]$ManifestPath = '', [AllowEmptyString()][string]$OutputPath = '')

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$exitCode = 2
$record = [ordered]@{ protocol='winbooksplit.support-export'; version=1; status='failed'; code='support_export_failed'; exit_code=2; written_count=0 }
try {
    . (Join-Path $PSScriptRoot 'engine\WinBookSplit.Support.ps1')
    $null = Export-WinBookSplitSupportSummary -ManifestPath $ManifestPath -OutputPath $OutputPath
    $exitCode = 0
    $record.status = 'success'; $record.code = 'support_export_complete'; $record.exit_code = 0; $record.written_count = 1
}
catch { Write-Host 'Diagnostic export failed. Choose a valid finalized run manifest and a new local output file.' }
Write-Host ('[SUPPORT-EXPORT] ' + ($record | ConvertTo-Json -Compress))
exit $exitCode
