Set-StrictMode -Version Latest

function Resolve-TestToolManifest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ToolRoot,
        [Parameter(Mandatory = $true)][ValidateSet('Pester', 'PSScriptAnalyzer')][string]$Name
    )
    if (-not [IO.Path]::IsPathRooted($ToolRoot)) {
        throw 'ToolRoot must be an absolute directory.'
    }
    $version = if ($Name -eq 'Pester') { '6.2.0' } else { '1.25.0' }
    $manifest = Join-Path -Path $ToolRoot -ChildPath ($Name + '\' + $version + '\' + $Name + '.psd1')
    if (-not (Test-Path -LiteralPath $manifest -PathType Leaf)) {
        throw ('Required isolated module manifest is missing: ' + $manifest)
    }
    return (Resolve-Path -LiteralPath $manifest -ErrorAction Stop).ProviderPath
}

function Get-TestOutcomeExitCode {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Outcomes)
    # An empty run cannot succeed. Later successful stages never erase a failure.
    if ($Outcomes.Count -eq 0) { return 1 }
    foreach ($outcome in $Outcomes) {
        if (-not $outcome.passed) { return 1 }
    }
    return 0
}

Export-ModuleMember -Function Resolve-TestToolManifest, Get-TestOutcomeExitCode
