[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$RepositoryRoot,
    [Parameter(Mandatory = $true)][string]$ToolRoot,
    [Parameter(Mandatory = $true)][string]$WorkRoot,
    [Parameter(Mandatory = $true)][string]$ReportPath,
    [switch]$FailureProbe
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$outcomes = New-Object -TypeName 'System.Collections.Generic.List[object]'
$report = [ordered]@{
    schema_version = 1
    acceptance_ids = @('AC-009', 'AC-010')
    scope = 'Local shell harness, syntax and static checks; no application success-path, Explorer or Calibre acceptance.'
    host = [ordered]@{
        version = $PSVersionTable.PSVersion.ToString()
        edition = $PSVersionTable.PSEdition
        clr = [Environment]::Version.ToString()
        execution_policy = @((Get-ExecutionPolicy -List) | ForEach-Object {
            @{ scope = $_.Scope.ToString(); policy = $_.ExecutionPolicy.ToString() }
        })
    }
    failure_probe = [bool]$FailureProbe
    imports = @()
    syntax = @()
    scaffold_analysis = @()
    legacy_application_analysis = @()
    legacy_analysis_is_acceptance_pass = $false
    pester = $null
    outcomes = @()
    exit_code = 1
    success = $false
}
$previousTemp = $env:TEMP
$previousTmp = $env:TMP
$reportValidated = $false

try {
    foreach ($directory in @($RepositoryRoot, $ToolRoot, $WorkRoot)) {
        if (-not [IO.Path]::IsPathRooted($directory) -or -not (Test-Path -LiteralPath $directory -PathType Container)) {
            throw 'RepositoryRoot, ToolRoot and WorkRoot must be existing absolute directories.'
        }
    }
    $repository = (Resolve-Path -LiteralPath $RepositoryRoot).ProviderPath.TrimEnd('\')
    $work = (Resolve-Path -LiteralPath $WorkRoot).ProviderPath.TrimEnd('\')
    $tools = (Resolve-Path -LiteralPath $ToolRoot).ProviderPath.TrimEnd('\')
    if ($work.Equals($repository, [StringComparison]::OrdinalIgnoreCase) -or
        $work.StartsWith($repository + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'WorkRoot must be outside the repository.'
    }
    if ($tools.Equals($repository, [StringComparison]::OrdinalIgnoreCase) -or
        $tools.StartsWith($repository + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'ToolRoot must be outside the repository.'
    }
    if (-not [IO.Path]::IsPathRooted($ReportPath)) {
        throw 'ReportPath must be an absolute path.'
    }
    $reportFile = [IO.Path]::GetFullPath($ReportPath)
    if (-not $reportFile.StartsWith($work + '\', [StringComparison]::OrdinalIgnoreCase) -or
        (Test-Path -LiteralPath $reportFile)) {
        throw 'ReportPath must be a new file inside the owned WorkRoot.'
    }
    $reportValidated = $true
    $env:TEMP = $work
    $env:TMP = $work
    Import-Module -Name (Join-Path -Path $repository -ChildPath 'tests\powershell\HarnessSupport.psm1') -Force -ErrorAction Stop
    foreach ($name in @('Pester', 'PSScriptAnalyzer')) {
        $manifest = Resolve-TestToolManifest -ToolRoot $ToolRoot -Name $name
        $expectedVersion = if ($name -eq 'Pester') { '6.2.0' } else { '1.25.0' }
        $module = Import-Module -Name $manifest -Force -PassThru -ErrorAction Stop
        if ($module.Version.ToString() -ne $expectedVersion -or
            -not $module.Path.StartsWith([IO.Path]::GetDirectoryName($manifest) + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw ('Unexpected module version or import origin: ' + $name)
        }
        $report.imports += @{ name = $name; version = $module.Version.ToString(); manifest = $manifest; path = $module.Path }
    }
    $outcomes.Add([pscustomobject]@{ name = 'explicit pinned imports'; passed = $true })

    $application = Join-Path -Path $repository -ChildPath 'WinBookSplit.ps1'
    $scaffoldPaths = @(
        (Join-Path -Path $repository -ChildPath 'tests\Invoke-ShellTests.ps1'),
        (Join-Path -Path $repository -ChildPath 'engine\WinBookSplit.Runtime.ps1'),
        (Join-Path -Path $repository -ChildPath 'engine\WinBookSplit.Process.ps1'),
        (Join-Path -Path $repository -ChildPath 'tests\PSScriptAnalyzerSettings.psd1')
    ) + @((Get-ChildItem -LiteralPath (Join-Path -Path $repository -ChildPath 'tests\powershell') -File) |
        Where-Object { $_.Extension -in @('.ps1', '.psm1') } | ForEach-Object { $_.FullName })
    foreach ($path in @($application) + $scaffoldPaths) {
        $tokens = $null
        $parseErrors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$parseErrors)
        $report.syntax += @{ path = $path.Substring($repository.Length + 1); errors = @($parseErrors | ForEach-Object { $_.Message }) }
        $outcomes.Add([pscustomobject]@{ name = 'syntax: ' + $path.Substring($repository.Length + 1); passed = (@($parseErrors).Count -eq 0) })
    }
    $settings = Join-Path -Path $repository -ChildPath 'tests\PSScriptAnalyzerSettings.psd1'
    foreach ($path in $scaffoldPaths) {
        $findings = @(Invoke-ScriptAnalyzer -Path $path -Settings $settings -ErrorAction Stop)
        $report.scaffold_analysis += @{ path = $path.Substring($repository.Length + 1); findings = @($findings | ForEach-Object {
            @{ rule = $_.RuleName; severity = $_.Severity.ToString(); line = $_.Line; message = $_.Message }
        }) }
        $outcomes.Add([pscustomobject]@{ name = 'scaffold analysis: ' + $path.Substring($repository.Length + 1); passed = ($findings.Count -eq 0) })
    }
    $report.legacy_application_analysis = @(Invoke-ScriptAnalyzer -Path $application -ErrorAction Stop | ForEach-Object {
        @{ rule = $_.RuleName; severity = $_.Severity.ToString(); line = $_.Line; message = $_.Message }
    })
    $outcomes.Add([pscustomobject]@{ name = 'legacy application findings recorded (observation only)'; passed = $true })

    $testFile = if ($FailureProbe) { 'tests\powershell\FailureProbe.ps1' } else { 'tests\powershell\Harness.Tests.ps1' }
    $config = New-PesterConfiguration
    $containers = @(New-PesterContainer -Path (Join-Path -Path $repository -ChildPath $testFile) -Data @{
        RepositoryRoot = $repository; ToolRoot = $ToolRoot; WorkRoot = $work
    })
    if (-not $FailureProbe) {
        $containers += New-PesterContainer -Path (Join-Path -Path $repository -ChildPath 'tests\powershell\Runtime.Tests.ps1') -Data @{
            RepositoryRoot = $repository; WorkRoot = $work; TestPython = $env:WBS_TEST_PYTHON
        }
        $containers += New-PesterContainer -Path (Join-Path -Path $repository -ChildPath 'tests\powershell\Process.Tests.ps1') -Data @{
            RepositoryRoot = $repository; WorkRoot = $work; TestPython = $env:WBS_TEST_PYTHON
        }
    }
    $config.Run.Container = $containers
    $config.Run.PassThru = $true
    $config.Run.Exit = $false
    $config.Output.Verbosity = 'None'
    $config.TestResult.Enabled = $false
    $pesterResult = Invoke-Pester -Configuration $config
    $report.pester = @{
        result = $pesterResult.Result.ToString()
        total = $pesterResult.TotalCount
        passed = $pesterResult.PassedCount
        failed = $pesterResult.FailedCount
        skipped = $pesterResult.SkippedCount
        not_run = $pesterResult.NotRunCount
        failed_containers = $pesterResult.FailedContainersCount
    }
    $outcomes.Add([pscustomobject]@{ name = 'Pester'; passed = ($pesterResult.Result -eq 'Passed' -and $pesterResult.TotalCount -gt 0) })

    # This successful native process is deliberately later than Pester. A known
    # failed fixture must still return nonzero after this and JSON serialization.
    & $env:ComSpec /d /c 'exit 0'
    $nativeExitCode = $LASTEXITCODE
    $outcomes.Add([pscustomobject]@{ name = 'later native success'; passed = ($nativeExitCode -eq 0); exit_code = $nativeExitCode })
    $report.exit_code = Get-TestOutcomeExitCode -Outcomes $outcomes.ToArray()
}
catch {
    $outcomes.Add([pscustomobject]@{ name = 'shell harness exception'; passed = $false; error = $_.Exception.Message })
    $report.exit_code = 1
    [Console]::Error.WriteLine($_.Exception.ToString())
}
finally {
    $env:TEMP = $previousTemp
    $env:TMP = $previousTmp
    $report.outcomes = $outcomes.ToArray()
    $report.success = ($report.exit_code -eq 0)
    if ($reportValidated) { try {
        $json = $report | ConvertTo-Json -Depth 12
        # CreateNew prevents overwriting even if an unexpected file appeared.
        $stream = New-Object -TypeName IO.FileStream -ArgumentList @($reportFile, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try {
            $writer = New-Object -TypeName IO.StreamWriter -ArgumentList @($stream, (New-Object -TypeName Text.UTF8Encoding -ArgumentList $false))
            try { $writer.Write($json) } finally { $writer.Dispose() }
        }
        finally { $stream.Dispose() }
    }
    catch {
        $report.exit_code = 1
        [Console]::Error.WriteLine('Shell report failed: ' + $_.Exception.Message)
    } }
}
exit $report.exit_code
