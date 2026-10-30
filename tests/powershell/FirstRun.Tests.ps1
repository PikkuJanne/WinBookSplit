param([string]$RepositoryRoot, [string]$WorkRoot)

BeforeAll {
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Runtime.ps1')
    $tokens = $null; $parseErrors = $null
    $application = [Management.Automation.Language.Parser]::ParseFile(
        (Join-Path $RepositoryRoot 'WinBookSplit.ps1'), [ref]$tokens, [ref]$parseErrors)
    if (@($parseErrors).Count -ne 0) { throw 'The application has syntax errors.' }
    $summary = @($application.FindAll({ param($node)
        $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -ceq 'Get-WinBookSplitDependencySummary'
    }, $true))
    if ($summary.Count -ne 1) { throw 'Expected one application dependency summary.' }
    . ([scriptblock]::Create($summary[0].Extent.Text))
}

Describe 'Functional repair missing-package classification (SYNTHETIC discovery input)' -Tag 'FUNCTIONAL-FIRST-RUN' {
    It 'reports setup_required after a verified supported interpreter lacks pypdf' {
        # This is a production resolver/summary regression with authored discovery
        # evidence. No interpreter runs and no environment or installer is created.
        # It is separate from the unchanged BAT/native cold-start test.
        $app = Join-Path $TestDrive 'unprepared application'
        $engine = Join-Path $app 'engine'
        $null = New-Item -ItemType Directory -Path $engine
        Copy-Item -LiteralPath (Join-Path $RepositoryRoot 'engine\WinBookSplit.Compatibility.json') -Destination $engine
        $script:firstRunVersion = '3.15.0'
        $script:firstRunPath = Join-Path $WorkRoot 'SYNTHETIC supported Python\python.exe'
        Mock Resolve-WinBookSplitAutomaticPython {
            $Attempts.Add([pscustomobject]@{
                Path = $script:firstRunPath; Source = 'manager'; Stage = 'package'; Accepted = $false
                Reason = 'missing_pypdf'; Message = 'SYNTHETIC: compatible interpreter, absent locked pypdf.'
                Probe = $null; ObservedVersion = $script:firstRunVersion; IdentityVerified = $true
            })
            return $null
        }
        Mock Find-WinBookSplitSetupBase { return [pscustomobject]@{ Path = $script:firstRunPath } }
        $failure = $null
        try { $null = Resolve-WinBookSplitRuntime -ApplicationRoot $app }
        catch { $failure = $_.Exception }
        $failure | Should -Not -BeNullOrEmpty
        $summary = Get-WinBookSplitDependencySummary -Failure $failure -ApplicationRoot $app
        $summary.Reason | Should -BeExactly 'missing_pypdf'
        $summary.Stage | Should -BeExactly 'package'
        $summary.ObservedVersion | Should -BeExactly $script:firstRunVersion
        $summary.ObservedPath | Should -BeExactly $script:firstRunPath
        $summary.ActionIsCommand | Should -BeTrue
        [IO.Directory]::Exists((Join-Path $app '.venv')) | Should -BeFalse
        # Intentionally red until checkpoint 3: a package failure is not a
        # missing interpreter. Do not retain the obsolete runtime_not_found result.
        $failure.Data['Code'] | Should -BeExactly 'setup_required'
    }
}
