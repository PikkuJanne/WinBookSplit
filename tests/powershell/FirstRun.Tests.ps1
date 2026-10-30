param([string]$RepositoryRoot, [string]$WorkRoot)

BeforeAll {
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Setup.ps1')
    $tokens = $null; $parseErrors = $null
    $application = [Management.Automation.Language.Parser]::ParseFile(
        (Join-Path $RepositoryRoot 'WinBookSplit.ps1'), [ref]$tokens, [ref]$parseErrors)
    if (@($parseErrors).Count -ne 0) { throw 'The application has syntax errors.' }
    foreach ($functionName in @('New-WinBookSplitFailure', 'Get-WinBookSplitDependencySummary',
            'Read-WinBookSplitSessionLine', 'Resolve-WinBookSplitProcessingRuntime',
            'ConvertTo-WinBookSplitDisplayText')) {
        $definitions = @($application.FindAll({ param($node)
            $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
            $node.Name -ceq $functionName
        }, $true))
        if ($definitions.Count -ne 1) { throw ('Expected one application function: ' + $functionName) }
        . ([scriptblock]::Create($definitions[0].Extent.Text))
    }
    function New-FirstRunDependencyFailure {
        param([string]$Code = 'setup_required', [string]$Reason = 'missing_pypdf',
            [string]$Source = 'manager', [bool]$Verified = $true)
        $failure = New-WinBookSplitFailure $Code 'SYNTHETIC dependency preflight failure.'
        $failure.Data['Attempts'] = @([pscustomobject]@{
            Path = 'C:\SYNTHETIC Python\python.exe'; Source = $Source; Stage = 'package'; Accepted = $false
            Reason = $Reason; Message = 'SYNTHETIC controlled discovery evidence.'; Probe = $null
            ObservedVersion = '3.15.0'; IdentityVerified = $Verified
        })
        return $failure
    }
    function New-FirstRunSetupResult {
        param([string]$Status = 'success', [string]$Code = 'setup_complete', [int]$ExitCode = 0)
        return [pscustomobject]@{ protocol = 'winbooksplit.setup'; schema_version = 1
            status = $Status; code = $Code; exit_code = $ExitCode
            message = 'SYNTHETIC setup result; no installer or interpreter was run.'
            backup_path = $null; incomplete_path = $null; rollback_status = 'not_needed'
            operations = @(); attempts = @() }
    }
    function Get-FirstRunCaughtFailure {
        param([scriptblock]$Action)
        try { $null = & $Action }
        catch { return $_.Exception }
        throw 'Expected the component operation to fail.'
    }
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
        # A verified supported interpreter is present; only the package is absent.
        $failure.Data['Code'] | Should -BeExactly 'setup_required'
    }

    It 'does not classify an unverified package allegation as setup_required' {
        $app = Join-Path $TestDrive 'unverified package allegation'
        $engine = Join-Path $app 'engine'
        $null = New-Item -ItemType Directory -Path $engine
        Copy-Item -LiteralPath (Join-Path $RepositoryRoot 'engine\WinBookSplit.Compatibility.json') -Destination $engine
        Mock Resolve-WinBookSplitAutomaticPython {
            $Attempts.Add([pscustomobject]@{
                Path = 'C:\SYNTHETIC unverified Python\python.exe'; Source = 'PATH'; Stage = 'package'; Accepted = $false
                Reason = 'missing_pypdf'; Message = 'SYNTHETIC unverified package allegation.'; Probe = $null
                ObservedVersion = '3.15.0'; IdentityVerified = $false
            })
            return $null
        }
        $failure = Get-FirstRunCaughtFailure { Resolve-WinBookSplitRuntime -ApplicationRoot $app }
        $failure.Data['Code'] | Should -BeExactly 'runtime_not_found'
        [IO.Directory]::Exists((Join-Path $app '.venv')) | Should -BeFalse
    }

    It 'keeps interpreter-only lookup read-only without package-dependent reclassification' {
        $app = Join-Path $TestDrive 'interpreter only'
        $engine = Join-Path $app 'engine'
        $null = New-Item -ItemType Directory -Path $engine
        Copy-Item -LiteralPath (Join-Path $RepositoryRoot 'engine\WinBookSplit.Compatibility.json') -Destination $engine
        Mock Resolve-WinBookSplitAutomaticPython {
            $Attempts.Add([pscustomobject]@{
                Path = 'C:\SYNTHETIC Python\python.exe'; Source = 'manager'; Stage = 'package'; Accepted = $false
                Reason = 'missing_pypdf'; Message = 'SYNTHETIC: proves InterpreterOnly does not apply a package decision.'; Probe = $null
                ObservedVersion = '3.15.0'; IdentityVerified = $true
            })
            return $null
        }
        $failure = Get-FirstRunCaughtFailure { Resolve-WinBookSplitRuntime -ApplicationRoot $app -InterpreterOnly }
        $failure.Data['Code'] | Should -BeExactly 'runtime_not_found'
        Should-Invoke Resolve-WinBookSplitAutomaticPython -Times 1 -Exactly -ParameterFilter { $InterpreterOnly }
        [IO.Directory]::Exists((Join-Path $app '.venv')) | Should -BeFalse
    }

    It 'does not turn below-floor Python into a setup offer because of a package allegation' {
        $app = Join-Path $TestDrive 'below floor package allegation'
        $engine = Join-Path $app 'engine'
        $null = New-Item -ItemType Directory -Path $engine
        Copy-Item -LiteralPath (Join-Path $RepositoryRoot 'engine\WinBookSplit.Compatibility.json') -Destination $engine
        Mock Resolve-WinBookSplitAutomaticPython {
            $Attempts.Add([pscustomobject]@{
                Path = 'C:\SYNTHETIC below-floor Python\python.exe'; Source = 'manager'; Stage = 'package'; Accepted = $false
                Reason = 'missing_pypdf'; Message = 'SYNTHETIC package claim from a below-floor version.'; Probe = $null
                ObservedVersion = '3.14.7'; IdentityVerified = $true
            })
            return $null
        }
        $failure = Get-FirstRunCaughtFailure { Resolve-WinBookSplitRuntime -ApplicationRoot $app }
        $failure.Data['Code'] | Should -BeExactly 'runtime_not_found'
        [IO.Directory]::Exists((Join-Path $app '.venv')) | Should -BeFalse
    }

    It 'classifies verified explicit-base <Reason> as setup_required' -ForEach @(
        @{ Reason = 'missing_pypdf' }, @{ Reason = 'pypdf_version_mismatch' }, @{ Reason = 'pypdf_origin_invalid' }
    ) {
        $app = Join-Path $TestDrive $Reason
        $engine = Join-Path $app 'engine'
        $null = New-Item -ItemType Directory -Path $engine
        Copy-Item -LiteralPath (Join-Path $RepositoryRoot 'engine\WinBookSplit.Compatibility.json') -Destination $engine
        $script:firstRunPackageReason = $Reason
        Mock Test-WinBookSplitPythonCandidate {
            $Attempts.Add([pscustomobject]@{
                Path = $Path; Source = $Source; Stage = 'package'; Accepted = $false
                Reason = $script:firstRunPackageReason; Message = 'SYNTHETIC verified base with unusable required package.'; Probe = $null
                ObservedVersion = '3.15.0'; IdentityVerified = $true
            })
            return $null
        }
        $failure = Get-FirstRunCaughtFailure {
            Resolve-WinBookSplitRuntime -ApplicationRoot $app -PythonPath 'C:\SYNTHETIC explicit Python\python.exe'
        }
        $failure.Data['Code'] | Should -BeExactly 'setup_required'
        Mock Find-WinBookSplitSetupBase { return [pscustomobject]@{ Path = $PythonPath } }
        $summary = Get-WinBookSplitDependencySummary -Failure $failure -ApplicationRoot $app -PythonPath 'C:\SYNTHETIC explicit Python\python.exe'
        $summary.ActionIsCommand | Should -BeTrue
        $summary.Action.Contains(" -PythonPath 'C:\SYNTHETIC explicit Python\python.exe'") | Should -BeTrue
        [IO.Directory]::Exists((Join-Path $app '.venv')) | Should -BeFalse
    }

    It 'forwards literal document exclusions during read-only setup and makes no environment changes' {
        $app = Join-Path $TestDrive ('read-only setup [' + [char]0x00f6 + '] & app')
        $engine = Join-Path $app 'engine'
        $null = New-Item -ItemType Directory -Path $engine
        Copy-Item -LiteralPath (Join-Path $RepositoryRoot 'engine\WinBookSplit.Compatibility.json') -Destination $engine
        $script:firstRunCheckDocument = Join-Path $TestDrive ('books [' + [char]0x00e4 + '] & literal\original.pdf')
        $script:firstRunCheckBase = 'C:\SYNTHETIC approved base\python.exe'
        Mock Resolve-WinBookSplitAutomaticPython {
            if ($InterpreterOnly) { return [pscustomobject]@{ Path = $script:firstRunCheckBase; Version = '3.15.0'; Source = 'manager' } }
            $Attempts.Add([pscustomobject]@{
                Path = $script:firstRunCheckBase; Source = 'manager'; Stage = 'package'; Accepted = $false
                Reason = 'missing_pypdf'; Message = 'SYNTHETIC verified base with missing package.'; Probe = $null
                ObservedVersion = '3.15.0'; IdentityVerified = $true
            })
            return $null
        }
        Mock Invoke-WinBookSplitSetupProcess { throw 'Read-only setup must never run venv or pip.' }
        Mock Read-Host { throw 'Read-only setup must never prompt.' }
        $before = @((Get-ChildItem -LiteralPath $app -Recurse -Force).FullName) -join "`n"
        $result = Invoke-WinBookSplitSetup -ApplicationRoot $app -DocumentPath $script:firstRunCheckDocument -Check -NonInteractive
        $result.code | Should -BeExactly 'setup_required'
        $result.exit_code | Should -Be 3
        $result.changed | Should -BeFalse
        (@((Get-ChildItem -LiteralPath $app -Recurse -Force).FullName) -join "`n") | Should -BeExactly $before
        Should-Invoke Resolve-WinBookSplitAutomaticPython -Times 2 -Exactly -ParameterFilter { $DocumentPath -ceq $script:firstRunCheckDocument }
        Should-Invoke Resolve-WinBookSplitAutomaticPython -Times 1 -Exactly -ParameterFilter {
            $InterpreterOnly -and $DocumentPath -ceq $script:firstRunCheckDocument -and
            $ExcludedDirectories -contains (Join-Path $app '.venv') -and $ExcludedDirectoryPrefixes.Count -eq 2
        }
        Should-Invoke Invoke-WinBookSplitSetupProcess -Times 0 -Exactly
        Should-Invoke Read-Host -Times 0 -Exactly
        [IO.Directory]::Exists((Join-Path $app '.venv')) | Should -BeFalse
    }
}

Describe 'Consent and resume application components (SYNTHETIC runtime/setup seams)' -Tag 'FUNCTIONAL-FIRST-RUN' {
    BeforeEach {
        # These deliberately replace discovery, setup and console input. They
        # prove the production coordinator's decisions, not native Store/BAT setup.
        # No .venv, Python, package, book or output is prepared by these cases.
        $script:firstRunApp = Join-Path $TestDrive ('application [' + [char]0x00f6 + '] & copy')
        $script:firstRunDocument = Join-Path $TestDrive ('books [' + [char]0x00e4 + '] & literal\original.pdf')
        $script:firstRunExplicit = 'C:\SYNTHETIC Python [base] & trusted\python.exe'
        $script:firstRunReady = [pscustomobject]@{
            Path = (Join-Path $script:firstRunApp '.venv\Scripts\python.exe')
            Version = '3.15.0'; Source = 'application_venv'; PypdfVersion = '6.21.4'
        }
        $script:firstRunInitialFailure = New-FirstRunDependencyFailure
        $script:firstRunSetup = New-FirstRunSetupResult
        $script:firstRunResolves = 0
        $script:firstRunSequence = [Collections.Generic.List[string]]::new()
        Mock Resolve-WinBookSplitRuntime {
            $script:firstRunSequence.Add('resolve')
            $script:firstRunResolves++
            if ($script:firstRunResolves -eq 1) { throw $script:firstRunInitialFailure }
            return $script:firstRunReady
        }
        Mock Read-WinBookSplitSessionLine {
            $script:firstRunSequence.Add('consent')
            return 'Y'
        }
        Mock Invoke-WinBookSplitSetup {
            $script:firstRunSequence.Add('setup')
            return $script:firstRunSetup
        }
    }

    It 'sets up once after consent, re-resolves the app environment and keeps source/root literal' {
        $ready = Resolve-WinBookSplitProcessingRuntime -ApplicationRoot $script:firstRunApp -DocumentPath $script:firstRunDocument
        $ready.Path | Should -BeExactly $script:firstRunReady.Path
        $ready.Source | Should -BeExactly 'application_venv'
        ($script:firstRunSequence -join ',') | Should -BeExactly 'resolve,consent,setup,resolve'
        Should-Invoke Read-WinBookSplitSessionLine -Times 1 -Exactly -ParameterFilter {
            $null -eq $Session -and $Prompt -ceq 'First-time setup is required. Download and install the PDF component into this copy of WinBookSplit now? [Y/N]'
        }
        Should-Invoke Invoke-WinBookSplitSetup -Times 1 -Exactly -ParameterFilter {
            $ApplicationRoot -ceq $script:firstRunApp -and $DocumentPath -ceq $script:firstRunDocument -and
            [string]::IsNullOrEmpty($PythonPath) -and $NonInteractive -and -not $Repair -and -not $Check
        }
        Should-Invoke Resolve-WinBookSplitRuntime -Times 2 -Exactly -ParameterFilter {
            $ApplicationRoot -ceq $script:firstRunApp -and $DocumentPath -ceq $script:firstRunDocument -and
            [string]::IsNullOrEmpty($PythonPath)
        }
        [IO.Directory]::Exists((Join-Path $script:firstRunApp '.venv')) | Should -BeFalse
    }

    It 'passes an original explicit base only to setup, then selects the newly prepared application environment' {
        $ready = Resolve-WinBookSplitProcessingRuntime -ApplicationRoot $script:firstRunApp -DocumentPath $script:firstRunDocument -PythonPath $script:firstRunExplicit
        $ready.Path | Should -BeExactly $script:firstRunReady.Path
        Should-Invoke Resolve-WinBookSplitRuntime -Times 1 -Exactly -ParameterFilter { $PythonPath -ceq $script:firstRunExplicit }
        Should-Invoke Invoke-WinBookSplitSetup -Times 1 -Exactly -ParameterFilter {
            $PythonPath -ceq $script:firstRunExplicit -and $DocumentPath -ceq $script:firstRunDocument -and $NonInteractive
        }
        Should-Invoke Resolve-WinBookSplitRuntime -Times 1 -Exactly -ParameterFilter {
            [string]::IsNullOrEmpty($PythonPath) -and $DocumentPath -ceq $script:firstRunDocument
        }
    }

    It 'refuses a healthy external fallback when the prepared app environment disappears' {
        $script:firstRunReady.Source = 'PATH'
        $script:firstRunReady.Path = 'C:\SYNTHETIC unrelated global Python\python.exe'
        $failure = Get-FirstRunCaughtFailure {
            Resolve-WinBookSplitProcessingRuntime -ApplicationRoot $script:firstRunApp -DocumentPath $script:firstRunDocument
        }
        $failure.Data['Code'] | Should -BeExactly 'setup_failed'
        $failure.Data['SetupResult'] | Should -Be $script:firstRunSetup
        Should-Invoke Resolve-WinBookSplitRuntime -Times 2 -Exactly
        Should-Invoke Invoke-WinBookSplitSetup -Times 1 -Exactly
        Should-Invoke Read-WinBookSplitSessionLine -Times 1 -Exactly
    }

    It 'repairs an originally explicit path inside this app environment without using it as a setup base' {
        $script:firstRunInitialFailure = New-FirstRunDependencyFailure -Code 'runtime_invalid' -Reason 'missing_pypdf' -Source 'explicit'
        $explicitAppPython = Join-Path $script:firstRunApp '.venv\Scripts\python.exe'
        $ready = Resolve-WinBookSplitProcessingRuntime -ApplicationRoot $script:firstRunApp -DocumentPath $script:firstRunDocument -PythonPath $explicitAppPython
        $ready.Path | Should -BeExactly $script:firstRunReady.Path
        Should-Invoke Read-WinBookSplitSessionLine -Times 1 -Exactly -ParameterFilter {
            $Prompt -match '(?i)repair' -and $Prompt -match '(?i)preserv|backup'
        }
        Should-Invoke Invoke-WinBookSplitSetup -Times 1 -Exactly -ParameterFilter {
            $Repair -and $NonInteractive -and [string]::IsNullOrEmpty($PythonPath)
        }
        Should-Invoke Resolve-WinBookSplitRuntime -Times 1 -Exactly -ParameterFilter { $PythonPath -ceq $explicitAppPython }
        Should-Invoke Resolve-WinBookSplitRuntime -Times 1 -Exactly -ParameterFilter { [string]::IsNullOrEmpty($PythonPath) }
    }

    It 'leaves a healthy repeated launch free of consent, setup and downloads' {
        Mock Resolve-WinBookSplitRuntime { return $script:firstRunReady }
        $ready = Resolve-WinBookSplitProcessingRuntime -ApplicationRoot $script:firstRunApp -DocumentPath $script:firstRunDocument
        $ready.Path | Should -BeExactly $script:firstRunReady.Path
        Should-Invoke Resolve-WinBookSplitRuntime -Times 1 -Exactly
        Should-Invoke Read-WinBookSplitSessionLine -Times 0 -Exactly
        Should-Invoke Invoke-WinBookSplitSetup -Times 0 -Exactly
    }

    It 'cancels without setup on <Name>' -ForEach @(
        @{ Name = 'No'; Answer = 'N' }, @{ Name = 'EOF'; Answer = $null },
        @{ Name = 'blank input'; Answer = '' }, @{ Name = 'explicit cancel'; Answer = 'C' }
    ) {
        $script:firstRunAnswer = $Answer
        Mock Read-WinBookSplitSessionLine { return $script:firstRunAnswer }
        $failure = Get-FirstRunCaughtFailure {
            Resolve-WinBookSplitProcessingRuntime -ApplicationRoot $script:firstRunApp -DocumentPath $script:firstRunDocument
        }
        $failure.Data['Code'] | Should -BeExactly 'cancelled'
        Should-Invoke Read-WinBookSplitSessionLine -Times 1 -Exactly
        Should-Invoke Invoke-WinBookSplitSetup -Times 0 -Exactly
        Should-Invoke Resolve-WinBookSplitRuntime -Times 1 -Exactly
    }

    It 'preserves console cancellation before setup' {
        Mock Read-WinBookSplitSessionLine { throw (New-WinBookSplitFailure 'processor_cancelled' 'SYNTHETIC console cancellation.') }
        $failure = Get-FirstRunCaughtFailure {
            Resolve-WinBookSplitProcessingRuntime -ApplicationRoot $script:firstRunApp -DocumentPath $script:firstRunDocument
        }
        $failure.Data['Code'] | Should -BeExactly 'processor_cancelled'
        Should-Invoke Invoke-WinBookSplitSetup -Times 0 -Exactly
        Should-Invoke Resolve-WinBookSplitRuntime -Times 1 -Exactly
    }

    It 'does not prompt or install in <Name>' -ForEach @(
        @{ Name = 'noninteractive processing'; NonInteractiveValue = $true; PreviewValue = $false },
        @{ Name = 'preview'; NonInteractiveValue = $false; PreviewValue = $true },
        @{ Name = 'noninteractive preview'; NonInteractiveValue = $true; PreviewValue = $true }
    ) {
        $failure = Get-FirstRunCaughtFailure {
            Resolve-WinBookSplitProcessingRuntime -ApplicationRoot $script:firstRunApp -DocumentPath $script:firstRunDocument -NonInteractive:$NonInteractiveValue -Preview:$PreviewValue
        }
        $failure.Data['Code'] | Should -BeExactly 'setup_required'
        Should-Invoke Read-WinBookSplitSessionLine -Times 0 -Exactly
        Should-Invoke Invoke-WinBookSplitSetup -Times 0 -Exactly
        Should-Invoke Resolve-WinBookSplitRuntime -Times 1 -Exactly
    }

    It 'does not offer setup for <Name>' -ForEach @(
        @{ Name = 'missing interpreter'; Code = 'runtime_not_found'; Reason = 'missing_python'; Source = 'PATH' },
        @{ Name = 'unsupported Python'; Code = 'runtime_invalid'; Reason = 'python_security_floor'; Source = 'explicit' },
        @{ Name = 'unproved process shutdown'; Code = 'processor_failed'; Reason = 'processor_failed'; Source = 'manager' }
    ) {
        $script:firstRunInitialFailure = New-FirstRunDependencyFailure -Code $Code -Reason $Reason -Source $Source
        $failure = Get-FirstRunCaughtFailure {
            Resolve-WinBookSplitProcessingRuntime -ApplicationRoot $script:firstRunApp -DocumentPath $script:firstRunDocument
        }
        $failure.Data['Code'] | Should -BeExactly $Code
        Should-Invoke Read-WinBookSplitSessionLine -Times 0 -Exactly
        Should-Invoke Invoke-WinBookSplitSetup -Times 0 -Exactly
        Should-Invoke Resolve-WinBookSplitRuntime -Times 1 -Exactly
    }

    It 'never offers repair when the authoritative failure is malformed shipped policy' {
        $script:firstRunInitialFailure = New-FirstRunDependencyFailure -Code 'runtime_invalid' -Reason 'invalid_application_venv' -Source 'application_venv' -Verified $false
        $script:firstRunInitialFailure.Data['Reason'] = 'malformed_policy'
        $failure = Get-FirstRunCaughtFailure {
            Resolve-WinBookSplitProcessingRuntime -ApplicationRoot $script:firstRunApp -DocumentPath $script:firstRunDocument
        }
        $failure.Data['Code'] | Should -BeExactly 'runtime_invalid'
        $failure.Data['Reason'] | Should -BeExactly 'malformed_policy'
        Should-Invoke Read-WinBookSplitSessionLine -Times 0 -Exactly
        Should-Invoke Invoke-WinBookSplitSetup -Times 0 -Exactly
    }

    It 'never offers setup when dependency process shutdown is uncertain' {
        $script:firstRunInitialFailure.Data['FatalDiscovery'] = $true
        $failure = Get-FirstRunCaughtFailure {
            Resolve-WinBookSplitProcessingRuntime -ApplicationRoot $script:firstRunApp -DocumentPath $script:firstRunDocument
        }
        $failure.Data['FatalDiscovery'] | Should -BeTrue
        Should-Invoke Read-WinBookSplitSessionLine -Times 0 -Exactly
        Should-Invoke Invoke-WinBookSplitSetup -Times 0 -Exactly
    }

    It 'requires a separate explicit repair offer for an unhealthy authoritative application environment' {
        $script:firstRunInitialFailure = New-FirstRunDependencyFailure -Code 'runtime_invalid' -Reason 'invalid_application_venv' -Source 'application_venv' -Verified $false
        $ready = Resolve-WinBookSplitProcessingRuntime -ApplicationRoot $script:firstRunApp -DocumentPath $script:firstRunDocument
        $ready.Path | Should -BeExactly $script:firstRunReady.Path
        Should-Invoke Read-WinBookSplitSessionLine -Times 1 -Exactly -ParameterFilter {
            $Prompt -match '(?i)repair' -and $Prompt -match '(?i)preserv|backup' -and $Prompt -match '\[Y/N\]'
        }
        Should-Invoke Invoke-WinBookSplitSetup -Times 1 -Exactly -ParameterFilter { $Repair -and $NonInteractive }
        Should-Invoke Resolve-WinBookSplitRuntime -Times 2 -Exactly
    }

    It 'preserves an unhealthy application environment when repair is declined' {
        $script:firstRunInitialFailure = New-FirstRunDependencyFailure -Code 'runtime_invalid' -Reason 'invalid_application_venv' -Source 'application_venv' -Verified $false
        Mock Read-WinBookSplitSessionLine { return 'N' }
        $failure = Get-FirstRunCaughtFailure {
            Resolve-WinBookSplitProcessingRuntime -ApplicationRoot $script:firstRunApp -DocumentPath $script:firstRunDocument
        }
        $failure.Data['Code'] | Should -BeExactly 'cancelled'
        Should-Invoke Invoke-WinBookSplitSetup -Times 0 -Exactly
        Should-Invoke Resolve-WinBookSplitRuntime -Times 1 -Exactly
    }

    It 'stops once and preserves the returned failure details for <Name>' -ForEach @(
        @{ Name = 'pip/network failure'; Status = 'failed'; Code = 'setup_failed'; ExitCode = 6 },
        @{ Name = 'cancelled setup'; Status = 'cancelled'; Code = 'setup_cancelled'; ExitCode = 130 },
        @{ Name = 'timed out setup'; Status = 'timeout'; Code = 'setup_timeout'; ExitCode = 130 },
        @{ Name = 'backup restoration failure'; Status = 'failed'; Code = 'setup_rollback_failed'; ExitCode = 6 }
    ) {
        $script:firstRunSetup = New-FirstRunSetupResult -Status $Status -Code $Code -ExitCode $ExitCode
        $script:firstRunSetup.message = 'SYNTHETIC actual pip stderr: offline proxy refused the request; retained artifacts.'
        $failure = Get-FirstRunCaughtFailure {
            Resolve-WinBookSplitProcessingRuntime -ApplicationRoot $script:firstRunApp -DocumentPath $script:firstRunDocument
        }
        $failure.Data['Code'] | Should -BeExactly $Code
        $failure.Message | Should -Match 'offline proxy refused'
        $failure.Data['SetupResult'] | Should -Be $script:firstRunSetup
        Should-Invoke Invoke-WinBookSplitSetup -Times 1 -Exactly
        Should-Invoke Resolve-WinBookSplitRuntime -Times 1 -Exactly
        Should-Invoke Read-WinBookSplitSessionLine -Times 1 -Exactly
    }

    It 'rejects a forged setup success with <Name> and never processes the document' -ForEach @(
        @{ Name = 'wrong protocol'; Field = 'protocol'; Value = 'unrelated.setup' },
        @{ Name = 'wrong schema'; Field = 'schema_version'; Value = 2 },
        @{ Name = 'nonzero exit'; Field = 'exit_code'; Value = 6 },
        @{ Name = 'failure code'; Field = 'code'; Value = 'setup_failed' },
        @{ Name = 'failure status'; Field = 'status'; Value = 'failed' }
    ) {
        $script:firstRunSetup.$Field = $Value
        $failure = Get-FirstRunCaughtFailure {
            Resolve-WinBookSplitProcessingRuntime -ApplicationRoot $script:firstRunApp -DocumentPath $script:firstRunDocument
        }
        $failure | Should -Not -BeNullOrEmpty
        Should-Invoke Invoke-WinBookSplitSetup -Times 1 -Exactly
        Should-Invoke Resolve-WinBookSplitRuntime -Times 1 -Exactly
    }

    It 'allows no second setup cycle when re-resolution still reports missing pypdf' {
        Mock Resolve-WinBookSplitRuntime { throw $script:firstRunInitialFailure }
        $failure = Get-FirstRunCaughtFailure {
            Resolve-WinBookSplitProcessingRuntime -ApplicationRoot $script:firstRunApp -DocumentPath $script:firstRunDocument
        }
        $failure.Data['Code'] | Should -BeExactly 'setup_required'
        Should-Invoke Resolve-WinBookSplitRuntime -Times 2 -Exactly
        Should-Invoke Invoke-WinBookSplitSetup -Times 1 -Exactly
        Should-Invoke Read-WinBookSplitSessionLine -Times 1 -Exactly
    }
}
