param([string]$RepositoryRoot, [string]$WorkRoot, [string]$TestPython, [string]$Wheelhouse)

BeforeAll {
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Diagnostics.ps1')
    . (Join-Path $RepositoryRoot 'engine\WinBookSplit.Setup.ps1')
    $tokens=$null;$parseErrors=$null
    $setupAst=[Management.Automation.Language.Parser]::ParseFile((Join-Path $RepositoryRoot 'Setup-WinBookSplit.ps1'),[ref]$tokens,[ref]$parseErrors)
    if(@($parseErrors).Count -ne 0){throw 'The trusted setup frontend has syntax errors.'}
    $hostFunctions=@($setupAst.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq 'Get-WinBookSplitSetupHostObservation'},$true))
    if($hostFunctions.Count -ne 1){throw 'Expected one trusted setup host observation function.'}
    . ([scriptblock]::Create($hostFunctions[0].Extent.Text))
    if (-not [IO.File]::Exists($TestPython)) { throw 'Setup tests require a concrete accepted TestPython.' }
    if (-not [IO.Directory]::Exists($Wheelhouse)) { throw 'Setup tests require an explicit reviewed Wheelhouse; set WBS_TEST_WHEELHOUSE.' }
    $script:setupPolicy=Get-WinBookSplitCompatibilityPolicy
    $script:setupWheel=Join-Path $Wheelhouse ('pypdf-'+$script:setupPolicy.pypdf.version+'-py3-none-any.whl')
    if (-not [IO.File]::Exists($script:setupWheel)) { throw 'The reviewed wheel fixture is absent.' }
    $script:setupNativePhase=(Get-Command Invoke-WinBookSplitSetupProcess).ScriptBlock
    $script:setupNativeCandidate=(Get-Command Test-WinBookSplitPythonCandidate).ScriptBlock
    $script:setupNativePolicy=(Get-Command Get-WinBookSplitCompatibilityPolicy).ScriptBlock
    $script:setupExpectedVersion=(Resolve-WinBookSplitRuntime -ApplicationRoot $RepositoryRoot -PythonPath $TestPython -InterpreterOnly).Version
    function Get-SetupFileDigest {
        param([string]$Path)
        $digest=[Security.Cryptography.SHA256]::Create()
        try { return ([BitConverter]::ToString($digest.ComputeHash([IO.File]::ReadAllBytes($Path)))).Replace('-','').ToLowerInvariant() }
        finally { $digest.Dispose() }
    }
    (Get-SetupFileDigest $script:setupWheel) | Should -BeExactly $script:setupPolicy.pypdf.wheel_sha256
    function New-SetupApplication {
        $app=Join-Path $TestDrive ('a'+[guid]::NewGuid().ToString('N').Substring(0,12))
        $engine=Join-Path $app 'engine'
        $null=New-Item -ItemType Directory -Path $engine
        foreach ($name in @('Setup-WinBookSplit.ps1','requirements.txt')) {
            Copy-Item -LiteralPath (Join-Path $RepositoryRoot $name) -Destination $app
        }
        foreach ($name in @('WinBookSplit.Setup.ps1','WinBookSplit.Runtime.ps1','WinBookSplit.Process.ps1','WinBookSplit.Diagnostics.ps1',
                'WinBookSplit.Support.ps1','WinBookSplit.Logging.ps1','WinBookSplit.Outcomes.json','WinBookSplit.Compatibility.ps1','WinBookSplit.Compatibility.json','winbooksplit_compatibility.py')) {
            Copy-Item -LiteralPath (Join-Path $RepositoryRoot ('engine\'+$name)) -Destination $engine
        }
        [IO.File]::WriteAllText((Join-Path $app 'foreign-neighbor.txt'),'Authored unrelated file must survive setup.')
        return $app
    }
    function New-BrokenSetupApplication {
        $app=New-SetupApplication
        $null=New-Item -ItemType Directory -Path (Join-Path $app '.venv\Scripts')
        [IO.File]::WriteAllText((Join-Path $app '.venv\preserve-old.txt'),'Authored old environment to restore unchanged.')
        return $app
    }
    function Get-SetupSnapshot {
        param([string]$Path)
        $records=@(Get-ChildItem -LiteralPath $Path -File -Recurse -Force | ForEach-Object {
            @{relative=$_.FullName.Substring($Path.Length);sha256=(Get-SetupFileDigest $_.FullName);ticks=$_.LastWriteTimeUtc.Ticks}
        } | Sort-Object -Property relative)
        return (ConvertTo-Json -InputObject $records -Depth 5 -Compress)
    }
    function Get-SetupEffectiveAcl {
        param($Acl)
        $rules=@($Acl.Access | ForEach-Object {
            @{identity=$_.IdentityReference.Value;rights=[string]$_.FileSystemRights;type=[string]$_.AccessControlType;
                inherited=$_.IsInherited;inheritance=[string]$_.InheritanceFlags;propagation=[string]$_.PropagationFlags}
        } | Sort-Object -Property identity,rights,type,inherited,inheritance,propagation)
        return (ConvertTo-Json -InputObject @{owner=$Acl.Owner;group=$Acl.Group;protected=$Acl.AreAccessRulesProtected;rules=$rules} -Depth 6 -Compress)
    }
    function Assert-SetupNoMutation {
        param([string]$App, [string]$Before)
        Get-SetupSnapshot $App | Should -BeExactly $Before
        [IO.Directory]::Exists((Join-Path $App '.venv')) | Should -BeFalse
        [IO.File]::Exists((Join-Path $App '.WinBookSplit-setup.lock')) | Should -BeFalse
    }
    function Write-SetupObservation {
        param([string]$Name, $Result)
        if (-not [IO.Directory]::Exists($WorkRoot)) { throw 'Setup observations require the owned WorkRoot.' }
        $path=Join-Path $WorkRoot ('setup-'+$Name+'-'+[guid]::NewGuid().ToString('N')+'.json')
        [IO.File]::WriteAllText($path,($Result | ConvertTo-Json -Depth 16),(New-Object Text.UTF8Encoding($false)))
    }
    $script:healthySetupApp=New-SetupApplication
}

Describe 'V102 local setup host and policy observations' -Tag 'V102-AC-041','V102-AC-042' {
    It 'observes the actual current host and policy scopes without changing them' {
        $before=@(Get-ExecutionPolicy -List | ForEach-Object {@{scope=$_.Scope.ToString();policy=$_.ExecutionPolicy.ToString()}}) | ConvertTo-Json -Compress
        $current=[Diagnostics.Process]::GetCurrentProcess()
        try{$expectedPath=$current.MainModule.FileName}finally{$current.Dispose()}
        $record=Get-WinBookSplitSetupHostObservation
        $record.executable_status | Should -BeExactly 'observed'
        $record.executable | Should -BeExactly $expectedPath
        $record.version | Should -BeExactly $PSVersionTable.PSVersion.ToString()
        $record.edition | Should -BeExactly $PSVersionTable.PSEdition
        $record.execution_policy.status | Should -BeExactly 'observed'
        $record.execution_policy.effective | Should -BeExactly (Get-ExecutionPolicy).ToString()
        @($record.execution_policy.scopes).Count | Should -Be 5
        foreach($scope in @('MachinePolicy','UserPolicy','Process','CurrentUser','LocalMachine')){
            @($record.execution_policy.scopes | Where-Object {$_.scope -ceq $scope})[0].policy | Should -BeExactly (Get-ExecutionPolicy -Scope $scope).ToString()
        }
        (@(Get-ExecutionPolicy -List | ForEach-Object {@{scope=$_.Scope.ToString();policy=$_.ExecutionPolicy.ToString()}}) | ConvertTo-Json -Compress) | Should -BeExactly $before
    }
    It 'reports a diagnostic query failure as unavailable with no partial or guessed scopes' {
        # Controlled query failure only; this does not enforce or change policy.
        Mock Get-ExecutionPolicy {throw 'SYNTHETIC inaccessible policy query'}
        $record=Get-WinBookSplitSetupHostObservation
        $record.execution_policy.status | Should -BeExactly 'unavailable'
        $record.execution_policy.reason_code | Should -BeExactly 'query_failed'
        $record.execution_policy.effective | Should -BeNullOrEmpty
        @($record.execution_policy.scopes).Count | Should -Be 0
        $record.executable_status | Should -BeExactly 'observed'
    }
    It 'keeps simulated organization effective signing distinct from child Process Bypass' {
        # SYNTHETIC query records, not native Group Policy enforcement evidence.
        Mock Get-ExecutionPolicy {
            if($List){return @('MachinePolicy','UserPolicy','Process','CurrentUser','LocalMachine') | ForEach-Object {
                [pscustomobject]@{Scope=$_;ExecutionPolicy=$(if($_ -eq 'MachinePolicy'){'AllSigned'}elseif($_ -eq 'Process'){'Bypass'}else{'Undefined'})}
            }}
            return 'AllSigned'
        }
        Mock Set-ExecutionPolicy {throw 'A diagnostic must never change policy.'}
        $record=Get-WinBookSplitSetupHostObservation
        $record.execution_policy.status | Should -BeExactly 'observed'
        $record.execution_policy.effective | Should -BeExactly 'AllSigned'
        @($record.execution_policy.scopes | Where-Object {$_.scope -eq 'Process'})[0].policy | Should -BeExactly 'Bypass'
        Should-Invoke Set-ExecutionPolicy -Times 0 -Exactly
    }
    It 'retains a real frontend import failure and nonzero exit with an actual local host receipt' {
        $app=New-SetupApplication
        Remove-Item -LiteralPath (Join-Path $app 'engine\WinBookSplit.Setup.ps1')
        $before=Get-SetupSnapshot $app
        $hostPath=(Get-WinBookSplitSetupHostObservation).executable
        $probe=Invoke-WinBookSplitProcess -Path $hostPath -Arguments @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $app 'Setup-WinBookSplit.ps1'),'-NonInteractive','-Check') -WorkingDirectory $app -TimeoutSeconds 60
        $probe.ExitCode | Should -Be 6
        $probe.ParentStopped | Should -BeTrue
        $probe.DescendantsStopped | Should -BeTrue
        $probe.StreamsComplete | Should -BeTrue
        $frames=@([regex]::Matches($probe.Stdout,'(?m)^\[SETUP\] ([^\r\n]+)'))
        $frames.Count | Should -Be 1
        $result=$frames[0].Groups[1].Value | ConvertFrom-Json
        $result.status | Should -BeExactly 'failed'
        $result.code | Should -BeExactly 'setup_failed'
        $result.exit_code | Should -Be 6
        $result.changed | Should -BeFalse
        $result.host.executable | Should -BeExactly $hostPath
        $result.host.version | Should -BeExactly $PSVersionTable.PSVersion.ToString()
        $result.host.execution_policy.status | Should -BeExactly 'observed'
        Get-SetupSnapshot $app | Should -BeExactly $before
        Write-SetupObservation 'frontend-import-failure' ([pscustomobject]@{probe=$probe;result=$result})
    }
    It 'observes inbox child policy with the actual parent host module search path inherited' {
        $app=New-SetupApplication
        Remove-Item -LiteralPath (Join-Path $app 'engine\WinBookSplit.Setup.ps1')
        $hostPath=Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::System)) 'WindowsPowerShell\v1.0\powershell.exe'
        $previousModules=$env:PSModulePath
        try {
            # The PS7 run passes actual PS7 builtin modules to actual inbox5.1.
            # This exercises module incompatibility, not simulated policy.
            $env:PSModulePath=Join-Path $PSHOME 'Modules'
            $probe=Invoke-WinBookSplitProcess -Path $hostPath -Arguments @('-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $app 'Setup-WinBookSplit.ps1'),'-NonInteractive','-Check') -WorkingDirectory $app -TimeoutSeconds 60
        } finally {$env:PSModulePath=$previousModules}
        $probe.ExitCode | Should -Be 6
        $probe.ParentStopped | Should -BeTrue
        $probe.DescendantsStopped | Should -BeTrue
        $probe.StreamsComplete | Should -BeTrue
        $frames=@([regex]::Matches($probe.Stdout,'(?m)^\[SETUP\] ([^\r\n]+)'))
        $frames.Count | Should -Be 1
        $result=$frames[0].Groups[1].Value | ConvertFrom-Json
        Write-SetupObservation 'frontend-inbox-parent-module-path' ([pscustomobject]@{probe=$probe;result=$result;parent_version=$PSVersionTable.PSVersion.ToString()})
        $result.host.executable | Should -BeExactly $hostPath
        $result.host.version | Should -Match '\A5\.1\.'
        $result.host.execution_policy.status | Should -BeExactly 'observed'
        $result.host.execution_policy.effective | Should -BeExactly 'Bypass'
        @($result.host.execution_policy.scopes | Where-Object {$_.scope -eq 'Process'})[0].policy | Should -BeExactly 'Bypass'
        $env:PSModulePath | Should -BeExactly $previousModules
    }
}

Describe 'V101 explicit application-local setup with the real reviewed offline wheel' -Tag 'V101-AC-025','V101-AC-026','V101-AC-029','V101-AC-030','V101-AC-031' {
    It 'creates the clean environment at its final path and verifies the locked package and concrete pip arguments' {
        $result=Invoke-WinBookSplitSetup -ApplicationRoot $script:healthySetupApp -PythonPath $TestPython -Wheelhouse $Wheelhouse -NonInteractive
        Write-SetupObservation 'offline-create' $result
        $result.code | Should -BeExactly 'setup_complete'
        $result.exit_code | Should -Be 0
        $result.changed | Should -BeTrue
        $venv=Join-Path $script:healthySetupApp '.venv'
        $result.environment_path | Should -BeExactly $venv
        $selected=Resolve-WinBookSplitRuntime -ApplicationRoot $script:healthySetupApp
        $selected.Path | Should -BeExactly (Join-Path $venv 'Scripts\python.exe')
        $selected.Version | Should -BeExactly $script:setupExpectedVersion
        $selected.PypdfVersion | Should -BeExactly '6.20.0'
        $selected.Source | Should -BeExactly 'application_venv'
        $install=@($result.operations | Where-Object { $_.phase -eq 'pip_install' })
        $install.Count | Should -Be 1
        $install[0].path | Should -BeExactly $selected.Path
        foreach ($argument in @('-I','-m','pip','--isolated','--no-input','--require-hashes','--only-binary=:all:','--no-deps','--no-index')) {
            @($install[0].arguments) | Should -Contain $argument
        }
        @($install[0].arguments) | Should -Contain $script:setupWheel
        @($install[0].arguments) | Should -Not -Contain $Wheelhouse
        $install[0].probe.JobAssigned | Should -BeTrue
        $install[0].probe.ParentStopped | Should -BeTrue
        $install[0].probe.DescendantsStopped | Should -BeTrue
        $install[0].probe.StreamsComplete | Should -BeTrue
        [IO.File]::ReadAllText((Join-Path $script:healthySetupApp 'foreign-neighbor.txt')) | Should -BeExactly 'Authored unrelated file must survive setup.'
    }
    It 'checks and reuses a healthy environment without package operations or changed file bytes and timestamps' {
        $before=Get-SetupSnapshot $script:healthySetupApp
        Mock Invoke-WinBookSplitSetupProcess { throw 'A healthy environment must not create or install anything.' }
        foreach ($check in @($true,$false)) {
            $result=Invoke-WinBookSplitSetup -ApplicationRoot $script:healthySetupApp -NonInteractive -Check:$check -PythonPath (Join-Path $TestDrive 'unused-missing-python.exe') -Wheelhouse (Join-Path $TestDrive 'unused-missing-wheelhouse')
            Write-SetupObservation 'healthy-unused-prerequisites' $result
            $result.code | Should -BeExactly 'setup_ready'
            $result.exit_code | Should -Be 0
            $result.changed | Should -BeFalse
            @($result.operations).Count | Should -Be 0
            Get-SetupSnapshot $script:healthySetupApp | Should -BeExactly $before
        }
        Should-Invoke Invoke-WinBookSplitSetupProcess -Times 0 -Exactly
    }
    It 'uses Check as a read-only negative diagnosis without creating an environment or setup lock' {
        $app=New-SetupApplication
        $before=Get-SetupSnapshot $app
        Mock Invoke-WinBookSplitSetupProcess { throw 'Check must never create or install anything.' }
        $result=Invoke-WinBookSplitSetup -ApplicationRoot $app -PythonPath (Join-Path $TestDrive 'missing-check-python.exe') -Check -NonInteractive
        $result.code | Should -BeExactly 'runtime_invalid'
        $result.exit_code | Should -Be 3
        $result.changed | Should -BeFalse
        Assert-SetupNoMutation $app $before
        Should-Invoke Invoke-WinBookSplitSetupProcess -Times 0 -Exactly
    }
}

Describe 'V101 damaged environment authority and ownership-limited rollback' -Tag 'V101-AC-027','V101-AC-028','V101-AC-032','V101-AC-033' {
    It 'preserves a broken existing environment by default and explains explicit Repair' {
        $app=New-BrokenSetupApplication
        $before=Get-SetupSnapshot $app
        $result=Invoke-WinBookSplitSetup -ApplicationRoot $app -PythonPath $TestPython -Wheelhouse $Wheelhouse -NonInteractive
        $result.code | Should -BeExactly 'setup_repair_required'
        $result.exit_code | Should -Be 3
        $result.changed | Should -BeFalse
        $result.message | Should -Match 'Repair'
        Get-SetupSnapshot $app | Should -BeExactly $before
    }
    It 'restores the old environment and preserves foreign neighbors after controlled native install failure' {
        $app=New-BrokenSetupApplication
        $before=Get-SetupSnapshot (Join-Path $app '.venv')
        Mock Invoke-WinBookSplitSetupProcess {
            if ($Phase -eq 'pip_install') {
                # SYNTHETIC failure command; its real owned native process exits17.
                return & $script:setupNativePhase -Path $TestPython -Arguments @('-I','-B','-c','raise SystemExit(17)') -ApplicationRoot $ApplicationRoot -Phase $Phase -Operations $Operations -TimeoutSeconds $TimeoutSeconds -CancellationToken $CancellationToken
            }
            return & $script:setupNativePhase -Path $Path -Arguments $Arguments -ApplicationRoot $ApplicationRoot -Phase $Phase -Operations $Operations -TimeoutSeconds $TimeoutSeconds -CancellationToken $CancellationToken -IsolatePipConfiguration:$IsolatePipConfiguration
        }
        $result=Invoke-WinBookSplitSetup -ApplicationRoot $app -PythonPath $TestPython -Wheelhouse $Wheelhouse -Repair -NonInteractive
        Write-SetupObservation 'controlled-native-install-failure' $result
        $result.exit_code | Should -Be 6
        $result.code | Should -BeExactly 'setup_failed'
        $result.message | Should -Match 'pip_install failed'
        $result.message | Should -Not -Match 'not serializable'
        $result.rollback_status | Should -BeExactly 'restored'
        Get-SetupSnapshot (Join-Path $app '.venv') | Should -BeExactly $before
        [IO.File]::ReadAllText((Join-Path $app 'foreign-neighbor.txt')) | Should -BeExactly 'Authored unrelated file must survive setup.'
        $failed=@($result.operations | Where-Object { $_.phase -eq 'pip_install' })
        $failed.Count | Should -Be 1
        $failed[0].probe.ExitCode | Should -Be 17
        $failed[0].probe.ParentStopped | Should -BeTrue
        $failed[0].probe.DescendantsStopped | Should -BeTrue
    }
    It 'proves native cancellation and preserves the old environment during explicit Repair' {
        $app=New-BrokenSetupApplication
        $before=Get-SetupSnapshot (Join-Path $app '.venv')
        $script:setupCancellation=New-Object Threading.CancellationTokenSource
        Mock Invoke-WinBookSplitSetupProcess {
            if ($Phase -eq 'pip_install') {
                # Controlled helper replaces only installation; cancellation is native.
                $script:setupCancellation.CancelAfter(150)
                return & $script:setupNativePhase -Path $TestPython -Arguments @('-I','-B','-c','import time;time.sleep(8)') -ApplicationRoot $ApplicationRoot -Phase $Phase -Operations $Operations -TimeoutSeconds 5 -CancellationToken $script:setupCancellation.Token
            }
            return & $script:setupNativePhase -Path $Path -Arguments $Arguments -ApplicationRoot $ApplicationRoot -Phase $Phase -Operations $Operations -TimeoutSeconds $TimeoutSeconds -CancellationToken $CancellationToken -IsolatePipConfiguration:$IsolatePipConfiguration
        }
        try {
            $result=Invoke-WinBookSplitSetup -ApplicationRoot $app -PythonPath $TestPython -Wheelhouse $Wheelhouse -Repair -NonInteractive -CancellationToken $script:setupCancellation.Token
            Write-SetupObservation 'controlled-native-cancellation' $result
            $result.code | Should -BeExactly 'setup_cancelled'
            $result.exit_code | Should -Be 130
            $result.rollback_status | Should -BeExactly 'restored'
            Get-SetupSnapshot (Join-Path $app '.venv') | Should -BeExactly $before
            $cancelled=@($result.operations | Where-Object { $_.phase -eq 'pip_install' })
            $cancelled[0].probe.Cancelled | Should -BeTrue
            $cancelled[0].probe.TimedOut | Should -BeFalse
            $cancelled[0].probe.ParentStopped | Should -BeTrue
            $cancelled[0].probe.DescendantsStopped | Should -BeTrue
        } finally { $script:setupCancellation.Dispose() }
    }
    It 'rejects an actual ordinary environment replacement during controlled preflight without adopting it' {
        $app=New-BrokenSetupApplication
        $original=Join-Path $app '.concurrent-original'
        $script:setupReplacementApp=$app
        $script:setupReplacementMoved=$false
        $before=Get-SetupSnapshot (Join-Path $app '.venv')
        Mock Test-WinBookSplitPythonCandidate {
            $candidate=& $script:setupNativeCandidate -Path $Path -Source $Source -ApplicationRoot $ApplicationRoot -Attempts $Attempts -TimeoutSeconds $TimeoutSeconds -InterpreterOnly:$InterpreterOnly
            if ($Source -eq 'application_venv' -and -not $script:setupReplacementMoved) {
                # SYNTHETIC scheduling gate; both directory identities and the replacement are native.
                $venv=Join-Path $script:setupReplacementApp '.venv'
                [IO.Directory]::Move($venv,(Join-Path $script:setupReplacementApp '.concurrent-original'))
                $null=[IO.Directory]::CreateDirectory($venv)
                [IO.File]::WriteAllText((Join-Path $venv 'foreign-replacement.txt'),'Authored concurrent replacement must survive.')
                $script:setupReplacementMoved=$true
            }
            return $candidate
        }
        $result=Invoke-WinBookSplitSetup -ApplicationRoot $app -PythonPath $TestPython -Wheelhouse $Wheelhouse -Repair -NonInteractive
        Write-SetupObservation 'concurrent-replacement' $result
        $script:setupReplacementMoved | Should -BeTrue
        $result.code | Should -BeExactly 'setup_busy'
        $result.changed | Should -BeFalse
        @($result.operations).Count | Should -Be 0
        Get-SetupSnapshot $original | Should -BeExactly $before
        [IO.File]::ReadAllText((Join-Path $app '.venv\foreign-replacement.txt')) | Should -BeExactly 'Authored concurrent replacement must survive.'
        [IO.Directory]::Exists((Join-Path $app '.venv\Scripts')) | Should -BeFalse
        @([IO.Directory]::EnumerateDirectories($app,'.WinBookSplit-*')).Count | Should -Be 0
        [IO.File]::Exists((Join-Path $app '.WinBookSplit-setup.lock')) | Should -BeFalse
    }
    It 'refuses an exclusive foreign setup lock without changing or deleting its bytes' {
        $app=New-SetupApplication
        $lock=Join-Path $app '.WinBookSplit-setup.lock'
        [IO.File]::WriteAllText($lock,'Authored held foreign lock.')
        $held=[IO.File]::Open($lock,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None)
        try {
            $result=Invoke-WinBookSplitSetup -ApplicationRoot $app -PythonPath $TestPython -Wheelhouse $Wheelhouse -NonInteractive
            $result.code | Should -BeExactly 'setup_busy'
            $result.changed | Should -BeFalse
            [IO.Directory]::Exists((Join-Path $app '.venv')) | Should -BeFalse
        } finally { $held.Dispose() }
        [IO.File]::ReadAllText($lock) | Should -BeExactly 'Authored held foreign lock.'
    }
    It 'rejects reparse environment destinations and leaves their owned target untouched' {
        $app=New-SetupApplication
        $target=Join-Path $TestDrive ('t'+[guid]::NewGuid().ToString('N').Substring(0,12))
        $null=New-Item -ItemType Directory -Path $target
        $sentinel=Join-Path $target 'preserve.txt'
        [IO.File]::WriteAllText($sentinel,'Authored reparse target must survive.')
        $venv=Join-Path $app '.venv'
        $null=New-Item -ItemType Junction -Path $venv -Target $target
        try {
            $result=Invoke-WinBookSplitSetup -ApplicationRoot $app -PythonPath $TestPython -Wheelhouse $Wheelhouse -Repair -NonInteractive
            $result.exit_code | Should -Not -Be 0
            $result.changed | Should -BeFalse
            [IO.File]::ReadAllText($sentinel) | Should -BeExactly 'Authored reparse target must survive.'
        } finally { [IO.Directory]::Delete($venv) }
        [IO.File]::ReadAllText($sentinel) | Should -BeExactly 'Authored reparse target must survive.'
    }
    It 'preserves held old files and foreign neighbors when Repair cannot finish' {
        $app=New-BrokenSetupApplication
        $oldFile=Join-Path $app '.venv\preserve-old.txt'
        $held=[IO.File]::Open($oldFile,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None)
        Mock Invoke-WinBookSplitSetupProcess {
            if ($Phase -eq 'pip_install') {
                return & $script:setupNativePhase -Path $TestPython -Arguments @('-I','-B','-c','raise SystemExit(17)') -ApplicationRoot $ApplicationRoot -Phase $Phase -Operations $Operations -TimeoutSeconds $TimeoutSeconds -CancellationToken $CancellationToken
            }
            return & $script:setupNativePhase -Path $Path -Arguments $Arguments -ApplicationRoot $ApplicationRoot -Phase $Phase -Operations $Operations -TimeoutSeconds $TimeoutSeconds -CancellationToken $CancellationToken -IsolatePipConfiguration:$IsolatePipConfiguration
        }
        try {
            $result=Invoke-WinBookSplitSetup -ApplicationRoot $app -PythonPath $TestPython -Wheelhouse $Wheelhouse -Repair -NonInteractive
            $result.exit_code | Should -Not -Be 0
            @('setup_failed','setup_rollback_failed','setup_busy') | Should -Contain $result.code
            [IO.File]::ReadAllText((Join-Path $app 'foreign-neighbor.txt')) | Should -BeExactly 'Authored unrelated file must survive setup.'
        } finally { $held.Dispose() }
        $preserved=$oldFile
        if (-not [IO.File]::Exists($preserved) -and $result.backup_path) { $preserved=Join-Path $result.backup_path 'preserve-old.txt' }
        [IO.File]::ReadAllText($preserved) | Should -BeExactly 'Authored old environment to restore unchanged.'
    }
    It 'fails an actually unwritable owned app root with actionable per-user guidance and no environment' {
        $app=New-SetupApplication
        $before=Get-SetupSnapshot $app
        $original=Get-Acl -LiteralPath $app
        $sddl=$original.Sddl
        $effectiveAcl=Get-SetupEffectiveAcl $original
        $restricted=Get-Acl -LiteralPath $app
        $identity=[Security.Principal.WindowsIdentity]::GetCurrent().User
        $rule=New-Object Security.AccessControl.FileSystemAccessRule($identity,[Security.AccessControl.FileSystemRights]::Write,
            [Security.AccessControl.InheritanceFlags]::None,[Security.AccessControl.PropagationFlags]::None,[Security.AccessControl.AccessControlType]::Deny)
        $restricted.AddAccessRule($rule)
        try {
            Set-Acl -LiteralPath $app -AclObject $restricted
            $result=Invoke-WinBookSplitSetup -ApplicationRoot $app -PythonPath $TestPython -Wheelhouse $Wheelhouse -NonInteractive
            $result.exit_code | Should -Not -Be 0
            $result.changed | Should -BeFalse
            $result.message | Should -Match 'unwrit|access|permission|per-user|extract'
            Assert-SetupNoMutation $app $before
        } finally {
            $restore=Get-Acl -LiteralPath $app
            $restore.SetSecurityDescriptorSddlForm($sddl)
            Set-Acl -LiteralPath $app -AclObject $restore
        }
        # Windows may normalize D: to D:AI while restoring identical effective ACEs.
        Get-SetupEffectiveAcl (Get-Acl -LiteralPath $app) | Should -BeExactly $effectiveAcl
    }
}

Describe 'V101 precise prerequisite and reviewed wheel failures' -Tag 'V101-AC-030','V101-AC-031','V101-AC-033' {
    It 'rejects preserved backup and incomplete directories as selected bases before any interpreter probe' {
        foreach ($name in @('.WinBookSplit-venv-backup-authored','.WinBookSplit-incomplete-authored')) {
            $app=New-SetupApplication
            $preserved=Join-Path $app $name
            $null=[IO.Directory]::CreateDirectory($preserved)
            $python=Join-Path $preserved 'python.exe'
            [IO.File]::WriteAllText($python,'SYNTHETIC non-executable fixture must be rejected before probing.')
            $before=Get-SetupSnapshot $app
            Mock Test-WinBookSplitPythonCandidate { throw 'Preserved artifacts must not be probed as working bases.' }
            $result=Invoke-WinBookSplitSetup -ApplicationRoot $app -PythonPath $python -Wheelhouse $Wheelhouse -NonInteractive
            Write-SetupObservation 'preserved-base' $result
            $result.code | Should -BeExactly 'runtime_invalid'
            $result.exit_code | Should -Be 3
            Assert-SetupNoMutation $app $before
        }
        Should-Invoke Test-WinBookSplitPythonCandidate -Times 0 -Exactly
    }
    It 'fails closed when the policy identity changes at the held preparation reread' {
        $app=New-SetupApplication
        $before=Get-SetupSnapshot $app
        $script:setupPolicyReadCount=0
        Mock Get-WinBookSplitCompatibilityPolicy {
            $identity=& $script:setupNativePolicy -PolicyPath $PolicyPath -WithIdentity:$WithIdentity
            if ($WithIdentity) {
                $script:setupPolicyReadCount++
                if ($script:setupPolicyReadCount -eq 3) {
                    # SYNTHETIC identity race; the policy bytes and parsed contents remain unchanged.
                    $identity.Sha256=('0'*64)
                }
            }
            return $identity
        }
        $result=Invoke-WinBookSplitSetup -ApplicationRoot $app -PythonPath $TestPython -Wheelhouse $Wheelhouse -NonInteractive
        Write-SetupObservation 'policy-identity-race' $result
        $script:setupPolicyReadCount | Should -Be 3
        $result.code | Should -BeExactly 'setup_source_invalid'
        $result.exit_code | Should -Be 3
        @($result.operations).Count | Should -Be 0
        Assert-SetupNoMutation $app $before
    }
    It 'fails missing or invalid selected Python with no installer fallback or prompts' {
        foreach ($python in @('python.exe',(Join-Path $TestDrive 'missing-python.exe'))) {
            $app=New-SetupApplication
            $before=Get-SetupSnapshot $app
            Mock Read-Host { throw 'NonInteractive must never prompt.' }
            Mock Resolve-WinBookSplitAutomaticPython { throw 'An explicit Python must never silently fall back.' }
            $result=Invoke-WinBookSplitSetup -ApplicationRoot $app -PythonPath $python -Wheelhouse $Wheelhouse -NonInteractive
            $result.code | Should -BeExactly 'runtime_invalid'
            $result.exit_code | Should -Be 3
            Assert-SetupNoMutation $app $before
        }
        Should-Invoke Read-Host -Times 0 -Exactly
        Should-Invoke Resolve-WinBookSplitAutomaticPython -Times 0 -Exactly
    }
    It 'fails missing and wrong-hash offline wheels before creating or repairing any environment' {
        foreach ($wrongHash in @($false,$true)) {
            $app=New-SetupApplication
            $before=Get-SetupSnapshot $app
            $bad=Join-Path $TestDrive ('w'+[guid]::NewGuid().ToString('N').Substring(0,12))
            $null=New-Item -ItemType Directory -Path $bad
            if ($wrongHash) { [IO.File]::WriteAllText((Join-Path $bad ([IO.Path]::GetFileName($script:setupWheel))),'Authored wrong-hash wheel.') }
            $result=Invoke-WinBookSplitSetup -ApplicationRoot $app -PythonPath $TestPython -Wheelhouse $bad -NonInteractive
            $result.exit_code | Should -Not -Be 0
            $result.changed | Should -BeFalse
            Assert-SetupNoMutation $app $before
        }
    }
    It 'uses only the reviewed wheel file when the wheelhouse contains an HTML remote-link trap' {
        $app=New-SetupApplication
        $trap=Join-Path $TestDrive ('w'+[guid]::NewGuid().ToString('N').Substring(0,12))
        $null=New-Item -ItemType Directory -Path $trap
        $wheel=Join-Path $trap ([IO.Path]::GetFileName($script:setupWheel))
        Copy-Item -LiteralPath $script:setupWheel -Destination $wheel
        [IO.File]::WriteAllText((Join-Path $trap 'index.html'),'<a href="http://127.0.0.1:1/forbidden/pypdf-6.20.0-py3-none-any.whl">remote trap</a>')
        $result=Invoke-WinBookSplitSetup -ApplicationRoot $app -PythonPath $TestPython -Wheelhouse $trap -NonInteractive
        $result.code | Should -BeExactly 'setup_complete'
        $install=@($result.operations | Where-Object { $_.phase -eq 'pip_install' })[0]
        @($install.arguments) | Should -Contain '--no-index'
        @($install.arguments) | Should -Contain $wheel
        @($install.arguments) | Should -Not -Contain $trap
        ($install.arguments -join ',') | Should -Not -Match 'index\.html|127\.0\.0\.1|--index-url'
    }
}
