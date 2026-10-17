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
    $applicationAst=[Management.Automation.Language.Parser]::ParseFile((Join-Path $RepositoryRoot 'WinBookSplit.ps1'),[ref]$tokens,[ref]$parseErrors)
    if(@($parseErrors).Count -ne 0){throw 'The trusted application has syntax errors.'}
    $summaryFunctions=@($applicationAst.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq 'Get-WinBookSplitDependencySummary'},$true))
    if($summaryFunctions.Count -ne 1){throw 'Expected one trusted application dependency summary function.'}
    . ([scriptblock]::Create($summaryFunctions[0].Extent.Text))
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
        param([string]$Name)
        if ([string]::IsNullOrWhiteSpace($Name)) { $Name='a'+[guid]::NewGuid().ToString('N').Substring(0,12) }
        $app=Join-Path $TestDrive $Name
        $engine=Join-Path $app 'engine'
        $null=New-Item -ItemType Directory -Path $engine
        foreach ($name in @('Setup-WinBookSplit.ps1','Setup-WinBookSplit.bat','requirements.txt')) {
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
    function Read-SetupFrontendResult {
        param($Probe)
        $Probe.ParentStopped | Should -BeTrue
        $Probe.DescendantsStopped | Should -BeTrue
        $Probe.StreamsComplete | Should -BeTrue
        $Probe.TimedOut | Should -BeFalse
        $frames=@([regex]::Matches($Probe.Stdout,'(?m)^\[SETUP\] ([^\r\n]+)'))
        $frames.Count | Should -Be 1
        $record=$frames[0].Groups[1].Value | ConvertFrom-Json
        $record.exit_code | Should -Be $Probe.ExitCode
        return $record
    }
    function Invoke-SetupOwnedRecipe {
        param([string]$WorkingDirectory,[string]$Action,[string]$Bat,[string[]]$Arguments,[string]$DirectScript)
        # The command is captured only from the current owned frontend. Other
        # paths/argument vectors remain JSON data, never interpolated shell code.
        $control=Join-Path $WorkRoot ('setup-control-'+[guid]::NewGuid().ToString('N')+'.json')
        $driver=Join-Path $WorkRoot ('setup-recipe-'+[guid]::NewGuid().ToString('N')+'.ps1')
        $data=[ordered]@{action=$Action;bat=$Bat;arguments=@($Arguments);direct_script=$DirectScript}
        [IO.File]::WriteAllText($control,($data | ConvertTo-Json -Depth 4),(New-Object Text.UTF8Encoding($false)))
        $source=@'
param([string]$Control)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
$p=[IO.File]::ReadAllText($Control)|ConvertFrom-Json
$before=@(Get-ExecutionPolicy -List | ForEach-Object {@{scope=$_.Scope.ToString();policy=$_.ExecutionPolicy.ToString()}})
Write-Host ('[POLICY_BEFORE] '+($before|ConvertTo-Json -Compress))
$code=1
try {
    if (-not [string]::IsNullOrEmpty($p.action)) {
        # Preserve the emitted text as the executed command. The parent already
        # verifies its finite literal AST and source before writing this control.
        Write-Host ('[EXACT_ACTION] '+([string]$p.action))
        & ([scriptblock]::Create([string]$p.action))
        $code=$LASTEXITCODE
    } elseif (-not [string]::IsNullOrEmpty($p.bat)) {
        & ([string]$p.bat) @([string[]]$p.arguments)
        $code=$LASTEXITCODE
    } else {
        try { & ([string]$p.direct_script) @([string[]]$p.arguments);$code=$LASTEXITCODE }
        catch [Management.Automation.PSSecurityException] {
            Write-Host '[DIRECT_POLICY_DENIAL] PSSecurityException'
            $code=1
        }
    }
} finally {
    $after=@(Get-ExecutionPolicy -List | ForEach-Object {@{scope=$_.Scope.ToString();policy=$_.ExecutionPolicy.ToString()}})
    Write-Host ('[POLICY_AFTER] '+($after|ConvertTo-Json -Compress))
}
exit $code
'@
        [IO.File]::WriteAllText($driver,$source,(New-Object Text.UTF8Encoding($true)))
        $hostPath=(Get-WinBookSplitSetupHostObservation).executable
        $probe=Invoke-WinBookSplitProcess -Path $hostPath -Arguments @('-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File',$driver,'-Control',$control) -WorkingDirectory $WorkingDirectory -TimeoutSeconds 120
        $probe.ParentStopped | Should -BeTrue
        $probe.DescendantsStopped | Should -BeTrue
        $probe.StreamsComplete | Should -BeTrue
        $probe.TimedOut | Should -BeFalse
        return $probe
    }
    function Assert-SetupActionLiterals {
        param([string]$Action,[string]$App,[string]$Base,[string]$Wheels,[switch]$Repair)
        $tokens=$null;$errors=$null
        $ast=[Management.Automation.Language.Parser]::ParseInput($Action,[ref]$tokens,[ref]$errors)
        @($errors).Count | Should -Be 0
        $ast.EndBlock.Statements.Count | Should -Be 1
        $pipeline=$ast.EndBlock.Statements[0]
        $pipeline | Should -BeOfType Management.Automation.Language.PipelineAst
        $pipeline.PipelineElements.Count | Should -Be 1
        $command=$pipeline.PipelineElements[0]
        $command | Should -BeOfType Management.Automation.Language.CommandAst
        $command.InvocationOperator.ToString() | Should -BeExactly 'Ampersand'
        @($command.Redirections).Count | Should -Be 0
        $values=@(foreach($element in $command.CommandElements) {
            if ($element -is [Management.Automation.Language.StringConstantExpressionAst]) { [string]$element.Value }
            elseif ($element -is [Management.Automation.Language.CommandParameterAst] -and $null -eq $element.Argument) { '-'+$element.ParameterName }
            else { throw 'An emitted setup command must contain only literal command elements.' }
        })
        $hostPath=Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::System)) 'WindowsPowerShell\v1.0\powershell.exe'
        $expected=@($hostPath,'-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $App 'Setup-WinBookSplit.ps1'),'-NonInteractive')
        if ($Repair) { $expected+='-Repair' }
        $expected+=@('-PythonPath',$Base)
        if (-not [string]::IsNullOrWhiteSpace($Wheels)) { $expected+=@('-Wheelhouse',$Wheels) }
        ($values|ConvertTo-Json -Compress) | Should -BeExactly ($expected|ConvertTo-Json -Compress)
    }
    function Assert-SetupRecipePolicies {
        param($Probe)
        $before=[regex]::Match($Probe.Stdout,'(?m)^\[POLICY_BEFORE\] ([^\r\n]+)').Groups[1].Value
        $after=[regex]::Match($Probe.Stdout,'(?m)^\[POLICY_AFTER\] ([^\r\n]+)').Groups[1].Value
        $before | Should -Not -BeNullOrEmpty
        $after | Should -BeExactly $before
        $scopes=ConvertFrom-Json -InputObject $before
        @($scopes|Where-Object {$_.scope -eq 'Process'})[0].policy | Should -BeExactly 'RemoteSigned'
    }
    $script:healthySetupApp=New-SetupApplication
}

Describe 'T04 exact emitted setup and native BAT contracts in explicit developer fixtures' -Tag 'STORE-021','STORE-022','STORE-023','STORE-024' {
    BeforeAll {
        # A disposable package-free developer base is intentionally explicit.
        # It is not the existing-Store automatic-discovery acceptance scenario.
        $script:bareSetupRoot=Join-Path $TestDrive 'base without packages'
        $create=Invoke-WinBookSplitProcess -Path $TestPython -Arguments @('-I','-B','-m','venv','--without-pip',$script:bareSetupRoot) -WorkingDirectory $TestDrive -TimeoutSeconds 60
        $create.ExitCode | Should -Be 0
        $create.ParentStopped | Should -BeTrue
        $create.DescendantsStopped | Should -BeTrue
        $script:bareSetupPython=Join-Path $script:bareSetupRoot 'Scripts\python.exe'
        $script:setupUnrelated=Join-Path $TestDrive ('cwd '+[char]0xc5+[char]0x65e5)
        $null=New-Item -ItemType Directory -Path $script:setupUnrelated
        $script:spacedSetupWheels=Join-Path $TestDrive ('wheels '+[char]0xc5+" O'Brien")
        $null=New-Item -ItemType Directory -Path $script:spacedSetupWheels
        Copy-Item -LiteralPath $script:setupWheel -Destination $script:spacedSetupWheels
    }
    It 'executes the actual printed Check action verbatim, then repeats it as a healthy no-op' {
        $app=New-SetupApplication -Name ('action '+[char]0xc5+[char]0x65e5+" O'Brien & [1]")
        $before=Get-SetupSnapshot $app
        $hostPath=(Get-WinBookSplitSetupHostObservation).executable
        $checkProbe=Invoke-WinBookSplitProcess -Path $hostPath -Arguments @('-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File',(Join-Path $app 'Setup-WinBookSplit.ps1'),'-NonInteractive','-Check','-PythonPath',$script:bareSetupPython,'-Wheelhouse',$script:spacedSetupWheels) -WorkingDirectory $script:setupUnrelated -TimeoutSeconds 60
        $check=Read-SetupFrontendResult $checkProbe
        $check.code | Should -BeExactly 'setup_required'
        $check.reason | Should -BeExactly 'missing_pypdf'
        $check.stage | Should -BeExactly 'package'
        $check.source | Should -BeExactly 'explicit'
        $check.observed_python | Should -BeExactly $script:bareSetupPython
        $check.observed_version | Should -BeExactly $script:setupExpectedVersion
        $check.observed_source | Should -BeExactly 'explicit'
        $check.selected_version | Should -BeExactly $script:setupExpectedVersion
        $check.message.Contains($check.action) | Should -BeTrue
        Assert-SetupActionLiterals $check.action $app $script:bareSetupPython $script:spacedSetupWheels
        Assert-SetupNoMutation $app $before
        $processingFailure=$null
        try { $null=Resolve-WinBookSplitRuntime -ApplicationRoot $app -PythonPath $script:bareSetupPython }
        catch { $processingFailure=$_.Exception }
        $processingFailure | Should -Not -BeNullOrEmpty
        $applicationSummary=Get-WinBookSplitDependencySummary -Failure $processingFailure -ApplicationRoot $app -PythonPath $script:bareSetupPython
        $applicationSummary.Reason | Should -BeExactly 'missing_pypdf'
        $applicationSummary.Stage | Should -BeExactly 'package'
        $applicationSummary.Source | Should -BeExactly 'explicit'
        $applicationSummary.ObservedVersion | Should -BeExactly $script:setupExpectedVersion
        $applicationSummary.ActionIsCommand | Should -BeTrue
        Assert-SetupActionLiterals -Action $applicationSummary.Action -App $app -Base $script:bareSetupPython
        Assert-SetupNoMutation $app $before
        $firstProbe=Invoke-SetupOwnedRecipe -WorkingDirectory $script:setupUnrelated -Action $check.action
        $first=Read-SetupFrontendResult $firstProbe
        $first.code | Should -BeExactly 'setup_complete'
        $first.changed | Should -BeTrue
        $firstProbe.Stdout.Contains('[EXACT_ACTION] '+$check.action) | Should -BeTrue
        Assert-SetupRecipePolicies $firstProbe
        $healthy=Get-SetupSnapshot $app
        $secondProbe=Invoke-SetupOwnedRecipe -WorkingDirectory $script:setupUnrelated -Action $check.action
        $second=Read-SetupFrontendResult $secondProbe
        $second.code | Should -BeExactly 'setup_ready'
        $second.changed | Should -BeFalse
        @($second.operations).Count | Should -Be 0
        Get-SetupSnapshot $app | Should -BeExactly $healthy
        Assert-SetupRecipePolicies $secondProbe
        # The application's own emitted Action has no offline-wheel option.
        # Execute it verbatim against the now-healthy authoritative .venv: this
        # proves the recipe while healthy no-op rules prohibit package activity.
        $applicationActionProbe=Invoke-SetupOwnedRecipe -WorkingDirectory $script:setupUnrelated -Action $applicationSummary.Action
        $applicationAction=Read-SetupFrontendResult $applicationActionProbe
        $applicationAction.code | Should -BeExactly 'setup_ready'
        $applicationAction.changed | Should -BeFalse
        @($applicationAction.operations).Count | Should -Be 0
        Get-SetupSnapshot $app | Should -BeExactly $healthy
        Assert-SetupRecipePolicies $applicationActionProbe
        Write-SetupObservation 't04-verbatim-action' ([pscustomobject]@{scope='EXPLICIT_DEVELOPER_FIXTURE_NOT_STORE_GATE';check_probe=$checkProbe;check=$check;action=$check.action;first_probe=$firstProbe;first=$first;second_probe=$secondProbe;second=$second;application_summary=$applicationSummary;application_action_probe=$applicationActionProbe;application_action=$applicationAction})
    }
    It 'forwards all setup BAT options and exits from unrelated Unicode CWD while preserving an explicitly repaired old environment' {
        $app=New-SetupApplication -Name ('bat '+[char]0xc5+[char]0x65e5+" O'Brien & [1]")
        $null=New-Item -ItemType Directory -Path (Join-Path $app '.venv\Scripts')
        [IO.File]::WriteAllText((Join-Path $app '.venv\preserve-old.txt'),'Authored old BAT environment must survive.')
        $old=Get-SetupSnapshot (Join-Path $app '.venv')
        $arguments=@('-NonInteractive','-PythonPath',$script:bareSetupPython,'-Wheelhouse',$script:spacedSetupWheels)
        $bat=Join-Path $app 'Setup-WinBookSplit.bat'
        $rejectProbe=Invoke-SetupOwnedRecipe -WorkingDirectory $script:setupUnrelated -Bat $bat -Arguments $arguments
        $rejected=Read-SetupFrontendResult $rejectProbe
        $rejectProbe.ExitCode | Should -Be 3
        $rejected.code | Should -BeExactly 'setup_repair_required'
        Get-SetupSnapshot (Join-Path $app '.venv') | Should -BeExactly $old
        Assert-SetupActionLiterals $rejected.action $app $script:bareSetupPython $script:spacedSetupWheels -Repair
        # Execute the actual displayed repair Action before independently testing
        # BAT's Repair forwarding against another owned broken environment.
        $repairActionProbe=Invoke-SetupOwnedRecipe -WorkingDirectory $script:setupUnrelated -Action $rejected.action
        $repairAction=Read-SetupFrontendResult $repairActionProbe
        $repairAction.code | Should -BeExactly 'setup_complete'
        Get-SetupSnapshot $repairAction.backup_path | Should -BeExactly $old
        $batApp=New-SetupApplication -Name ('repair bat '+[char]0xc5)
        $null=New-Item -ItemType Directory -Path (Join-Path $batApp '.venv\Scripts')
        [IO.File]::WriteAllText((Join-Path $batApp '.venv\preserve-old.txt'),'Authored old BAT environment must survive.')
        $batOld=Get-SetupSnapshot (Join-Path $batApp '.venv')
        $bat=Join-Path $batApp 'Setup-WinBookSplit.bat'
        $repairProbe=Invoke-SetupOwnedRecipe -WorkingDirectory $script:setupUnrelated -Bat $bat -Arguments ($arguments+@('-Repair'))
        $repaired=Read-SetupFrontendResult $repairProbe
        $repairProbe.ExitCode | Should -Be 0
        $repaired.code | Should -BeExactly 'setup_complete'
        Get-SetupSnapshot $repaired.backup_path | Should -BeExactly $batOld
        $healthy=Get-SetupSnapshot $batApp
        $checkProbe=Invoke-SetupOwnedRecipe -WorkingDirectory $script:setupUnrelated -Bat $bat -Arguments ($arguments+@('-Check'))
        $check=Read-SetupFrontendResult $checkProbe
        $check.code | Should -BeExactly 'setup_ready'
        $check.changed | Should -BeFalse
        @($check.operations).Count | Should -Be 0
        Get-SetupSnapshot $batApp | Should -BeExactly $healthy
        foreach($probe in @($rejectProbe,$repairActionProbe,$repairProbe,$checkProbe)){Assert-SetupRecipePolicies $probe}
        Write-SetupObservation 't04-bat-repair' ([pscustomobject]@{scope='EXPLICIT_DEVELOPER_FIXTURE_NOT_STORE_GATE';rejected_probe=$rejectProbe;rejected=$rejected;repair_action_probe=$repairActionProbe;repair_action=$repairAction;bat_repair_probe=$repairProbe;bat_repair=$repaired;bat_check_probe=$checkProbe;bat_check=$check;old_preserved=$true})
    }
    It 'retains real MOTW through RemoteSigned direct denial and process-only BAT Check invocation' {
        $app=New-SetupApplication
        $marks=@(foreach($file in Get-ChildItem -LiteralPath $app -Filter '*.ps1' -File -Recurse) {
            $ads="[ZoneTransfer]`r`nZoneId=3`r`n"
            Set-Content -LiteralPath $file.FullName -Stream Zone.Identifier -Value $ads -Encoding ASCII -NoNewline
            [pscustomobject]@{path=$file.FullName;ads=$ads;sha256=Get-SetupFileDigest $file.FullName}
        })
        $marks.Count | Should -BeGreaterThan 3
        $before=Get-SetupSnapshot $app
        $arguments=@('-NonInteractive','-Check','-PythonPath',$script:bareSetupPython,'-Wheelhouse',$script:spacedSetupWheels)
        $denied=Invoke-SetupOwnedRecipe -WorkingDirectory $script:setupUnrelated -DirectScript (Join-Path $app 'Setup-WinBookSplit.ps1') -Arguments $arguments
        $denied.ExitCode | Should -Be 1
        $denied.Stdout | Should -Match '\[DIRECT_POLICY_DENIAL\] PSSecurityException'
        $denied.Stdout | Should -Not -Match '(?m)^\[SETUP\] '
        Assert-SetupNoMutation $app $before
        # PowerShell Core and inbox Windows PowerShell can read different
        # persistent policy stores. Observe the actual inbox host separately;
        # compare each process only with its own policy domain.
        $inboxHost=Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::System)) 'WindowsPowerShell\v1.0\powershell.exe'
        $inboxPolicyScript=Join-Path $WorkRoot ('setup-inbox-policy-'+[guid]::NewGuid().ToString('N')+'.ps1')
        $inboxPolicySource="[Console]::OutputEncoding=New-Object Text.UTF8Encoding(`$false)`r`n"+$hostFunctions[0].Extent.Text+"`r`nWrite-Host ('[INBOX_POLICY] '+(Get-WinBookSplitSetupHostObservation|ConvertTo-Json -Depth 6 -Compress))`r`n"
        [IO.File]::WriteAllText($inboxPolicyScript,$inboxPolicySource,(New-Object Text.UTF8Encoding($true)))
        $inboxPolicyProbe=Invoke-WinBookSplitProcess -Path $inboxHost -Arguments @('-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File',$inboxPolicyScript) -WorkingDirectory $script:setupUnrelated -TimeoutSeconds 60
        $inboxPolicyProbe.ExitCode | Should -Be 0
        $inboxPolicyProbe.ParentStopped | Should -BeTrue
        $inboxPolicyProbe.DescendantsStopped | Should -BeTrue
        $inboxPolicyProbe.StreamsComplete | Should -BeTrue
        $inboxPolicyFrames=@([regex]::Matches($inboxPolicyProbe.Stdout,'(?m)^\[INBOX_POLICY\] ([^\r\n]+)'))
        $inboxPolicyFrames.Count | Should -Be 1
        $inboxPolicy=$inboxPolicyFrames[0].Groups[1].Value|ConvertFrom-Json
        $inboxPolicy.executable | Should -BeExactly $inboxHost
        $inboxPolicy.edition | Should -BeExactly 'Desktop'
        $inboxPolicy.version | Should -Match '\A5\.1\.'
        $inboxPolicy.execution_policy.status | Should -BeExactly 'observed'
        $batProbe=Invoke-SetupOwnedRecipe -WorkingDirectory $script:setupUnrelated -Bat (Join-Path $app 'Setup-WinBookSplit.bat') -Arguments $arguments
        $checked=Read-SetupFrontendResult $batProbe
        $batProbe.ExitCode | Should -Be 3
        $checked.code | Should -BeExactly 'setup_required'
        $checked.changed | Should -BeFalse
        @($checked.operations).Count | Should -Be 0
        $checked.host.execution_policy.effective | Should -BeExactly 'Bypass'
        foreach($scope in @('MachinePolicy','UserPolicy','Process','CurrentUser','LocalMachine')) {
            $childPolicy=@($checked.host.execution_policy.scopes | Where-Object {$_.scope -eq $scope})[0].policy
            if ($scope -eq 'Process') { $childPolicy | Should -BeExactly 'Bypass' }
            else { $childPolicy | Should -BeExactly @($inboxPolicy.execution_policy.scopes | Where-Object {$_.scope -eq $scope})[0].policy }
        }
        Assert-SetupNoMutation $app $before
        Assert-SetupActionLiterals $checked.action $app $script:bareSetupPython $script:spacedSetupWheels
        $actionProbe=Invoke-SetupOwnedRecipe -WorkingDirectory $script:setupUnrelated -Action $checked.action
        $installed=Read-SetupFrontendResult $actionProbe
        $installed.code | Should -BeExactly 'setup_complete'
        foreach($mark in $marks) {
            Get-SetupFileDigest $mark.path | Should -BeExactly $mark.sha256
            Get-Content -LiteralPath $mark.path -Stream Zone.Identifier -Raw -Encoding ASCII | Should -BeExactly $mark.ads
        }
        Assert-SetupRecipePolicies $denied
        Assert-SetupRecipePolicies $batProbe
        Assert-SetupRecipePolicies $actionProbe
        Write-SetupObservation 't04-motw-policy' ([pscustomobject]@{scope='ACTUAL_WINDOWS_OWNED_COPY_MOTW_NOT_ENFORCED_GPO';direct_probe=$denied;inbox_policy_probe=$inboxPolicyProbe;inbox_policy=$inboxPolicy;bat_probe=$batProbe;bat_check=$checked;action_probe=$actionProbe;installed=$installed;script_marks_preserved=$true;enforced_organization_policy='NOT_RUN: no enforced policy on this host'})
    }
}

Describe 'V102 local setup host and policy observations' -Tag 'V102-AC-041','V102-AC-042' {
    It 'escapes human frontend controls while preserving raw structured local diagnostics' {
        $app=New-SetupApplication
        # SYNTHETIC provider output, actual unchanged frontend and host process.
        $raw='Authored local detail '+[char]27+"[31m`r`nnext line"
        $stub=@'
. (Join-Path $PSScriptRoot 'WinBookSplit.Diagnostics.ps1')
function Invoke-WinBookSplitSetup {
    $raw='Authored local detail '+[char]27+"[31m`r`nnext line"
    return [pscustomobject]@{protocol='winbooksplit.setup';schema_version=1;status='failed';code='runtime_invalid';exit_code=3;changed=$false;
        message=$raw;display_message=$raw;reason='rejected_entry_point';stage='entry';source='explicit';observed_python=$raw;observed_version='3.14.8';observed_source='explicit';action=$null}
}
'@
        [IO.File]::WriteAllText((Join-Path $app 'engine\WinBookSplit.Setup.ps1'),$stub,(New-Object Text.UTF8Encoding($true)))
        $hostPath=(Get-WinBookSplitSetupHostObservation).executable
        $probe=Invoke-WinBookSplitProcess -Path $hostPath -Arguments @('-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File',(Join-Path $app 'Setup-WinBookSplit.ps1'),'-NonInteractive','-Check') -WorkingDirectory $app -TimeoutSeconds 60
        $record=Read-SetupFrontendResult $probe
        $record.message | Should -BeExactly $raw
        $record.display_message | Should -BeExactly $raw
        $record.observed_python | Should -BeExactly $raw
        $human=$probe.Stdout.Substring(0,$probe.Stdout.IndexOf('[SETUP] '))
        $human.Contains([string][char]27) | Should -BeFalse
        $human | Should -Match '\\u001b\[31m\\u000d\\u000anext line'
        $human | Should -Match 'Discovery: rejected_entry_point; stage=entry; source=explicit'
        Write-SetupObservation 't04-display-privacy' ([pscustomobject]@{scope='SYNTHETIC_PROVIDER_ACTUAL_FRONTEND';probe=$probe;record=$record;raw_local_detail_preserved=$true;human_controls_escaped=$true})
    }
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
        $selected.PypdfVersion | Should -BeExactly $script:setupPolicy.pypdf.version
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
    It 'preserves repair-required authority when an optional base hint fails' {
        $app=New-BrokenSetupApplication
        $before=Get-SetupSnapshot $app
        Mock Find-WinBookSplitSetupBase { throw 'SYNTHETIC optional base hint failure.' }
        $result=Invoke-WinBookSplitSetup -ApplicationRoot $app -PythonPath $TestPython -Wheelhouse $Wheelhouse -NonInteractive
        $result.code | Should -BeExactly 'setup_repair_required'
        $result.exit_code | Should -Be 3
        $result.changed | Should -BeFalse
        $result.source | Should -BeExactly 'application_venv'
        $result.action | Should -Match ' -Repair(?: |$)'
        Get-SetupSnapshot $app | Should -BeExactly $before
    }
    It 'keeps interruption fatal when finding an optional repair base hint' {
        $app=New-BrokenSetupApplication
        $before=Get-SetupSnapshot $app
        Mock Find-WinBookSplitSetupBase {
            $failure=New-Object InvalidOperationException('SYNTHETIC interruption in the optional hint route.')
            $failure.Data['Code']='processor_cancelled'
            throw $failure
        }
        $result=Invoke-WinBookSplitSetup -ApplicationRoot $app -PythonPath $TestPython -Wheelhouse $Wheelhouse -NonInteractive
        $result.code | Should -BeExactly 'processor_cancelled'
        $result.exit_code | Should -Be 130
        $result.changed | Should -BeFalse
        $result.action | Should -BeNullOrEmpty
        Get-SetupSnapshot $app | Should -BeExactly $before
    }
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
        [IO.File]::WriteAllText((Join-Path $trap 'index.html'),('<a href="http://127.0.0.1:1/forbidden/'+[IO.Path]::GetFileName($script:setupWheel)+'">remote trap</a>'))
        $result=Invoke-WinBookSplitSetup -ApplicationRoot $app -PythonPath $TestPython -Wheelhouse $trap -NonInteractive
        $result.code | Should -BeExactly 'setup_complete'
        $install=@($result.operations | Where-Object { $_.phase -eq 'pip_install' })[0]
        @($install.arguments) | Should -Contain '--no-index'
        @($install.arguments) | Should -Contain $wheel
        @($install.arguments) | Should -Not -Contain $trap
        ($install.arguments -join ',') | Should -Not -Match 'index\.html|127\.0\.0\.1|--index-url'
    }
}
