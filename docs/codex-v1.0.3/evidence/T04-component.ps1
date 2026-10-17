# Public LF normalized, sanitized recipe from the actual external collector.
# Replace <EXISTING_MANAGED_3147> only for verification hashing of the prior
# observed interpreter. It is never passed to application discovery or setup.
# Exact local script hash and receipt correspondence: T04-native-originals.json.
[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$ApplicationRoot,[Parameter(Mandatory=$true)][string]$OutputPath)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false)
$script:checks=New-Object 'Collections.Generic.List[object]'
function Add-Check {param([string]$Name,[bool]$Passed);$script:checks.Add([pscustomobject]@{name=$Name;passed=$Passed})}
$files=@('WinBookSplit.ps1','Setup-WinBookSplit.ps1','Setup-WinBookSplit.bat','engine\WinBookSplit.Runtime.ps1','engine\WinBookSplit.Diagnostics.ps1','engine\WinBookSplit.Setup.ps1','engine\WinBookSplit.Process.ps1','engine\WinBookSplit.Compatibility.ps1','engine\WinBookSplit.Compatibility.json','engine\winbooksplit_compatibility.py')
# Only a previously observed T03 verification path is hashed; it is never
# supplied to runtime resolution, summary/remedy discovery or setup.
$verificationRuntime='<EXISTING_MANAGED_3147>'
function Get-Observation {
    $source=[ordered]@{};foreach($file in $files){$source[$file]=(Get-FileHash -LiteralPath (Join-Path $ApplicationRoot $file) -Algorithm SHA256).Hash.ToLowerInvariant()}
    $runtimeHash=$null;if([IO.File]::Exists($verificationRuntime)){$runtimeHash=(Get-FileHash -LiteralPath $verificationRuntime -Algorithm SHA256).Hash.ToLowerInvariant()}
    return [ordered]@{source_sha256=$source;path=$env:PATH;local_app_data=$env:LOCALAPPDATA;
        execution_policy=@(Get-ExecutionPolicy -List|ForEach-Object{[ordered]@{scope=$_.Scope.ToString();policy=$_.ExecutionPolicy.ToString()}});
        effective_policy=(Get-ExecutionPolicy).ToString();cwd=(Get-Location).ProviderPath;
        previously_observed_runtime_verification_path=$verificationRuntime;runtime_sha256=$runtimeHash;
        app_environment_exists=(Test-Path -LiteralPath (Join-Path $ApplicationRoot '.venv'));
        setup_lock_exists=(Test-Path -LiteralPath (Join-Path $ApplicationRoot '.WinBookSplit-setup.lock'));
        setup_artifacts=@(Get-ChildItem -LiteralPath $ApplicationRoot -Force |Where-Object{$_.Name.StartsWith('.WinBookSplit-venv-backup-') -or $_.Name.StartsWith('.WinBookSplit-incomplete-')}|Select-Object -ExpandProperty Name)}
}
$before=Get-Observation
$hostProcess=[Diagnostics.Process]::GetCurrentProcess();try{$hostPath=$hostProcess.MainModule.FileName}finally{$hostProcess.Dispose()}
$receipt=[ordered]@{schema_version=1;kind='ACTUAL_WINDOWS11_PYTHONORG_T04_DIAGNOSTIC_READINESS';
    meaning='Unchanged installed regular x64 Python 3.14.7 is reached automatically and correctly rejected by the retained 3.14.8 security floor. This is not compatible automatic-runtime or Store-channel/owner Windows10 acceptance.';
    host=[ordered]@{powershell=$PSVersionTable.PSVersion.ToString();edition=$PSVersionTable.PSEdition;executable=$hostPath;os=[Environment]::OSVersion.VersionString;process64bit=[Environment]::Is64BitProcess};
    source=[ordered]@{head=(& git -C $ApplicationRoot rev-parse HEAD);tree=(& git -C $ApplicationRoot rev-parse 'HEAD^{tree}')};
    invocation=[ordered]@{collector=$PSCommandPath;collector_sha256=(Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash.ToLowerInvariant();argv=@('-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File',$PSCommandPath,'-ApplicationRoot',$ApplicationRoot,'-OutputPath',$OutputPath)};
    commands=@('Resolve-WinBookSplitRuntime -ApplicationRoot $ApplicationRoot -TimeoutSeconds 3','Get-WinBookSplitDependencySummary -Failure $failure -ApplicationRoot $ApplicationRoot','Invoke-WinBookSplitSetup -ApplicationRoot $ApplicationRoot -Check -NonInteractive -TimeoutSeconds 3');
    before=$before;processing=$null;summary=$null;setup_check=$null;frontend=$null;after=$null;checks=@();passed=$false}
try {
    . (Join-Path $ApplicationRoot 'engine\WinBookSplit.Runtime.ps1')
    . (Join-Path $ApplicationRoot 'engine\WinBookSplit.Setup.ps1')
    $tokens=$null;$parseErrors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $ApplicationRoot 'WinBookSplit.ps1'),[ref]$tokens,[ref]$parseErrors)
    if(@($parseErrors).Count -ne 0){throw 'Application source parse failed.'}
    $function=@($ast.FindAll({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq 'Get-WinBookSplitDependencySummary'},$true))
    if($function.Count -ne 1){throw 'Expected exactly one actual application dependency summary.'}
    . ([scriptblock]::Create($function[0].Extent.Text))
    $failure=$null
    try {$selected=Resolve-WinBookSplitRuntime -ApplicationRoot $ApplicationRoot -TimeoutSeconds 3;$receipt.processing=[ordered]@{unexpected_selected=$selected}}
    catch {$failure=$_.Exception;$receipt.processing=[ordered]@{type=$failure.GetType().FullName;message=$failure.Message;code=$failure.Data['Code'];attempts=@($failure.Data['Attempts'])}}
    Add-Check 'actual automatic processing discovery rejects below-floor runtime' ($null -ne $failure -and $failure.Data['Code'] -ceq 'runtime_not_found')
    if($null -eq $failure){throw 'Unexpected compatible automatic selection; inspect actual result instead of forcing expectations.'}
    $summary=Get-WinBookSplitDependencySummary -Failure $failure -ApplicationRoot $ApplicationRoot
    $receipt.summary=$summary
    Add-Check 'summary reason security floor' ($summary.Reason -ceq 'python_security_floor')
    Add-Check 'summary stage compatibility' ($summary.Stage -ceq 'compatibility')
    Add-Check 'summary source manager' ($summary.Source -ceq 'manager')
    Add-Check 'summary observed Python 3.14.7 from manager' ($summary.ObservedVersion -ceq '3.14.7' -and $summary.ObservedSource -ceq 'manager')
    Add-Check 'summary actual observed path agrees with verification-only prior observation' ([string]::Equals($summary.ObservedPath,$verificationRuntime,[StringComparison]::OrdinalIgnoreCase))
    $concrete=@($failure.Data['Attempts']|Where-Object{$_.Source -ceq 'manager' -and $_.Reason -ceq 'python_security_floor'})
    Add-Check 'concrete child identity independently verified' ($concrete.Count -eq 1 -and $concrete[0].IdentityVerified -and $concrete[0].Stage -ceq 'compatibility')
    $setup=Invoke-WinBookSplitSetup -ApplicationRoot $ApplicationRoot -Check -NonInteractive -TimeoutSeconds 3
    $receipt.setup_check=$setup
    Add-Check 'inprocess readiness preserves reason stage source and observed version' ($setup.reason -ceq 'python_security_floor' -and $setup.stage -ceq 'compatibility' -and $setup.source -ceq 'manager' -and $setup.observed_version -ceq '3.14.7')
    Add-Check 'inprocess readiness exits 3 and performs no setup operations' ($setup.exit_code -eq 3 -and -not $setup.changed -and @($setup.operations).Count -eq 0)
    $frontendArguments=@('-NoProfile','-NonInteractive','-ExecutionPolicy','RemoteSigned','-File',(Join-Path $ApplicationRoot 'Setup-WinBookSplit.ps1'),'-Check','-NonInteractive')
    $probe=Invoke-WinBookSplitProcess -Path $hostPath -Arguments $frontendArguments -WorkingDirectory $ApplicationRoot -TimeoutSeconds 60 -CaptureLimitBytes 1048576
    $records=@($probe.Stdout -split '\r?\n'|Where-Object{$_.StartsWith('[SETUP] ')})
    $frontendResult=$null;if($records.Count -eq 1){$frontendResult=$records[0].Substring(8)|ConvertFrom-Json}
    $receipt.frontend=[ordered]@{path=$hostPath;argv=$frontendArguments;probe=$probe;result=$frontendResult}
    Add-Check 'actual frontend readiness exits native 3 safely' ($probe.ExitCode -eq 3 -and $probe.JobAssigned -and $probe.ParentStopped -and $probe.DescendantsStopped -and $probe.StreamsComplete -and -not $probe.StartError -and -not $probe.StopError -and -not $probe.StreamError -and -not $probe.TimedOut -and -not $probe.Cancelled -and -not $probe.StdoutTruncated -and -not $probe.StderrTruncated)
    Add-Check 'frontend structured readiness retains reason stage source and observed version' ($null -ne $frontendResult -and $frontendResult.reason -ceq 'python_security_floor' -and $frontendResult.stage -ceq 'compatibility' -and $frontendResult.source -ceq 'manager' -and $frontendResult.observed_version -ceq '3.14.7' -and $frontendResult.exit_code -eq 3 -and -not $frontendResult.changed -and @($frontendResult.operations).Count -eq 0)
    Add-Check 'frontend human output names reason and actual selected version' ($probe.Stdout.Contains('Discovery: python_security_floor; stage=compatibility; source=manager') -and $probe.Stdout.Contains('(3.14.7; manager)'))
} catch {$receipt['collector_error']=[ordered]@{message=$_.Exception.Message;stack=$_.ScriptStackTrace};Add-Check 'collector completed without unexpected error' $false}
finally {
    $after=Get-Observation;$receipt.after=$after
    Add-Check 'source bytes unchanged' (($before.source_sha256|ConvertTo-Json -Compress) -ceq ($after.source_sha256|ConvertTo-Json -Compress))
    Add-Check 'PATH LOCALAPPDATA CWD and all policy scopes unchanged' ($before.path -ceq $after.path -and $before.local_app_data -ceq $after.local_app_data -and $before.cwd -ceq $after.cwd -and ($before.execution_policy|ConvertTo-Json -Compress) -ceq ($after.execution_policy|ConvertTo-Json -Compress))
    Add-Check 'RemoteSigned remains process policy and effective policy' ($before.effective_policy -ceq 'RemoteSigned' -and $after.effective_policy -ceq 'RemoteSigned' -and @($after.execution_policy|Where-Object{$_.scope -ceq 'Process' -and $_.policy -ceq 'RemoteSigned'}).Count -eq 1)
    Add-Check 'no app venv lock or setup artifact created' (-not $before.app_environment_exists -and -not $after.app_environment_exists -and -not $before.setup_lock_exists -and -not $after.setup_lock_exists -and ($before.setup_artifacts|ConvertTo-Json -Compress) -ceq ($after.setup_artifacts|ConvertTo-Json -Compress))
    Add-Check 'pre-existing runtime bytes unchanged' ($null -ne $before.runtime_sha256 -and $before.runtime_sha256 -ceq $after.runtime_sha256)
    $receipt.checks=@($script:checks.ToArray());$receipt.passed=@($script:checks|Where-Object{-not $_.passed}).Count -eq 0
    $json=$receipt|ConvertTo-Json -Depth 50
    [IO.File]::WriteAllText($OutputPath,$json,(New-Object Text.UTF8Encoding($false)))
}
if($receipt.passed){Write-Output ('PASS '+$script:checks.Count+' checks: '+$OutputPath);exit 0}
Write-Output ('FAIL: '+$OutputPath);exit 1