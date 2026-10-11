[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$RepositoryRoot,
      [Parameter(Mandatory=$true)][string]$CoreRoot,
      [Parameter(Mandatory=$true)][string]$TestPython,
      [Parameter(Mandatory=$true)][string]$WorkRoot,
      [Parameter(Mandatory=$true)][string]$ReportPath)
$ErrorActionPreference='Stop'
$PSModuleAutoloadingPreference='None'
Import-Module ([IO.Path]::Combine($PSHOME,'Modules','Microsoft.PowerShell.Utility','Microsoft.PowerShell.Utility.psd1')) -ErrorAction Stop
Import-Module ([IO.Path]::Combine($PSHOME,'Modules','Microsoft.PowerShell.Management','Microsoft.PowerShell.Management.psd1')) -ErrorAction Stop
. ([IO.Path]::Combine($RepositoryRoot,'engine','WinBookSplit.Diagnostics.ps1'))
. ([IO.Path]::Combine($RepositoryRoot,'engine','WinBookSplit.Process.ps1'))
. ([IO.Path]::Combine($RepositoryRoot,'engine','WinBookSplit.Logging.ps1'))
. ([IO.Path]::Combine($RepositoryRoot,'engine','WinBookSplit.Support.ps1'))
. ([IO.Path]::Combine($CoreRoot,'Invoke-NativeAcceptance.ps1'))
Initialize-WinBookSplitLogTypes
Initialize-WinBookSplitSupport
Initialize-NativeAcceptanceDirectory
[WinBookSplitNative.Directory]::CreateExclusive($WorkRoot)
$cases=New-Object 'Collections.Generic.List[object]'
$started=[DateTime]::UtcNow.ToString('o')
$paths=@([IO.Path]::Combine($CoreRoot,'Invoke-NativeAcceptance.ps1'),
    [IO.Path]::Combine($RepositoryRoot,'engine','WinBookSplit.Process.ps1'),
    [IO.Path]::Combine($RepositoryRoot,'engine','WinBookSplit.Logging.ps1'),
    [IO.Path]::Combine($RepositoryRoot,'engine','WinBookSplit.Support.ps1'),
    [IO.Path]::Combine($RepositoryRoot,'engine','WinBookSplit.Diagnostics.ps1'))
$before=@{};foreach ($path in $paths) { $before[$path]=(Get-NativeAcceptanceFileHash -Path $path -MaximumBytes 1048576).sha256 }
function Require-Control($Value,[string]$Reason) { if (-not $Value) { throw $Reason } }
function Run-Control([string]$Id,[scriptblock]$Action) {
    $row=[ordered]@{id=$Id;status='FAIL';error=$null;scope='CONTROLLED_NATIVE_TOOLING_ONLY'}
    try { & $Action; $row.status='PASS' } catch { $row.error=$_.Exception.Message }
    $cases.Add([pscustomobject]$row)
}

Run-Control 'bounded-hash-and-json' {
    $path=[IO.Path]::Combine($WorkRoot,'small.json')
    Write-NativeAcceptanceJson -Path $path -Record @{authored='finite'}
    $parsed=Read-NativeAcceptanceJson -Path $path -MaximumBytes 1024
    Require-Control ($parsed.value.authored -ceq 'finite') 'Bounded exact JSON did not roundtrip.'
    $hashed=Get-NativeAcceptanceFileHash -Path $path -MaximumBytes 1024
    Require-Control ($hashed.sha256 -ceq $parsed.sha256 -and $hashed.bytes -eq $parsed.bytes) 'Hash and parsed bytes differ.'
}
Run-Control 'duplicate-json-rejected' {
    $path=[IO.Path]::Combine($WorkRoot,'duplicate.json')
    [IO.File]::WriteAllText($path,'{"field":1,"Field":2}',[Text.UTF8Encoding]::new($false))
    $rejected=$false;try { $null=Read-NativeAcceptanceJson -Path $path -MaximumBytes 1024 } catch { $rejected=$true }
    Require-Control $rejected 'Duplicate/case JSON keys were accepted.'
}
Run-Control 'overbound-json-rejected-before-read' {
    $path=[IO.Path]::Combine($WorkRoot,'overbound.json')
    [IO.File]::WriteAllText($path,('X'*1025),[Text.UTF8Encoding]::new($false))
    $rejected=$false;try { $null=Read-NativeAcceptanceJson -Path $path -MaximumBytes 1024 } catch { $rejected=$true }
    Require-Control $rejected 'Overbound JSON was accepted.'
}
Run-Control 'portable-vendor-utf16-path-bound' {
    $atLimit='C:\'+('a'*38)
    Require-Control ([IO.Path]::Combine($atLimit,'Calibre Portable').Length -eq 58) 'Authored limit fixture has the wrong length.'
    Require-Control (Test-NativeAcceptancePortableTargetLength $atLimit) 'Exact vendor path bound was rejected.'
    Require-Control (-not (Test-NativeAcceptancePortableTargetLength ($atLimit+'a'))) 'Overbound portable path was accepted.'
    $surrogate=[char]::ConvertFromUtf32(0x1F642)
    Require-Control (-not (Test-NativeAcceptancePortableTargetLength ($atLimit+$surrogate))) 'UTF16 surrogate units were not bounded.'
}
Run-Control 'existing-record-preserved' {
    $path=[IO.Path]::Combine($WorkRoot,'small.json'); $hash=(Get-NativeAcceptanceFileHash -Path $path -MaximumBytes 1024).sha256
    $rejected=$false;try { Write-NativeAcceptanceJson -Path $path -Record @{wrong='overwrite'} } catch { $rejected=$true }
    Require-Control ($rejected -and (Get-NativeAcceptanceFileHash -Path $path -MaximumBytes 1024).sha256 -ceq $hash) 'Existing record was overwritten.'
}

$script:normal=$null
Run-Control 'actual-plain-console-session-and-environment-restoration' {
    $prior=[Environment]::GetEnvironmentVariables('Process')
    $had=$prior.Contains('WBS_KIT_EXPAND');$old=$prior['WBS_KIT_EXPAND']
    try {
        [WinBookSplitNative.Directory]::SetConsoleFixtureEnvironment('WBS_KIT_EXPAND','Authored previous value')
        $code='import json,os,sys; a=[sys.stdin.readline().rstrip("\n") for _ in range(2)]; print(json.dumps({"answers":a,"environment":os.environ.get("WBS_KIT_EXPAND")},ensure_ascii=True))'
        $record=Invoke-NativeAcceptanceProcessCase -Path $TestPython -Arguments @('-I','-B','-X','utf8','-c',$code) -WorkingDirectory $WorkRoot -TimeoutSeconds 5 -Answers @('M','1,4,7') -Environment @{WBS_KIT_EXPAND='Authored temporary value'}
        Require-Control ($null -eq $record.failure) 'Plain console session raised a failure.'
        Assert-NativeAcceptanceProcess $record.probe
        Require-Control ($record.probe.ExitCode -eq 0 -and $record.environment_before_after_restored) 'Native session/restoration failed.'
        $body=[WinBookSplitSupport.StrictJson]::Parse($record.probe.Stdout.Trim())
        Require-Control ($body.answers.Count -eq 2 -and $body.answers[0] -ceq 'M' -and $body.answers[1] -ceq '1,4,7' -and $body.environment -ceq 'Authored temporary value') 'Authored plain answers/environment differ.'
        Require-Control ([Environment]::GetEnvironmentVariable('WBS_KIT_EXPAND','Process') -ceq 'Authored previous value') 'Previous nonempty environment value was not restored.'
        $script:normal=$record.probe
        Write-NativeAcceptanceJson -Path ([IO.Path]::Combine($WorkRoot,'plain-session.json')) -Record $record.probe
    } finally { $value=$null;if ($had) {$value=[string]$old};[WinBookSplitNative.Directory]::RestoreConsoleFixtureEnvironment('WBS_KIT_EXPAND',$value,$had) }
}
Run-Control 'empty-existing-environment-value-restored' {
    $prior=[Environment]::GetEnvironmentVariables('Process');$had=$prior.Contains('WBS_KIT_EXPAND');$old=$prior['WBS_KIT_EXPAND']
    try {
        [WinBookSplitNative.Directory]::SetConsoleFixtureEnvironment('WBS_KIT_EXPAND','')
        Require-Control ([Environment]::GetEnvironmentVariables('Process').Contains('WBS_KIT_EXPAND')) 'Win32 empty fixture value was not retained.'
        $code='import os; print(os.environ.get("WBS_KIT_EXPAND","missing"))'
        $record=Invoke-NativeAcceptanceProcessCase -Path $TestPython -Arguments @('-I','-B','-c',$code) -WorkingDirectory $WorkRoot -TimeoutSeconds 5 -Environment @{WBS_KIT_EXPAND='Authored temporary value'}
        Assert-NativeAcceptanceProcess $record.probe
        $after=[Environment]::GetEnvironmentVariables('Process')
        Require-Control ($record.environment_before_after_restored -and $after.Contains('WBS_KIT_EXPAND') -and $after['WBS_KIT_EXPAND'] -ceq '') 'Empty existing environment value was not restored.'
    } finally { $value=$null;if ($had) {$value=[string]$old};[WinBookSplitNative.Directory]::RestoreConsoleFixtureEnvironment('WBS_KIT_EXPAND',$value,$had) }
}
Run-Control 'absent-environment-variable-restored-as-absent' {
    $prior=[Environment]::GetEnvironmentVariables('Process');$had=$prior.Contains('WBS_CLI_HELP_APP');$old=$prior['WBS_CLI_HELP_APP']
    try{
        [WinBookSplitNative.Directory]::RestoreConsoleFixtureEnvironment('WBS_CLI_HELP_APP','',$false)
        $record=Invoke-NativeAcceptanceProcessCase $TestPython @('-I','-B','-c','import os; print(os.environ.get("WBS_CLI_HELP_APP"))') $WorkRoot 5 -Environment @{WBS_CLI_HELP_APP='Authored temporary helper path'}
        Assert-NativeAcceptanceProcess $record.probe
        Require-Control ($record.environment_before_after_restored -and -not [Environment]::GetEnvironmentVariables('Process').Contains('WBS_CLI_HELP_APP')) 'Absent variable became an empty present value.'
    }finally{$value=$null;if($had){$value=[string]$old};[WinBookSplitNative.Directory]::RestoreConsoleFixtureEnvironment('WBS_CLI_HELP_APP',$value,$had)}
}
Run-Control 'unknown-environment-name-rejected' {
    $old=[Environment]::GetEnvironmentVariable('PATH','Process');$rejected=$false
    try { [WinBookSplitNative.Directory]::SetConsoleFixtureEnvironment('PATH','Authored forbidden override') } catch { $rejected=$true }
    Require-Control ($rejected -and [Environment]::GetEnvironmentVariable('PATH','Process') -ceq $old) 'Unknown environment override was accepted.'
}
Run-Control 'unsafe-proof-and-bool-pid-rejected' {
    Require-Control ($null -ne $script:normal) 'Positive native control did not provide a proof fixture.'
    $positive=[WinBookSplitSupport.StrictJson]::Parse(($script:normal|ConvertTo-Json -Depth 20 -Compress))
    Assert-NativeAcceptanceProcess $positive
    foreach ($mutation in @(@{field='ParentStopped';value=$false},@{field='DescendantsStopped';value=$false},
        @{field='InputWriterStopped';value=$false},@{field='Pid';value=$true},@{field='StdoutTotalBytes';value=0})) {
        $clone=[WinBookSplitSupport.StrictJson]::Parse(($script:normal | ConvertTo-Json -Depth 20 -Compress))
        $clone[$mutation.field]=$mutation.value;$rejected=$false
        try { Assert-NativeAcceptanceProcess $clone } catch {$rejected=$true}
        Require-Control $rejected ('Unsafe mutated proof accepted: '+$mutation.field)
    }
    foreach($key in @('StartError','StopError','StreamError','ResultError','InteractionError','InputError')){
        $clone=[WinBookSplitSupport.StrictJson]::Parse(($script:normal|ConvertTo-Json -Depth 20 -Compress));$null=$clone.Remove($key)
        $rejected=$false;try{Assert-NativeAcceptanceProcess $clone}catch{$rejected=$true}
        Require-Control $rejected ('Absent error member accepted: '+$key)
    }
}
Run-Control 'safe-os-export-rejects-planted-raw-markers' {
    $os=@{Caption='C:\Users\PRIVATE_MARKER\Host';Version='10.0.19045';BuildNumber='19045';OSArchitecture='64-bit';ProductType=1;UBR=123;
        powershell_version='5.1.19041.1';process_bits=64;DisplayVersion='22H2';processor_architectures=@(9);os_bits=64;platform_scope='BLOCKED'}
    $safe=ConvertTo-NativeAcceptanceSafeOs $os
    Require-Control (($safe|ConvertTo-Json -Depth 8 -Compress) -notmatch 'PRIVATE_MARKER') 'Raw OS marker escaped redaction.'
    Require-Control ($null -eq $safe.Caption -and $safe.Version -ceq '10.0.19045') 'Safe OS observation was invented/dropped.'
}
foreach($kind in @('timeout','cancel')){
    Run-Control ('actual-owned-descendant-'+$kind) {
        $helper=[IO.Path]::Combine($RepositoryRoot,'tests','process','session_child.py')
        $control=Invoke-NativeAcceptanceTransportControl $TestPython $helper $WorkRoot $kind
        Require-Control ($control.status -ceq 'PASS' -and $control.scope -ceq 'CONTROLLED_NATIVE') 'Native descendant stop proof missing.'
    }
}
$after=@{};foreach ($path in $paths) { $after[$path]=(Get-NativeAcceptanceFileHash -Path $path -MaximumBytes 1048576).sha256 }
$unchanged=$true;foreach ($path in $paths) { if ($before[$path] -cne $after[$path]) {$unchanged=$false} }
$record=[ordered]@{scope='CONTROLLED_NATIVE_TOOLING_ONLY; no application/kit/Windows10 acceptance';
    started_utc=$started;finished_utc=[DateTime]::UtcNow.ToString('o');host_version=$PSVersionTable.PSVersion.ToString();
    process_bits=([IntPtr]::Size*8);source_before=$before;source_after=$after;source_unchanged=$unchanged;
    cases=$cases.ToArray();passed_count=@($cases | Where-Object status -eq PASS).Count;
    failed_count=@($cases | Where-Object status -ne PASS).Count;success=($unchanged -and @($cases | Where-Object status -ne PASS).Count -eq 0)}
Write-NativeAcceptanceJson -Path $ReportPath -Record $record
$record | ConvertTo-Json -Depth 12 -Compress
if (-not $record.success) {exit 1}
