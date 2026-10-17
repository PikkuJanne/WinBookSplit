[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$RepositoryRoot,
      [Parameter(Mandatory=$true)][string]$WorkRoot,
      [Parameter(Mandatory=$true)][string]$ReportPath,
      [Parameter(Mandatory=$true)][string]$ChildPowerShell)
$ErrorActionPreference='Stop'
$ProgressPreference='SilentlyContinue'
$PSModuleAutoloadingPreference='None'
Import-Module ([IO.Path]::Combine($PSHOME,'Modules','Microsoft.PowerShell.Utility','Microsoft.PowerShell.Utility.psd1')) -ErrorAction Stop
Import-Module ([IO.Path]::Combine($PSHOME,'Modules','Microsoft.PowerShell.Management','Microsoft.PowerShell.Management.psd1')) -ErrorAction Stop
. ([IO.Path]::Combine($RepositoryRoot,'engine','WinBookSplit.Diagnostics.ps1'))
. ([IO.Path]::Combine($RepositoryRoot,'engine','WinBookSplit.Process.ps1'))
. ([IO.Path]::Combine($RepositoryRoot,'engine','WinBookSplit.Logging.ps1'))
. ([IO.Path]::Combine($RepositoryRoot,'engine','WinBookSplit.Support.ps1'))
. ([IO.Path]::Combine($RepositoryRoot,'tools','compatibility','Invoke-NativeAcceptance.ps1'))
. ([IO.Path]::Combine($RepositoryRoot,'tools','compatibility','Invoke-NativeDiscoveryCollector.ps1'))
Initialize-WinBookSplitLogTypes
Initialize-WinBookSplitSupport
Initialize-NativeAcceptanceDirectory
[WinBookSplitNative.Directory]::CreateExclusive($WorkRoot)
$cases=New-Object 'Collections.Generic.List[object]'
$before=[Environment]::GetEnvironmentVariables('Process')
function Require-Collector($Value,[string]$Reason) { if (-not $Value) { throw $Reason } }
function Run-CollectorControl([string]$Id,[scriptblock]$Action) {
    $row=[ordered]@{id=$Id;status='FAIL';error=$null;scope='CONTROLLED_NATIVE_COLLECTOR_ONLY'}
    try { & $Action; $row.status='PASS' } catch { $row.error=$_.Exception.Message }
    $cases.Add([pscustomobject]$row)
}
function New-CollectorArguments([string]$Text) {
    # Text transport is explicit on both editions; authored scripts write their
    # actual bytes to Console, never PowerShell's serialized ErrorRecord stream.
    $prefix='$ErrorActionPreference="Stop";$ProgressPreference="SilentlyContinue";[Console]::OutputEncoding=New-Object Text.UTF8Encoding($false);'
    return @('-NoLogo','-NoProfile','-NonInteractive','-InputFormat','Text','-OutputFormat','Text',
        '-ExecutionPolicy','RemoteSigned','-EncodedCommand',[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($prefix+$Text)))
}
function Invoke-CollectorControl([string]$Name,[string]$Text,[double]$Timeout=15,[int]$Capture=1048576) {
    return Invoke-NativeDiscoveryCollector -Path $ChildPowerShell -Arguments (New-CollectorArguments $Text) `
        -WorkingDirectory $WorkRoot -TimeoutSeconds $Timeout -CaptureLimitBytes $Capture `
        -OutputPath ([IO.Path]::Combine($WorkRoot,$Name+'.json'))
}
$unicode=([char]0x65e5).ToString()+[char]0x672c+[char]0xe4
Run-CollectorControl 'dual-unicode-streams-and-plain-json-positive-control' {
    $text='[Console]::Out.Write("AUTHORED_stdout_'+$unicode+'");[Console]::Error.Write("AUTHORED_stderr_'+$unicode+'");exit 0'
    $record=Invoke-CollectorControl 'unicode' $text
    Require-Collector ($record.status -ceq 'PASS' -and $record.exit_code -eq 0 -and $record.reason -ceq 'complete') 'The actual collector positive control did not pass.'
    Assert-NativeAcceptanceProcess $record.probe
    Require-Collector ($record.probe.Stdout -ceq ('AUTHORED_stdout_'+$unicode) -and $record.probe.Stderr -ceq ('AUTHORED_stderr_'+$unicode)) 'Unicode stream bytes did not roundtrip.'
    $path=[IO.Path]::Combine($WorkRoot,'unicode.json');$bytes=[IO.File]::ReadAllBytes($path)
    Require-Collector ($bytes.Length -gt 3 -and $bytes[0] -eq 123) 'Receipt has a BOM or is not plain JSON.'
    $read=Read-NativeAcceptanceJson -Path $path
    Require-Collector ($read.value.protocol -ceq 'winbooksplit.native-discovery-collector' -and $read.value.schema_version -eq 1 -and
        $read.value.probe.Stderr -ceq ('AUTHORED_stderr_'+$unicode)) 'Strict UTF8 receipt roundtrip differs.'
}
Run-CollectorControl 'nonzero-listing-retains-both-streams' {
    $record=Invoke-CollectorControl 'nonzero' '[Console]::Out.Write("AUTHORED_LIST_REACHED");[Console]::Error.Write("AUTHORED_LIST_FAILURE");exit 23'
    Assert-NativeAcceptanceProcess $record.probe
    Require-Collector ($record.status -ceq 'FAIL' -and $record.reason -ceq 'child_failed' -and $record.exit_code -eq 23 -and
        $record.probe.Stdout -ceq 'AUTHORED_LIST_REACHED' -and $record.probe.Stderr -ceq 'AUTHORED_LIST_FAILURE') 'Listing failure was hidden or its stub was not reached.'
}
Run-CollectorControl 'missing-launch-is-explicit-failure' {
    $record=Invoke-NativeDiscoveryCollector -Path ([IO.Path]::Combine($WorkRoot,'authored-missing.exe')) -Arguments @() -WorkingDirectory $WorkRoot `
        -OutputPath ([IO.Path]::Combine($WorkRoot,'launch-failure.json'))
    Require-Collector ($record.status -ceq 'FAIL' -and $record.reason -ceq 'launch_failed' -and $null -eq $record.exit_code -and
        $null -ne $record.probe.StartError -and $record.probe.ParentStopped -and $record.probe.DescendantsStopped) 'Launch failure was not recorded explicitly.'
}
Run-CollectorControl 'timeout-stops-child-after-positive-reach-marker' {
    $record=Invoke-CollectorControl 'timeout' '[Console]::Out.Write("AUTHORED_TIMEOUT_REACHED");Start-Sleep -Seconds 60' 5
    Assert-NativeAcceptanceProcess $record.probe -ExpectedStop timeout
    Require-Collector ($record.status -ceq 'FAIL' -and $record.reason -ceq 'timed_out' -and
        $record.probe.Stdout -ceq 'AUTHORED_TIMEOUT_REACHED' -and $record.probe.ElapsedSeconds -lt 9) 'Actual timeout/positive reach/finite shutdown was not proved.'
}
function New-CollectorDescendant([string]$Seconds) {
    $daughter=New-CollectorArguments ('[Console]::Out.Write("AUTHORED_DESCENDANT_REACHED");[Console]::Error.Write("AUTHORED_DESCENDANT_STDERR");Start-Sleep -Seconds '+$Seconds+';exit 0')
    $joined=($daughter|ForEach-Object {ConvertTo-NativeArgument $_}) -join ' '
    # Both strings come from this controlled fixture, encoded literally. No
    # application path, book text, or runtime-selection input is evaluated.
    $exe=$ChildPowerShell.Replace("'","''");$argv=$joined.Replace("'","''")
    return ('$s=New-Object Diagnostics.ProcessStartInfo;$s.FileName='''+$exe+''';$s.Arguments='''+$argv+''';$s.UseShellExecute=$false;$p=[Diagnostics.Process]::Start($s);[Console]::Out.Write("AUTHORED_PARENT_REACHED");exit 0')
}
Run-CollectorControl 'descendant-inherited-stream-positive-control' {
    $record=Invoke-CollectorControl 'descendant-positive' (New-CollectorDescendant '0')
    Assert-NativeAcceptanceProcess $record.probe
    Require-Collector ($record.status -ceq 'PASS' -and $record.probe.Stdout.Contains('AUTHORED_PARENT_REACHED') -and
        $record.probe.Stdout.Contains('AUTHORED_DESCENDANT_REACHED') -and $record.probe.Stderr -ceq 'AUTHORED_DESCENDANT_STDERR') 'Inherited-stream descendant fixture was not reached.'
}
Run-CollectorControl 'hung-descendant-and-inherited-streams-are-bounded' {
    $record=Invoke-CollectorControl 'descendant-hung' (New-CollectorDescendant '60') 7
    Assert-NativeAcceptanceProcess $record.probe -ExpectedStop timeout
    Require-Collector ($record.status -ceq 'FAIL' -and $record.reason -ceq 'timed_out' -and $record.probe.ExitCode -eq 0 -and
        $record.probe.Stdout.Contains('AUTHORED_PARENT_REACHED') -and $record.probe.Stdout.Contains('AUTHORED_DESCENDANT_REACHED') -and
        $record.probe.Stderr -ceq 'AUTHORED_DESCENDANT_STDERR' -and $record.probe.ElapsedSeconds -lt 11) 'Hung descendant/stream fixture did not prove finite owned shutdown.'
}
Run-CollectorControl 'bounded-output-cannot-become-success' {
    $record=Invoke-CollectorControl 'stream-limit' '[Console]::Out.Write(("A"*4096));[Console]::Error.Write(("B"*4096));exit 0' 15 1024
    Require-Collector ($record.status -ceq 'FAIL' -and $record.reason -ceq 'stream_limit' -and $record.exit_code -eq 0 -and
        $record.probe.StdoutTruncated -and $record.probe.StderrTruncated -and $record.probe.StreamsComplete) 'Truncated stream was accepted or fixture did not run.'
}
Run-CollectorControl 'existing-receipt-is-preserved' {
    $path=[IO.Path]::Combine($WorkRoot,'unicode.json');$hash=(Get-NativeAcceptanceFileHash $path -MaximumBytes 1048576).sha256
    $rejected=$false
    try { $null=Invoke-NativeDiscoveryCollector -Path $ChildPowerShell -Arguments (New-CollectorArguments 'exit 0') `
        -WorkingDirectory $WorkRoot -OutputPath $path } catch { $rejected=$true }
    Require-Collector ($rejected -and (Get-NativeAcceptanceFileHash $path -MaximumBytes 1048576).sha256 -ceq $hash) 'Collector replaced an existing receipt.'
}
Run-CollectorControl 'application-environment-is-unchanged' {
    $after=[Environment]::GetEnvironmentVariables('Process')
    foreach ($name in @('PATH','LOCALAPPDATA','VIRTUAL_ENV','PYTHONPATH','PYTHONHOME','WBS_PYTHON','WBS_PYTHONPATH')) {
        Require-Collector ($before.Contains($name) -eq $after.Contains($name) -and $before[$name] -ceq $after[$name]) ('Collector changed process environment: '+$name)
    }
}
$failures=@($cases|Where-Object {$_.status -cne 'PASS'})
$report=[ordered]@{protocol='winbooksplit.native-discovery-collector-tests';schema_version=1;scope='CONTROLLED_NATIVE_COLLECTOR_ONLY';
    controller_version=$PSVersionTable.PSVersion.ToString();controller_edition=$PSVersionTable.PSEdition;child_powershell=$ChildPowerShell;
    app_python_required=$false;status=$(if($failures.Count -eq 0){'PASS'}else{'FAIL'});cases=@($cases.ToArray())}
Write-NativeAcceptanceJson -Path $ReportPath -Record $report
if ($failures.Count -ne 0) { exit 1 }
exit 0
