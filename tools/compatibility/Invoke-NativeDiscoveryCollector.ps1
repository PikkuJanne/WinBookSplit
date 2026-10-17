# Native evidence transport only. Dot-sourcing neither selects nor installs Python.
# Load the maintained Process and native JSON helpers before calling this function.
function Invoke-NativeDiscoveryCollector {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$Path,
          [Parameter(Mandatory=$true)][AllowEmptyString()][AllowEmptyCollection()][string[]]$Arguments,
          [Parameter(Mandatory=$true)][string]$WorkingDirectory,
          [ValidateRange(0.05,600)][double]$TimeoutSeconds=30,
          [ValidateRange(1024,1048576)][int]$CaptureLimitBytes=1048576,
          [string]$OutputPath,
          [string]$ExpectedPackageFullName,[string]$ExpectedImagePath)
    $probe=$null; $collectorError=$null
    try {
        # CreateProcess and the existing job supervisor drain both byte streams.
        # Do not invoke a nested PowerShell through &, 2>&1, or a PS pipeline:
        # PS5.1 -EncodedCommand can otherwise reinterpret raw stderr as CLI XML.
        $parameters=@{Path=$Path;Arguments=$Arguments;WorkingDirectory=$WorkingDirectory;
            TimeoutSeconds=$TimeoutSeconds;CaptureLimitBytes=$CaptureLimitBytes;ResultLimitBytes=8388608}
        if (-not [string]::IsNullOrEmpty($ExpectedPackageFullName)) { $parameters.ExpectedPackageFullName=$ExpectedPackageFullName }
        if (-not [string]::IsNullOrEmpty($ExpectedImagePath)) { $parameters.ExpectedImagePath=$ExpectedImagePath }
        $probe=Invoke-WinBookSplitProcess @parameters
    } catch { $collectorError=$_.Exception.Message }
    $reason='complete';$status='FAIL';$exitCode=$null
    if ($null -eq $probe -or $null -ne $collectorError -or $null -ne $probe.StartError) { $reason='launch_failed' }
    else {
        $exitCode=$probe.ExitCode
        if ($probe.TimedOut) { $reason='timed_out' }
        elseif ($probe.Cancelled) { $reason='cancelled' }
        elseif (-not $probe.JobAssigned -or -not $probe.ParentStopped -or -not $probe.DescendantsStopped -or
                -not $probe.StreamsComplete -or -not $probe.InputWriterStopped -or
                $null -ne $probe.StopError -or $null -ne $probe.StreamError -or $null -ne $probe.ResultError -or
                $null -ne $probe.InteractionError -or $null -ne $probe.InputError) { $reason='supervision_failed' }
        elseif ($probe.StdoutTruncated -or $probe.StderrTruncated) { $reason='stream_limit' }
        elseif ($null -eq $exitCode -or $exitCode -ne 0) { $reason='child_failed' }
        else { $status='PASS' }
    }
    $record=[pscustomobject][ordered]@{protocol='winbooksplit.native-discovery-collector';schema_version=1;
        controller=[ordered]@{powershell_version=$PSVersionTable.PSVersion.ToString();edition=$PSVersionTable.PSEdition};
        command=(@($Path)+@($Arguments));cwd=$WorkingDirectory;timeout_seconds=$TimeoutSeconds;
        capture_limit_bytes=$CaptureLimitBytes;encoding='utf-8';status=$status;reason=$reason;
        exit_code=$exitCode;collector_error=$collectorError;probe=$probe}
    if (-not [string]::IsNullOrEmpty($OutputPath)) {
        # Guarded, bounded, CreateNew plain UTF-8 JSON; an existing receipt is never replaced.
        Write-NativeAcceptanceJson -Path $OutputPath -Record $record
    }
    return $record
}
