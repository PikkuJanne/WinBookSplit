[CmdletBinding()]
param([Parameter(Mandatory = $true)][string]$PayloadPath,
      [Parameter(Mandatory = $true)][string]$ReportPath)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
$payload = [IO.File]::ReadAllText($PayloadPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
. $payload.helper
. $payload.diagnostics
$clock = [Diagnostics.Stopwatch]::StartNew()
$cases = @()
foreach ($case in $payload.cases) {
    $cancel = $null
    try {
        $options = @{Path=$payload.python; Arguments=[string[]]$case.arguments;
            WorkingDirectory=$payload.work; TimeoutSeconds=[double]$case.timeout;
            CaptureLimitBytes=65536; ResultLimitBytes=[int]$case.result_limit}
        if ($case.cancel) {
            $cancel = New-Object Threading.CancellationTokenSource
            $cancel.CancelAfter(500)
            $options.CancellationToken = $cancel.Token
        }
        $result = Invoke-WinBookSplitProcess @options
        $parsed = $null
        $protocolError = $null
        if ($case.protocol) {
            try {
                if ($result.ResultError) { throw $result.ResultError }
                $parsed = ConvertFrom-SplitResult -Stdout ($result.ResultRecords -join "`n") -ExitCode $result.ExitCode -Mode 'manual'
            }
            catch { $protocolError = $_.Exception.Message }
        }
        $cases += [ordered]@{id=$case.id; kind=$case.kind; actual_process=$true;
            arguments=[string[]]$case.arguments; process=$result; parsed_result=$parsed; protocol_error=$protocolError}
    }
    finally { if ($null -ne $cancel) { $cancel.Dispose() } }
}
$report = [ordered]@{schema_version=1; host_version=$PSVersionTable.PSVersion.ToString();
    host_major=$PSVersionTable.PSVersion.Major; clr=[Environment]::Version.ToString();
    shell_executable=(Get-Process -Id $PID).Path; cases=$cases; elapsed_seconds=$clock.Elapsed.TotalSeconds;
    stored_policies=@(Get-ExecutionPolicy -List | Where-Object {$_.Scope -ne 'Process'} | ForEach-Object {
        [ordered]@{scope=$_.Scope.ToString();policy=$_.ExecutionPolicy.ToString()}})}
if ([IO.File]::Exists($ReportPath)) { throw 'Process report already exists.' }
[IO.File]::WriteAllText($ReportPath, ($report | ConvertTo-Json -Depth 100), (New-Object Text.UTF8Encoding($false)))
