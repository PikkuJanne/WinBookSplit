[CmdletBinding()]
param([Parameter(Mandatory = $true)][string]$PayloadPath,
      [Parameter(Mandatory = $true)][string]$ReportPath)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$payload = [IO.File]::ReadAllText($PayloadPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
. $payload.helper
$neighborBefore = [IO.File]::ReadAllBytes($payload.neighbor)
$categories = @()
foreach ($case in $payload.cases) {
    if ($case.id -in @('invalid-mode', 'missing-arguments')) { continue }
    $result = ConvertFrom-SplitResult -Stdout $case.stdout -ExitCode $case.exit_code -Mode $case.mode
    $decisions = @()
    foreach ($choice in $payload.choices) {
        $decision = Get-SplitDecision -Result $result -Choice $choice
        $decisions += [ordered]@{ choice = $choice; message = $decision.message;
            fallback_modes = @($decision.fallback_modes); decision = $decision.decision; retry_mode = $decision.retry_mode }
    }
    $categories += [ordered]@{ id = $case.id; passed = $true; parsed_result = $result; decisions = $decisions }
}
$protocol = @()
foreach ($case in $payload.protocol_cases) {
    $rejected = $false
    try { $null = ConvertFrom-SplitResult -Stdout $case.stdout -ExitCode $case.exit_code -Mode $case.mode }
    catch { $rejected = $true }
    if (-not $rejected) { throw ('Invalid result was accepted: ' + $case.id) }
    $protocol += [ordered]@{ id = $case.id; passed = $true; rejected = $true }
}

function Invoke-OwnedNative {
    param([string[]]$Arguments)
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = $payload.python
    $start.Arguments = ($Arguments | ForEach-Object { ConvertTo-NativeArgument -Value $_ }) -join ' '
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.StandardOutputEncoding = [Text.Encoding]::UTF8
    $start.StandardErrorEncoding = [Text.Encoding]::UTF8
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $start
    try {
        $null = $process.Start()
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(30000)) { $process.Kill(); throw 'Owned native argument probe timed out.' }
        return [pscustomobject]@{ exit_code = $process.ExitCode; stdout = $stdout.GetAwaiter().GetResult();
                                 stderr = $stderr.GetAwaiter().GetResult() }
    }
    finally { $process.Dispose() }
}
$echo = Invoke-OwnedNative -Arguments (@('-I', '-B', '-c', 'import json,sys; print(json.dumps(sys.argv[1:],ensure_ascii=True))') + @($payload.native_arguments))
if ($echo.exit_code -ne 0 -or $echo.stderr.Length -ne 0) { throw 'Native argument echo failed.' }
$actualArguments = $echo.stdout | ConvertFrom-Json
if (($actualArguments | ConvertTo-Json -Compress) -cne ($payload.native_arguments | ConvertTo-Json -Compress)) {
    throw 'Native argument quoting changed literal data.'
}

# Load only the authored execution function, without app startup or Documents.
$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($payload.application, [ref]$tokens, [ref]$errors)
if (@($errors).Count -ne 0) { throw 'Application syntax errors.' }
$functions = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
                            $node.Name -ceq 'Run-PythonSplitter' }, $true))
if ($functions.Count -ne 1) { throw 'Exactly one trusted execution function is required.' }
. ([scriptblock]::Create($functions[0].Extent.Text))
$enginePath = $payload.flood_engine
$pythonExe = $payload.python
$InputFile = 'quote " [space] å & $(literal)\'
$outputDir = $payload.output + '\'
$logFile = $payload.log
$manualData = '1,4,7 "literal" %value%!data!'
$stream = [IO.File]::Open($logFile, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
$script:consoleLogWriter = New-Object IO.StreamWriter($stream, (New-Object Text.UTF8Encoding($false)))
$script:consoleLogWriter.AutoFlush = $true
try { $flood = Run-PythonSplitter -mode 'manual' -manualData $manualData 6>$null }
finally { $script:consoleLogWriter.Dispose() }
$echoed = $flood.message | ConvertFrom-Json
$expectedArguments = @($InputFile, $outputDir, 'manual', $manualData)
if (($echoed.arguments | ConvertTo-Json -Compress) -cne ($expectedArguments | ConvertTo-Json -Compress) -or
    -not $echoed.python_executable.Equals($payload.python, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Actual execution function changed arguments or used another Python.'
}
$log = [IO.File]::ReadAllText($logFile, [Text.Encoding]::UTF8)
if ($flood.exit_code -ne 1 -or $flood.code -cne 'invalid_start_pages' -or
    -not $log.Contains(('o' * 200000)) -or -not $log.Contains(('e' * 200000))) {
    throw 'Actual execution function did not preserve both large process streams/final result.'
}
$neighborAfter = [IO.File]::ReadAllBytes($payload.neighbor)
if ([Convert]::ToBase64String($neighborBefore) -cne [Convert]::ToBase64String($neighborAfter) -or
    @(Get-ChildItem -LiteralPath $payload.output -File).Count -ne 0) { throw 'Handler/probe changed output or neighbor.' }
$report = [ordered]@{
    passed = $true; exit_code = 0; host_major = $PSVersionTable.PSVersion.Major
    host_version = $PSVersionTable.PSVersion.ToString(); category_cases = $categories; protocol_cases = $protocol
    native_argument_probe = [ordered]@{ passed = $true; exit_code = $echo.exit_code; actual_arguments = [string[]]$actualArguments }
    stream_probe = [ordered]@{ passed = $true; exit_code = $flood.exit_code; stdout_length = 200000; stderr_length = 200000
        actual_run_function = $true; arguments_preserved = $true; python_executable = $echoed.python_executable }
    no_automatic_execution = $true; owned_neighbor_unchanged = $true
}
if (Test-Path -LiteralPath $ReportPath) { throw 'Report already exists.' }
$json = $report | ConvertTo-Json -Depth 100
[IO.File]::WriteAllText($ReportPath, $json, (New-Object Text.UTF8Encoding($false)))
