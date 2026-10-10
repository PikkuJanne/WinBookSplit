<#
.SYNOPSIS
Split a local PDF, EPUB or AZW3 into chapter PDFs without changing its source.
.DESCRIPTION
Preserves every physical PDF page using the existing Python/pypdf planner.
Without Mode, select a method interactively. Scripted calls supply Mode plus
BookmarkLevel or StartPages. NonInteractive and Preview require complete choices.
EPUB/AZW3 analysis uses Calibre in an owned temporary workspace, even for Preview.
Requires regular Windows x64 CPython 3.14.8 and pypdf 6.19.0; ebooks also require
Calibre 9.15.0. See docs/codex-v1.0.0/SUPPORT_AND_SETUP.md for explicit setup.
.PARAMETER InputFile
One literal local PDF, EPUB or AZW3 source file. Missing interactive input is
prompted; blank or C cancels. Required for NonInteractive and Preview.
.PARAMETER OutputDirectory
Existing output base. Each execution publishes a unique new child. Defaults to Documents.
.PARAMETER Mode
Auto or Manual. Auto requires BookmarkLevel; Manual requires StartPages.
.PARAMETER BookmarkLevel
1 or 2, for Auto only. Level 2 children stay within their Level 1 parents.
.PARAMETER StartPages
Comma-separated ASCII decimal physical PDF start pages, for Manual only.
Every token must be valid. Starts are sorted/deduplicated and page 1 is included.
.PARAMETER Preview
Print the validated plan and exit with no chapter PDFs or console log files.
Ebooks require temporary conversion, safely cleaned before the plan returns.
Cannot combine with KeepConvertedPdf. Preview never prompts or retries another mode.
.PARAMETER NonInteractive
Never prompt, pause, clear the console, open dialogs/Explorer or attempt a fallback.
Supply InputFile, Mode and the selected level or starts; no usable plan returns 5.
.PARAMETER NoPause
Suppress exit pauses; interactive method, plan confirmation and open choices still apply.
.PARAMETER Version
Print the canonical development version and exit 0 without input or dependencies.
Cannot combine with processing parameters.
.PARAMETER PythonPath
Explicit trusted absolute Python executable. Otherwise use validated discovery.
.PARAMETER CalibrePath
Explicit trusted absolute ebook-convert executable; PDF input never probes Calibre.
.PARAMETER KeepConvertedPdf
Retain the exact full converted PDF inside successful ebook output only.
.PARAMETER ConversionTimeout
Converter deadline in seconds, 1 through 86400; default 1800.
.PARAMETER ProcessTimeout
Engine/tree/stream deadline in seconds, 1 through 172800; default at least 3600.
.EXAMPLE
.\WinBookSplit.ps1 -InputFile 'C:\Books\Manual.pdf'
.EXAMPLE
.\WinBookSplit.ps1 -InputFile 'C:\Books\Manual.pdf' -Mode Auto -BookmarkLevel 2 -Preview -NonInteractive
.EXAMPLE
.\WinBookSplit.ps1 -InputFile 'C:\Books\Manual.pdf' -Mode Manual -StartPages '1,4,7' -NonInteractive
.EXAMPLE
.\WinBookSplit.ps1 -InputFile 'C:\Books\Book.epub' -Mode Auto -BookmarkLevel 1 -KeepConvertedPdf -NonInteractive
.EXAMPLE
.\WinBookSplit.ps1 -Version
.NOTES
Application exits: 0 execution/preview/version, 2 invalid input/arguments/path,
3 dependency, 4 conversion, 5 no requested bookmark plan, 6 PDF/output failure,
7 unsupported features, 130 cancellation/timeout. Native PowerShell invocation
errors before this script runs use the host's own exit status.
Author: Janne Vuorela. MIT; unsigned application. Windows 11 x64, PS5.1/PS7.
#>

param (
    [string]$InputFile,
    [string]$OutputDirectory,
    [string]$CalibrePath,
    [switch]$KeepConvertedPdf,
    [ValidateRange(1, 86400)][int]$ConversionTimeout = 1800,
    [string]$PythonPath,
    [ValidateRange(1, 172800)][int]$ProcessTimeout = [Math]::Max(3600, $ConversionTimeout + 1800),
    [string]$Mode,
    [string]$BookmarkLevel,
    [string]$StartPages,
    [switch]$Preview,
    [switch]$NonInteractive,
    [switch]$NoPause,
    [switch]$Version
)

# Machine outcomes and path diagnostics use UTF-8 through every launcher.
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
[Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)

# --- Configuration ---
$AppName = "WinBookSplit"
$versionContract = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'engine\WinBookSplit.Outcomes.json') -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
$ver = $versionContract.application_version
if ($ver -isnot [string] -or $ver -cnotmatch '^\d+\.\d+\.\d+(?:-[a-z0-9.]+)?$') { throw 'Invalid shipped application version.' }
# Version precedes every input check, helper load and runtime probe.
if ($Version -and $PSBoundParameters.Count -eq 1 -and $args.Count -eq 0) {
    Write-Output ($AppName + ' ' + $ver)
    exit 0
}
$documentsPath = [Environment]::GetFolderPath("MyDocuments")

# --- UI Functions ---
function Draw-Header {
    if (-not $NonInteractive -and -not $Preview -and -not [Console]::IsOutputRedirected) { Clear-Host }
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "      $AppName v$ver" -ForegroundColor Yellow
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
}

. (Join-Path $PSScriptRoot 'engine\WinBookSplit.Paths.ps1')
. (Join-Path $PSScriptRoot 'engine\WinBookSplit.Diagnostics.ps1')
. (Join-Path $PSScriptRoot 'engine\WinBookSplit.Runtime.ps1')
. (Join-Path $PSScriptRoot 'engine\WinBookSplit.Process.ps1')
. (Join-Path $PSScriptRoot 'engine\WinBookSplit.Logging.ps1')

function New-WinBookSplitFailure {
    param([string]$Code, [string]$Message)
    $failure = New-Object InvalidOperationException($Message)
    $failure.Data['Code'] = $Code
    return $failure
}

function New-WinBookSplitOutcome {
    param([string]$Status, [string]$Code, [string]$Message, [string]$Mode = '', $EngineResult = $null)
    $count = 0
    $finalDirectory = $null
    if ($null -ne $EngineResult -and $null -ne $EngineResult.execution) {
        $finalDirectory = $EngineResult.execution.final_directory
        if ($Status -ceq 'success') { $count = $EngineResult.written_count }
    }
    $exitCode = Get-SplitExitCode -Code $Code
    if ($null -ne $EngineResult -and $EngineResult.code -ceq $Code) { $exitCode = $EngineResult.exit_code }
    return [pscustomobject][ordered]@{ protocol = 'winbooksplit.outcome'; version = 1;
        status = $Status; code = $Code; message = $Message; exit_code = $exitCode;
        mode = $Mode; written_count = $count; final_directory = $finalDirectory; engine_result = $EngineResult }
}

function Complete-WinBookSplitConsoleLog {
    param([Parameter(Mandatory = $true)]$Outcome)
    $script:closedConsoleLogWriter = $script:consoleLogWriter
    $script:consoleLogWriter.WriteFinal('[OPERATION-OUTCOME] ' + ($Outcome | ConvertTo-Json -Depth 100 -Compress))
    $script:consoleLogWriter.Flush()
    $script:consoleLogWriter.Dispose()
    $script:consoleLogWriter = $null
    if (-not $Preview) {
        $manifest = New-WinBookSplitRunRecord -Outcome $Outcome
        Write-WinBookSplitRunManifest -Directory $consoleDir -Manifest $manifest -Leases $script:consoleRecordLeases
    }
}

function New-WinBookSplitRunRecord {
    param([Parameter(Mandatory = $true)]$Outcome)
    $engine = $Outcome.engine_result
    $plan = $null
    if ($null -ne $engine -and $engine.PSObject.Properties.Name -contains 'plan' -and $null -ne $engine.plan) { $plan = $engine.plan }
    elseif ($null -ne $script:confirmedPlan) { $plan = $script:confirmedPlan }
    $identity = $script:sourceObservation
    $normalized = $null
    if ($null -ne $plan) {
        $identity = $plan.source_identity
        if ($plan.PSObject.Properties.Name -contains 'original_ebook_identity') { $identity = $plan.original_ebook_identity }
        $normalized = $plan.normalized_inputs
    }
    $planning = @(); $parser = @(); $suppressed = 0; $truncated = 0
    if ($null -ne $engine) {
        $planning = @($engine.warnings)
        if ($null -ne $engine.diagnostic -and $engine.diagnostic.PSObject.Properties.Name -contains 'parser_warnings') {
            $parser = @($engine.diagnostic.parser_warnings.records)
            $suppressed = $engine.diagnostic.parser_warnings.suppressed_count
            $truncated = $engine.diagnostic.parser_warnings.message_truncated_count
        }
    }
    $calibreVersion = $null
    if ($null -ne $converter) { $calibreVersion = $converter.Version }
    return [pscustomobject][ordered]@{
        protocol = 'winbooksplit.run'; version = 1; run_id = $consoleRunId; application_version = $ver
        started_utc = $script:runStartedUtc; finished_utc = [DateTime]::UtcNow.ToString('o')
        diagnostics_finalized = $true
        runtime_versions = [ordered]@{ powershell = $PSVersionTable.PSVersion.ToString(); python = $runtime.Version;
            pypdf = $runtime.PypdfVersion; calibre = $calibreVersion }
        settings = [ordered]@{ mode = $Outcome.mode; input_kind = $inputExt.Substring(1); preview = [bool]$Preview;
            non_interactive = [bool]$NonInteractive; no_pause = [bool]$NoPause; keep_converted_pdf = [bool]$KeepConvertedPdf;
            conversion_timeout = $ConversionTimeout; process_timeout = $ProcessTimeout; normalized_inputs = $normalized }
        source_identity = $identity; plan = $plan; engine_result = $engine; outcome = $Outcome
        warnings = [ordered]@{ planning = $planning; parser = $parser; parser_suppressed_count = $suppressed;
            parser_message_truncated_count = $truncated }
        log = [ordered]@{ filename = 'console.log'; max_bytes = $script:closedConsoleLogWriter.MaxBytes;
            body_limit_bytes = $script:closedConsoleLogWriter.BodyLimitBytes;
            bytes_written = $script:closedConsoleLogWriter.BytesWritten; limit_reached = $script:closedConsoleLogWriter.LimitReached }
    }
}

function Wait-WinBookSplitExit {
    if (-not $NonInteractive -and -not $Preview -and -not $NoPause) {
        try { $null = Read-WinBookSplitSessionLine $null 'Press Enter to exit' }
        catch { Write-Host ('Cannot pause the console: ' + (ConvertTo-WinBookSplitDisplayText $_.Exception.Message)) -ForegroundColor Yellow }
    }
}

function ConvertTo-WinBookSplitDisplayText {
    param([AllowNull()][string]$Text)
    # Paths and bookmarks are data: controls cannot become console records.
    # Only the human rendering changes; the immutable plan keeps its raw text.
    return [regex]::Replace([string]$Text, '[\x00-\x1f\x7f-\x9f\u2028\u2029]', {
        param($match)
        '\u{0:x4}' -f [int][char]$match.Value
    })
}

function Show-WinBookSplitPlan {
    param($Plan, [bool]$PreviewOnly = $true, $WorkingPdf = $null)
    Write-Host 'Split plan:' -ForegroundColor Cyan
    $originalPath = $Plan.source_identity.path
    if ($null -ne $Plan.original_ebook_identity) { $originalPath = $Plan.original_ebook_identity.path }
    Write-Host ('Source: ' + (ConvertTo-WinBookSplitDisplayText $originalPath))
    if ($null -ne $Plan.conversion) {
        if ($null -ne $WorkingPdf) {
            Write-Host ('Working generated PDF: ' + (ConvertTo-WinBookSplitDisplayText $WorkingPdf.path))
            Write-Host 'The working PDF remains available through page selection and confirmation.'
        }
        else {
            Write-Host ('Generated PDF: ' + (ConvertTo-WinBookSplitDisplayText $Plan.conversion.generated_pdf_identity.path))
            Write-Host 'This conversion path was cleaned; the captured PDF remains in memory.'
        }
        Write-Host 'Ranges use physical pages in the generated PDF, not ebook locations.'
        if ($Plan.PSObject.Properties.Name -contains 'keep_converted_pdf' -and $Plan.keep_converted_pdf) { Write-Host 'Full generated PDF: retained with successful chapter output.' }
        else { Write-Host 'Full generated PDF: not retained after this operation.' }
    }
    Write-Host ('Physical PDF pages: ' + $Plan.total_pages)
    $method = 'Manual'
    if ($Plan.mode -cne 'manual') { $method = 'Auto, bookmark level ' + $Plan.mode }
    Write-Host ('Method: ' + $method)
    Write-Host ('Output base: ' + (ConvertTo-WinBookSplitDisplayText $Plan.output_naming.resolved_base))
    foreach ($entry in $Plan.entries) {
        Write-Host ('{0}. {1} | physical pages {2}-{3} | {4} | {5}' -f $entry.sequence,
            (ConvertTo-WinBookSplitDisplayText $entry.title), ($entry.start + 1), $entry.end,
            (ConvertTo-WinBookSplitDisplayText $entry.filename), (ConvertTo-WinBookSplitDisplayText $entry.reason))
        if ($null -ne $entry.parent_id) { Write-Host ('   Parent: ' + (ConvertTo-WinBookSplitDisplayText $entry.parent_id)) }
    }
    foreach ($notice in $Plan.notices) { Write-Host ('Notice: ' + (ConvertTo-WinBookSplitDisplayText $notice)) }
    foreach ($warning in $Plan.warnings) {
        Write-Host ('Warning: ' + (ConvertTo-WinBookSplitDisplayText $warning.code) + ': ' +
            (ConvertTo-WinBookSplitDisplayText $warning.message)) -ForegroundColor Yellow
    }
    Write-Host ('Coverage: every physical page exactly once; {0} pages, {1} sections.' -f
        $Plan.coverage.covered_pages, $Plan.coverage.section_count)
    if ($PreviewOnly) { Write-Host 'Preview only: no chapter PDFs were written.' }
    else { Write-Host 'No chapter PDFs have been written. Confirm this plan to create the sections above.' }
}

function Read-WinBookSplitSessionLine {
    param($Session, [string]$Prompt)
    Write-Host ($Prompt + ': ') -NoNewline
    $pending = Start-WinBookSplitConsoleLine
    while (-not $pending.Completed) {
        if ($null -ne $Session -and $Session.Completed) {
            $pending.AbandonForExit()
            $script:abandonedConsoleInput = $true
            throw (New-WinBookSplitFailure 'processor_cancelled' 'The engine ended while awaiting console input.')
        }
        Start-Sleep -Milliseconds 20
    }
    Write-Host ''
    if ($pending.Error) { throw (New-WinBookSplitFailure 'processor_protocol_failed' ('Cannot read console input: ' + $pending.Error)) }
    return $pending.Value
}

function Read-WinBookSplitManualStarts {
    param($Session, $Event, [string]$SuppliedStarts)
    Write-Host ('Physical PDF pages: ' + $Event.total_pages)
    if ($inputExt -ne '.pdf') { Write-Host 'Use physical page numbers in the generated PDF, not ebook locations.' -ForegroundColor Yellow }
    if ($Event.manual_starts_supplied) { return $SuppliedStarts }
    Write-Host 'Enter page numbers where new files should START.' -ForegroundColor Yellow
    return Read-WinBookSplitSessionLine $Session 'Pages (comma separated; blank or C cancels)'
}

function Assert-WinBookSplitWorkingPdf {
    param($Event)
    $working = $Event.working_pdf
    Assert-SplitPreviewSourceIdentity $Event.source_identity 'reader_snapshot'
    if ($null -eq $working -or $working.binding -cne 'working_pdf_copy' -or
        $working.run_id -isnot [string] -or $working.run_id -cnotmatch '^[0-9a-f]{32}$' -or
        $working.path -isnot [string] -or $working.owned_directory -isnot [string] -or
        -not [StringComparer]::OrdinalIgnoreCase.Equals($working.owned_directory,
            [IO.Path]::Combine($outputDir, '.WinBookSplit-stage-' + $working.run_id)) -or
        -not [StringComparer]::OrdinalIgnoreCase.Equals($working.path,
            [IO.Path]::Combine($working.owned_directory, 'WinBookSplit_Working.pdf')) -or
        -not (Test-SplitInteger $Event.total_pages) -or $Event.total_pages -lt 1 -or
        -not (Test-SplitInteger $working.page_count) -or $working.page_count -ne $Event.total_pages -or
        $working.sha256 -cne $Event.source_identity.sha256 -or
        $working.size_bytes -ne $Event.source_identity.size_bytes) { throw 'Invalid owned working PDF identity.' }
    $null = Resolve-WinBookSplitOutputBase -Path $working.owned_directory
    $item = Get-Item -LiteralPath $working.path -ErrorAction Stop
    if ($item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -or
        $item.Length -ne $working.size_bytes) {
        throw 'The requested working PDF differs from the captured source.'
    }
    $stream, $sha = $null, $null
    try {
        $stream = [IO.File]::Open($working.path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        $sha = [Security.Cryptography.SHA256]::Create()
        $actualHash = [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-', '').ToLowerInvariant()
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
        if ($null -ne $sha) { $sha.Dispose() }
    }
    if ($actualHash -cne $working.sha256) { throw 'The requested working PDF differs from the captured source.' }
    return $working
}

function Invoke-WinBookSplitInteractiveEngine {
    param([string[]]$Arguments, [string]$SessionId, [string]$InitialMode, [string]$SuppliedStarts)
    $session = Start-WinBookSplitProcessSession -Path $pythonExe -Arguments $Arguments `
        -WorkingDirectory ([IO.Path]::GetDirectoryName($enginePath)) -TimeoutSeconds $ProcessTimeout
    $sequence, $currentMode, $nextStage = 0, $InitialMode, 'initial'
    $workingPdf = $null
    $capturedSource, $capturedPages = $null, $null
    $script:confirmedPlan, $script:interactionFailure = $null, $null
    $script:interactionStopRequested = $false
    $script:noPlanCancelled = $false
    try {
        while (-not $session.Completed) {
            $raw = $session.TryReadRequest()
            if ($null -eq $raw) { Start-Sleep -Milliseconds 20; continue }
            $sequence++
            $event = ConvertFrom-WinBookSplitInteraction $raw $SessionId $sequence $currentMode
            $reply = [ordered]@{ protocol = 'winbooksplit.interaction'; version = 1;
                session = $SessionId; sequence = $sequence; action = 'cancel' }
            switch -CaseSensitive ($event.stage) {
                'input_ready' {
                    if ($currentMode -cne 'manual' -or $nextStage -cne 'initial' -or
                        $event.can_request_working_pdf -isnot [bool] -or
                        $event.can_request_working_pdf -ne ($inputExt -ne '.pdf') -or
                        $event.manual_starts_supplied -isnot [bool] -or
                        $event.manual_starts_supplied -ne (-not [string]::IsNullOrEmpty($SuppliedStarts)) -or
                        -not (Test-SplitInteger $event.total_pages) -or $event.total_pages -lt 1) {
                        throw 'Unexpected manual preparation event.'
                    }
                    Assert-SplitPreviewSourceIdentity $event.source_identity 'reader_snapshot'
                    if ($inputExt -eq '.pdf' -and
                        -not [StringComparer]::OrdinalIgnoreCase.Equals($event.source_identity.path, $InputFile)) {
                        throw 'The prepared PDF differs from the selected input.'
                    }
                    $capturedSource, $capturedPages = $event.source_identity, $event.total_pages
                    $open = 'N'
                    if ($event.can_request_working_pdf) {
                        Assert-SplitPreviewSourceIdentity $event.original_ebook_identity 'ebook_snapshot'
                        if (-not [StringComparer]::OrdinalIgnoreCase.Equals($event.original_ebook_identity.path, $InputFile)) {
                            throw 'The manual ebook identity differs from the selected input.'
                        }
                        Write-Host ('Converted PDF: ' + $event.total_pages + ' physical pages. Use these numbers, not ebook locations.') -ForegroundColor Yellow
                        do {
                            $open = Read-WinBookSplitSessionLine $session 'Open the generated PDF? [O] open, [N] continue, [C] cancel'
                            if ($null -eq $open -or [string]::IsNullOrWhiteSpace($open)) { $open = 'C' }
                            $open = $open.Trim().ToUpperInvariant()
                            if ($open -cnotin @('O', 'N', 'C')) { Write-Host 'Choose exactly O, N or C.' }
                        } while ($open -cnotin @('O', 'N', 'C'))
                    }
                    if ($open -ceq 'O') { $reply.action = 'request_working_pdf'; $nextStage = 'working' }
                    elseif ($open -ceq 'N') {
                        $starts = Read-WinBookSplitManualStarts $session $event $SuppliedStarts
                        if (-not [string]::IsNullOrWhiteSpace($starts) -and $starts.Trim() -ine 'C') {
                            $reply.action = 'starts'; $reply.starts = $starts; $nextStage = 'plan'
                        }
                    }
                }
                'working_pdf_ready' {
                    if ($nextStage -cne 'working' -or $currentMode -cne 'manual') { throw 'Unexpected working PDF event.' }
                    if ($event.total_pages -ne $capturedPages -or
                        -not (Test-WinBookSplitJsonEqual $event.source_identity $capturedSource)) {
                        throw 'The working PDF source changed after manual preparation.'
                    }
                    $workingPdf = Assert-WinBookSplitWorkingPdf $event
                    Write-Host ('Working generated PDF: ' + (ConvertTo-WinBookSplitDisplayText $workingPdf.path))
                    try { Invoke-Item -LiteralPath $workingPdf.path -ErrorAction Stop }
                    catch { Write-Host ('Cannot open the PDF viewer: ' + (ConvertTo-WinBookSplitDisplayText $_.Exception.Message)) -ForegroundColor Yellow }
                    $event | Add-Member -NotePropertyName manual_starts_supplied -NotePropertyValue (-not [string]::IsNullOrEmpty($SuppliedStarts))
                    $starts = Read-WinBookSplitManualStarts $session $event $SuppliedStarts
                    if (-not [string]::IsNullOrWhiteSpace($starts) -and $starts.Trim() -ine 'C') {
                        $reply.action = 'starts'; $reply.starts = $starts; $nextStage = 'plan'
                    }
                }
                'plan_ready' {
                    if (($currentMode -ceq 'manual' -and $nextStage -cne 'plan') -or
                        ($currentMode -cne 'manual' -and $nextStage -cne 'initial')) { throw 'Unexpected plan confirmation event.' }
                    $plan = Assert-WinBookSplitInteractionPlan $event $currentMode $InputFile $outputDir ([bool]$KeepConvertedPdf)
                    if ($null -ne $capturedSource -and
                        (-not (Test-WinBookSplitJsonEqual $plan.source_identity $capturedSource) -or
                         $plan.total_pages -ne $capturedPages)) { throw 'The plan differs from the captured manual PDF.' }
                    Show-WinBookSplitPlan -Plan $plan -PreviewOnly $false -WorkingPdf $workingPdf
                    $script:consoleLogWriter.WriteLine('[PLAN] ' + $raw)
                    do {
                        $choice = Read-WinBookSplitSessionLine $session 'Create these chapter PDFs? [Y] yes, [C] cancel'
                        if ($null -eq $choice -or [string]::IsNullOrWhiteSpace($choice)) { $choice = 'C' }
                        $choice = $choice.Trim().ToUpperInvariant()
                        if ($choice -cnotin @('Y', 'C')) { Write-Host 'Choose exactly Y or C.' }
                    } while ($choice -cnotin @('Y', 'C'))
                    if ($choice -ceq 'Y') {
                        $script:confirmedPlan = $plan
                        $reply.action = 'execute'; $reply.plan_sha256 = $event.plan_sha256
                    }
                    $nextStage = 'terminal'
                }
                'no_plan' {
                    if ($nextStage -cne 'initial' -or $currentMode -ceq 'manual') { throw 'Unexpected no-plan event.' }
                    $noPlan = ConvertFrom-SplitResult -Stdout ($event.result | ConvertTo-Json -Depth 100 -Compress) -ExitCode 5 -Mode $currentMode
                    if (($event.fallback_modes | ConvertTo-Json -Compress) -cne ($noPlan.fallback_modes | ConvertTo-Json -Compress)) { throw 'Fallback choices disagree.' }
                    $decision = Get-SplitDecision -Result $noPlan
                    Write-Host (ConvertTo-WinBookSplitDisplayText $decision.message) -ForegroundColor Yellow
                    do {
                        $prompt = 'Choose [M] manual or [C] cancel'
                        if ($decision.fallback_modes -contains '1') { $prompt = 'Choose [1] Level 1, [M] manual or [C] cancel' }
                        $choice = Read-WinBookSplitSessionLine $session $prompt
                        if ($null -eq $choice) { $choice = 'C' }
                        $decision = Get-SplitDecision -Result $noPlan -Choice $choice
                        $script:consoleLogWriter.WriteLine("[DECISION] $($decision.decision); retry mode: $($decision.retry_mode)")
                        if ($decision.decision -in @('pending', 'invalid')) { Write-Host 'Choose one of the displayed options.' }
                    } while ($decision.decision -in @('pending', 'invalid'))
                    if ($decision.decision -ceq 'retry') {
                        $currentMode = $decision.retry_mode; $SuppliedStarts = ''
                        $reply.action = 'retry'; $reply.mode = $currentMode
                    }
                    else { $script:noPlanCancelled = $true; $nextStage = 'terminal' }
                }
            }
            $script:consoleLogWriter.WriteLine('[INTERACTION-REPLY] ' + ($reply | ConvertTo-Json -Compress))
            $session.SendReply(($reply | ConvertTo-Json -Compress))
        }
    }
    catch {
        $script:interactionFailure = $_.Exception
        $script:interactionStopRequested = -not $session.Completed
        $session.Cancel()
    }
    finally {
        $session.Dispose()
        $script:interactiveMode = $currentMode
    }
    return $session.Result
}

# --- Validation ---
Draw-Header
try {
    # Explicit choices form a validated processing set; no binding prompt may
    # supply a missing choice or silently discard a contradictory option.
    if ($args.Count -gt 0) { throw 'Unrecognized or extra arguments. Use Get-Help for the supported parameters.' }
    if ($Version) { throw 'Version cannot be combined with processing parameters.' }
    $hasMode = $PSBoundParameters.ContainsKey('Mode')
    $hasStarts = $PSBoundParameters.ContainsKey('StartPages')
    $hasLevel = $PSBoundParameters.ContainsKey('BookmarkLevel')
    if ($hasMode) {
        if ($Mode -notin @('Auto', 'Manual')) { throw 'Mode must be Auto or Manual.' }
        if ($Mode -ieq 'Auto') {
            if ($hasStarts) { throw 'Auto cannot be combined with StartPages.' }
            if (-not $hasLevel -or $BookmarkLevel -cnotin @('1', '2')) { throw 'Auto requires BookmarkLevel 1 or 2.' }
        }
        else {
            if ($hasLevel) { throw 'Manual cannot be combined with BookmarkLevel.' }
            if (-not $hasStarts -or $StartPages -notmatch '^\s*[0-9]+\s*(,\s*[0-9]+\s*)*$') {
                throw 'Manual requires comma-separated ASCII decimal physical PDF StartPages; every token must be valid.'
            }
        }
    }
    elseif ($hasStarts -or $hasLevel -or $NonInteractive -or $Preview) { throw 'Supply Mode and its BookmarkLevel or StartPages.' }
    if ($Preview -and $KeepConvertedPdf) { throw 'Preview cannot retain a converted PDF; omit KeepConvertedPdf.' }
    $InputFile = Resolve-WinBookSplitInputChoice -InputFile $InputFile -NonInteractive ([bool]$NonInteractive) -Preview ([bool]$Preview)
    $inputItem = Resolve-WinBookSplitInput -Path $InputFile
    $InputFile = $inputItem.FullName
    if ($KeepConvertedPdf -and [IO.Path]::GetExtension($InputFile) -ieq '.pdf') { throw 'KeepConvertedPdf applies only to EPUB or AZW3 conversion.' }
    $requestedBase = $documentsPath
    if (-not [string]::IsNullOrWhiteSpace($OutputDirectory)) { $requestedBase = $OutputDirectory }
    $outputDir = Resolve-WinBookSplitOutputBase -Path $requestedBase
}
catch {
    $validationCode, $validationStatus = 'invalid_arguments', 'failed'
    if ($_.Exception.Data.Contains('Code') -and $_.Exception.Data['Code'] -eq 'cancelled') {
        $validationCode, $validationStatus = 'cancelled', 'cancelled'
    }
    $outcome = New-WinBookSplitOutcome -Status $validationStatus -Code $validationCode -Message $_.Exception.Message
    if ($outcome.status -ceq 'cancelled') { Write-Host (ConvertTo-WinBookSplitDisplayText $outcome.message) -ForegroundColor Yellow }
    else { Write-Host ("[!] Error: " + (ConvertTo-WinBookSplitDisplayText $outcome.message)) -ForegroundColor Red }
    Write-Host ('[OUTCOME] ' + ($outcome | ConvertTo-Json -Depth 100 -Compress))
    if ($outcome.exit_code -ne 130) { Wait-WinBookSplitExit }
    exit $outcome.exit_code
}

$inputExt = [System.IO.Path]::GetExtension($InputFile).ToLower()

# Validate dependencies before reserving console records or starting conversion.
# The selected absolute interpreter is also used for every engine attempt.
$runtime = $null
$converter = $null
$converterExe = $null
try {
    $runtime = Resolve-WinBookSplitRuntime -ApplicationRoot $PSScriptRoot -PythonPath $PythonPath -DocumentPath $InputFile
    $pythonExe = $runtime.Path
    if ($inputExt -ne ".pdf") {
        $converter = Resolve-WinBookSplitConverter -ApplicationRoot $PSScriptRoot -CalibrePath $CalibrePath -DocumentPath $InputFile
        $converterExe = $converter.Path
    }
    elseif ($KeepConvertedPdf) { throw (New-WinBookSplitFailure 'invalid_arguments' 'KeepConvertedPdf applies only to EPUB or AZW3 conversion.') }
}
catch {
    $preflightFailure = $_.Exception
    Write-Host ('[!] Dependency preflight failed: ' + (ConvertTo-WinBookSplitDisplayText $preflightFailure.Message)) -ForegroundColor Red
    if ($preflightFailure.Data.Contains('Code')) {
        Write-Host ('[DEPENDENCY-ERROR] ' + (@{ Code = $preflightFailure.Data['Code']; Attempts = @($preflightFailure.Data['Attempts']); Probe = $preflightFailure.Data['Probe'] } | ConvertTo-Json -Depth 12 -Compress))
    }
    $failureCode = 'runtime_invalid'
    if ($preflightFailure.Data.Contains('Code')) { $failureCode = [string]$preflightFailure.Data['Code'] }
    $failureStatus = 'failed'
    if ($failureCode -eq 'processor_cancelled') { $failureStatus = 'cancelled' }
    elseif ($failureCode -eq 'processor_timeout') { $failureStatus = 'timeout' }
    $outcome = New-WinBookSplitOutcome -Status $failureStatus -Code $failureCode -Message $preflightFailure.Message
    Write-Host ('[OUTCOME] ' + ($outcome | ConvertTo-Json -Depth 100 -Compress))
    if ($outcome.exit_code -ne 130) { Wait-WinBookSplitExit }
    exit $outcome.exit_code
}

# Book metadata always describes the original source.
$fileName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
$fileSize = "{0:N2} MB" -f ((Get-Item -LiteralPath $InputFile -ErrorAction Stop).Length / 1MB)

# Console diagnostics have their own exclusive run folder; chapters are
# published by the engine only after the complete stage has been validated.
$markerStream = $null
$logStream = $null
$script:consoleLogWriter = $null
$script:closedConsoleLogWriter = $null
$script:consoleRecordLeases = @()
$script:confirmedPlan = $null
$script:runStartedUtc = [DateTime]::UtcNow.ToString('o')
$consoleDir = $null
try {
    $sourceItem = Get-Item -LiteralPath $InputFile -Force -ErrorAction Stop
    $script:sourceObservation = [pscustomobject]@{ path = $InputFile; size_bytes = $sourceItem.Length;
        last_write_utc = $sourceItem.LastWriteTimeUtc.ToString('o'); binding = 'metadata_only'; sha256 = $null }
    if ($Preview) {
        # Retain the same bounded transport/diagnostic path without disk logs.
        $script:consoleLogWriter = New-WinBookSplitLogWriter -Stream ([IO.MemoryStream]::new())
        $script:closedConsoleLogWriter = $script:consoleLogWriter
    }
    else {
    $script:consoleRecordLeases = New-WinBookSplitRecordLeases -Directory $outputDir -IncludeMarker $false
    foreach ($lease in $script:consoleRecordLeases) { $lease.AssertUnchanged() }
    $consoleReservation = New-WinBookSplitConsoleDirectory -Base $outputDir
    $consoleRunId = $consoleReservation.run_id
    $consoleDir = $consoleReservation.path
    $script:consoleRecordLeases.Add([WinBookSplitLogging.PathLease]::new($consoleDir, $true))
    foreach ($lease in $script:consoleRecordLeases) { $lease.AssertUnchanged() }
    $consoleMarker = [IO.Path]::Combine($consoleDir, '.WinBookSplit-console-owner.json')
    $markerStream = [IO.File]::Open($consoleMarker, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
    $createdMarkerLease = $null; $sealedMarkerLease = $null
    try {
        $createdMarkerLease = [WinBookSplitLogging.PathLease]::new($consoleMarker, $false)
        $createdMarkerLease.AssertSameHandle($markerStream.SafeFileHandle)
        $markerBytes = [Text.Encoding]::UTF8.GetBytes((@{ run_id = $consoleRunId; kind = 'console' } | ConvertTo-Json -Compress))
        $markerStream.Write($markerBytes, 0, $markerBytes.Length)
        $markerStream.Flush($true); $markerStream.Dispose(); $markerStream = $null
        $sealedMarkerLease = [WinBookSplitLogging.PathLease]::new($consoleMarker, $false, $false, $true)
        $createdMarkerLease.AssertSameLease($sealedMarkerLease)
        $script:consoleRecordLeases.Add($sealedMarkerLease); $sealedMarkerLease = $null
    }
    finally {
        if ($null -ne $markerStream) { $markerStream.Dispose(); $markerStream = $null }
        if ($null -ne $createdMarkerLease) { $createdMarkerLease.Dispose() }
        if ($null -ne $sealedMarkerLease) { $sealedMarkerLease.Dispose() }
    }
    $logFile = [IO.Path]::Combine($consoleDir, 'console.log')
    $logStream = [IO.File]::Open($logFile, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
    $script:consoleLogWriter = New-WinBookSplitLogWriter -Stream $logStream -Path $logFile
    $script:closedConsoleLogWriter = $script:consoleLogWriter
    }
}
catch {
    $setupMessage = $_.Exception.Message
    foreach ($stream in @($script:consoleLogWriter, $logStream, $markerStream)) {
        if ($null -ne $stream) {
            try { $stream.Dispose() }
            catch { $setupMessage += '; Cannot close a console stream: ' + $_.Exception.Message }
        }
    }
    foreach ($lease in $script:consoleRecordLeases) { $lease.Dispose() }
    Write-Host ('[!] Error: Cannot prepare console records: ' + (ConvertTo-WinBookSplitDisplayText $setupMessage)) -ForegroundColor Red
    if ($null -ne $consoleDir) { Write-Host ('Reserved console directory retained: ' + (ConvertTo-WinBookSplitDisplayText $consoleDir)) -ForegroundColor Gray }
    $outcome = New-WinBookSplitOutcome -Status 'failed' -Code 'console_setup_failed' -Message $setupMessage
    Write-Host ('[OUTCOME] ' + ($outcome | ConvertTo-Json -Depth 100 -Compress))
    Wait-WinBookSplitExit
    exit $outcome.exit_code
}

# --- The Python Engine ---
$enginePath = Join-Path $PSScriptRoot 'engine\winbooksplit_engine.py'

# --- Execution Function ---
function Run-PythonSplitter ($mode, $manualData, [bool]$PlanOnly = $false, [bool]$Interactive = $false) {
    $script:lastEngineResult = $null
    $script:consoleRecordError = $null
    # -I ignores PYTHON* environment settings. -X utf8 explicitly applies to
    # isolated Python as well as the UTF-8 environment inherited by Calibre.
    $engineArguments = @('-I', '-B', '-X', 'utf8', $enginePath, $InputFile, $outputDir, $mode, $manualData)
    if ($PlanOnly) { $engineArguments += '--preview' }
    if ($Interactive) { $sessionId = [Guid]::NewGuid().ToString('N'); $engineArguments += @('--interactive', $sessionId) }
    if ($InputFile -match '\.(epub|azw3)$') {
        $engineArguments += @('--calibre-path', $converterExe, '--conversion-timeout', [string]$ConversionTimeout)
        if ($KeepConvertedPdf) { $engineArguments += '--keep-converted-pdf' }
    }
    $script:consoleLogWriter.WriteLine('[ENGINE] ' + (@{ path = $pythonExe; arguments = @($engineArguments) } | ConvertTo-Json -Compress))
    if ($Interactive) {
        $script:lastProcessResult = Invoke-WinBookSplitInteractiveEngine $engineArguments $sessionId $mode $manualData
        $mode = $script:interactiveMode
    }
    else {
        $script:lastProcessResult = Invoke-WinBookSplitProcess -Path $pythonExe -Arguments $engineArguments `
            -WorkingDirectory ([IO.Path]::GetDirectoryName($enginePath)) -TimeoutSeconds $ProcessTimeout
    }
    $transport = $script:lastProcessResult
    # Classify this attempt before logging. A secondary log error cannot hide
    # its cancellation/failure or the location of completed engine output.
    $attemptFailure = $null
    $shutdownMessage = 'Its owned process tree was proved stopped; any incomplete output is retained.'
    if (-not $transport.ParentStopped -or -not $transport.DescendantsStopped -or $transport.StopError) {
        $shutdownMessage = 'Shutdown could not be proved complete; any incomplete output is retained. ' + $transport.StopError
    }
    if ($Interactive -and $null -ne $script:interactionFailure -and $script:interactionStopRequested -and
        -not $script:abandonedConsoleInput -and
        -not $transport.TimedOut) {
        $attemptFailure = New-WinBookSplitFailure 'processor_protocol_failed' ($script:interactionFailure.Message + ' ' + $shutdownMessage)
    }
    elseif ($transport.Cancelled) { $attemptFailure = New-WinBookSplitFailure 'processor_cancelled' ('Engine processing was cancelled. ' + $shutdownMessage) }
    elseif ($transport.TimedOut) { $attemptFailure = New-WinBookSplitFailure 'processor_timeout' ('Engine processing exceeded ProcessTimeout (' + $ProcessTimeout + ' seconds), including pipe EOF. ' + $shutdownMessage) }
    else {
        $errorFields = @('StartError', 'StopError', 'StreamError', 'ResultError')
        if ($Interactive) { $errorFields += @('InteractionError', 'InputError') }
        foreach ($name in $errorFields) {
            if ($transport.$name) {
                $failureCode = 'processor_protocol_failed'
                if ($name -eq 'StartError') { $failureCode = 'processor_start_failed' }
                $attemptFailure = New-WinBookSplitFailure $failureCode ($name + ': ' + $transport.$name)
                break
            }
        }
        if ($null -eq $attemptFailure -and (($Interactive -and -not $transport.InputWriterStopped) -or
            -not $transport.ParentStopped -or -not $transport.DescendantsStopped -or
            -not $transport.StreamsComplete -or $null -eq $transport.ExitCode)) {
            $attemptFailure = New-WinBookSplitFailure 'processor_protocol_failed' 'The engine process tree and streams did not complete safely.'
        }
    }
    if (-not $transport.Cancelled -and -not $transport.TimedOut -and
        $transport.ParentStopped -and $transport.DescendantsStopped -and $transport.StreamsComplete -and
        $null -ne $transport.ExitCode -and -not $transport.StartError -and
        -not $transport.StreamError -and -not $transport.ResultError) {
        # A handle/handler close error can follow a fully validated publication.
        # Keep its location while preserving the primary transport failure.
        try {
            $script:lastEngineResult = ConvertFrom-SplitResult -Stdout ($transport.ResultRecords -join "`n") -ExitCode $transport.ExitCode -Mode $mode -Preview $PlanOnly
            if ($script:lastEngineResult.PSObject.Properties.Name -contains 'plan' -and $null -ne $script:lastEngineResult.plan) {
                Assert-SplitPreviewPlan -Plan $script:lastEngineResult.plan -Mode $mode -Interactive $true -KeepConvertedPdf ([bool]$KeepConvertedPdf)
            }
            if ($Interactive) { Assert-WinBookSplitConfirmedExecution $script:confirmedPlan $script:lastEngineResult }
        }
        catch {
            if ($null -eq $attemptFailure) { $attemptFailure = New-WinBookSplitFailure 'processor_protocol_failed' $_.Exception.Message }
        }
    }
    try {
    $summary = [ordered]@{}
    foreach ($name in @('Pid', 'ExitCode', 'TimedOut', 'Cancelled', 'ParentStopped', 'DescendantsStopped',
        'StreamsComplete', 'JobAssigned', 'StartError', 'StopError', 'StreamError', 'ResultError',
        'StdoutTotalBytes', 'StderrTotalBytes', 'StdoutTruncated', 'StderrTruncated', 'ElapsedSeconds')) {
        $summary[$name] = $transport.$name
    }
    if ($Interactive) {
        foreach ($name in @('InteractionError', 'InputError', 'InputWriterStopped', 'InteractionCount',
                             'InteractionTotalBytes', 'QueuedReplyCount', 'ReplyCount', 'ReplyTotalBytes')) {
            $summary[$name] = $transport.$name
        }
    }
    $summary['ResultRecordCount'] = @($transport.ResultRecords).Count
    $script:consoleLogWriter.WriteLine('[PROCESS] ' + ($summary | ConvertTo-Json -Compress))
    foreach ($name in @('Stdout', 'Stderr')) {
        $text = $transport.$name
        $truncated = $transport.($name + 'Truncated')
        $script:consoleLogWriter.WriteLine('[' + $name.ToUpperInvariant() + ']')
        if ($truncated) {
            $notice = '[TRUNCATED] Retained the final 65536 bytes of ' + $name.ToLowerInvariant() + '; see [PROCESS] totals.'
            Write-Host $notice -ForegroundColor Yellow
            $script:consoleLogWriter.WriteLine($notice)
        }
        # Raw Write preserves blank lines and a final unterminated line.
        $script:consoleLogWriter.Write($text)
        $script:consoleLogWriter.WriteLine()
    }
    # A result survives even when subsequent human output evicts it from the
    # bounded stdout tail. Its separate cap fails explicitly on oversized JSON.
    foreach ($record in $transport.ResultRecords) {
        if (-not $transport.Stdout.Contains($record)) { $script:consoleLogWriter.WriteLine($record) }
    }
    }
    catch {
        $script:consoleRecordError = 'Cannot record engine diagnostics: ' + $_.Exception.Message
        if ($null -eq $attemptFailure -and ($null -eq $script:lastEngineResult -or $script:lastEngineResult.exit_code -eq 0)) {
            $attemptFailure = New-WinBookSplitFailure 'console_finalize_failed' $script:consoleRecordError
        }
    }
    # A disk-record failure must not hide either drained process stream.
    foreach ($name in @('Stdout', 'Stderr')) {
        $display = $transport.$name
        if ($name -eq 'Stdout') { $display = $transport.HumanStdout }
        if ($display.Length -gt 0) {
            $color = 'Green'; if ($name -eq 'Stderr') { $color = 'Red' }
            $displayLines = foreach ($line in [regex]::Split($display, '\r\n|\r|\n')) {
                $safeLine = ConvertTo-WinBookSplitDisplayText $line
                if ($safeLine -cmatch '^(\[OUTCOME\] |Log: )') { $safeLine = '[PROCESS] ' + $safeLine }
                $safeLine
            }
            Write-Host ($displayLines -join "`n") -ForegroundColor $color
        }
    }
    if ($null -ne $attemptFailure) { throw $attemptFailure }
    return $script:lastEngineResult
}

# --- Main Logic Flow ---
$result = $null
$outcome = $null
$script:lastEngineResult = $null
$script:consoleRecordError = $null
$requestedMode = $Mode
$mode = ''
$script:abandonedConsoleInput = $false
$script:interactiveMode = ''
try {

# 1. Initial TUI, shows original book metadata
$script:consoleLogWriter.WriteLine('[DEPENDENCY] ' + (@{ Runtime = $runtime; Converter = $converter } | ConvertTo-Json -Depth 12 -Compress))
$runtimeLine = 'Python: ' + $runtime.Path + ' (' + $runtime.Version + '; ' + $runtime.Source + ')'
$pypdfLine = 'pypdf: ' + $runtime.PypdfVersion + ' (' + $runtime.PypdfPath + ')'
foreach ($line in @($runtimeLine, $pypdfLine)) {
    Write-Host (ConvertTo-WinBookSplitDisplayText $line) -ForegroundColor Gray
    $script:consoleLogWriter.WriteLine($line)
}
if ($null -ne $converter) {
    $converterLine = 'Calibre: ' + $converter.Path + ' (' + $converter.Version + '; ' + $converter.Source + ')'
    Write-Host (ConvertTo-WinBookSplitDisplayText $converterLine) -ForegroundColor Gray
    $script:consoleLogWriter.WriteLine($converterLine)
}
Write-Host ('Target Book: ' + (ConvertTo-WinBookSplitDisplayText $fileName)) -ForegroundColor Green
Write-Host "Size:        $fileSize" -ForegroundColor Gray
if ($inputExt -ne ".pdf") {
     Write-Host "Format:      $inputExt (Calibre PDF conversion before splitting)" -ForegroundColor DarkGray
     Write-Host 'Conversion:  tablet profile; validated in a temporary owned workspace' -ForegroundColor DarkGray
     if ($KeepConvertedPdf) { Write-Host 'Full PDF:    retained with successful chapter output' -ForegroundColor DarkGray }
}
Write-Host ""
if ($hasMode) {
    $manualInput = ''
    if ($requestedMode -ieq 'Manual') { $mode = 'manual'; $manualInput = $StartPages }
    else { $mode = $BookmarkLevel }
}
else {
Write-Host "Select Splitting Method:" -ForegroundColor White
Write-Host " [1] Level 1 Bookmarks (Auto)" -ForegroundColor Cyan
Write-Host " [2] Level 2 Bookmarks (Auto)" -ForegroundColor Cyan
Write-Host " [M] Manual Entry (Page Numbers)" -ForegroundColor Magenta
Write-Host " [C] Cancel" -ForegroundColor Gray
Write-Host ""

$selection = Read-WinBookSplitInitialDecision
if ($selection.decision -eq 'cancel') { throw (New-WinBookSplitFailure 'cancelled' 'Method selection was cancelled before processing.') }
$mode = $selection.mode

# Manual starts are requested only after the supervised engine captures the PDF.
$manualInput = ""
}

# 3. Execution
Draw-Header
if ($Preview) { Write-Host 'Preparing plan only...' -ForegroundColor Yellow }
else {
    Write-Host "Running Processor..." -ForegroundColor Yellow
    Write-Host ('Log: ' + (ConvertTo-WinBookSplitDisplayText $logFile)) -ForegroundColor Gray
}
Write-Host ""

    $interactive = -not $NonInteractive -and -not $Preview
    $result = Run-PythonSplitter -mode $mode -manualData $manualInput -PlanOnly ([bool]$Preview) -Interactive $interactive
    if ($interactive) { $mode = $script:interactiveMode }
    if ($interactive -and $script:noPlanCancelled) {
        $outcome = New-WinBookSplitOutcome -Status 'cancelled' -Code 'cancelled' -Message 'Splitting was cancelled before a fallback attempt.' -Mode $mode -EngineResult $result
    }
    if ($null -eq $outcome) {
        $outcome = New-WinBookSplitOutcome -Status $result.status -Code $result.code -Message (Get-SplitDecision -Result $result).message -Mode $mode -EngineResult $result
    }
}
catch {
    $failure = $_.Exception
    if (-not [string]::IsNullOrEmpty($script:interactiveMode)) { $mode = $script:interactiveMode }
    $failureCode = 'processor_protocol_failed'
    if ($failure.Data.Contains('Code')) { $failureCode = [string]$failure.Data['Code'] }
    $failureStatus = 'failed'
    if ($failureCode -in @('cancelled', 'processor_cancelled')) { $failureStatus = 'cancelled' }
    elseif ($failureCode -eq 'processor_timeout') { $failureStatus = 'timeout' }
    elseif ($null -ne $script:lastEngineResult -and $script:lastEngineResult.status -ceq 'success') {
        $failureStatus = 'incomplete'
        if (-not $failure.Data.Contains('Code')) { $failureCode = 'console_finalize_failed' }
    }
    $outcome = New-WinBookSplitOutcome -Status $failureStatus -Code $failureCode -Message $failure.Message -Mode $mode -EngineResult $script:lastEngineResult
    if ($outcome.status -cnotin @('cancelled', 'timeout')) {
        Write-Host ('[!] Processor failure: ' + (ConvertTo-WinBookSplitDisplayText $failure.Message)) -ForegroundColor Red
    }
}
if ($script:consoleRecordError -and -not $outcome.message.Contains($script:consoleRecordError)) { $outcome.message += '; ' + $script:consoleRecordError }
try { Complete-WinBookSplitConsoleLog -Outcome $outcome }
catch {
    $finalizeMessage = 'Cannot finalize local run diagnostics: ' + $_.Exception.Message
    $finalizeStatus = 'failed'
    if ($outcome.status -ceq 'success') { $finalizeStatus = 'incomplete' }
    if ($outcome.exit_code -ne 0) {
        # A secondary log failure must not replace a primary failure or cancellation.
        $outcome.message += '; ' + $finalizeMessage
    }
    else {
        $outcome = New-WinBookSplitOutcome -Status $finalizeStatus -Code 'console_finalize_failed' -Message ($outcome.message + '; ' + $finalizeMessage) -Mode $mode -EngineResult $outcome.engine_result
    }
    if ($null -ne $script:consoleLogWriter) {
        try { $script:consoleLogWriter.Dispose() }
        catch { $outcome.message += '; Cannot close the console log: ' + $_.Exception.Message }
        $script:consoleLogWriter = $null
    }
    if ($null -ne $script:closedConsoleLogWriter) {
        try { $script:closedConsoleLogWriter.AppendCorrection('[OPERATION-OUTCOME] ' + ($outcome | ConvertTo-Json -Depth 100 -Compress)) }
        catch { $outcome.message += '; Cannot record the final diagnostic failure: ' + $_.Exception.Message }
    }
    if (-not $Preview -and -not (Test-Path -LiteralPath ([IO.Path]::Combine($consoleDir, 'WinBookSplit_Run.pending.json'))) -and
        -not (Test-Path -LiteralPath ([IO.Path]::Combine($consoleDir, 'WinBookSplit_Run.json')))) {
        try { Write-WinBookSplitRunManifest -Directory $consoleDir -Manifest (New-WinBookSplitRunRecord -Outcome $outcome) -Leases $script:consoleRecordLeases }
        catch {
            $outcome.message += '; Cannot preserve a finalized failure manifest: ' + $_.Exception.Message
            # The fallback failure changes the authoritative outcome after the
            # first correction. Record that final outcome when the log is usable.
            if ($null -ne $script:closedConsoleLogWriter) {
                try { $script:closedConsoleLogWriter.AppendCorrection('[OPERATION-OUTCOME] ' + ($outcome | ConvertTo-Json -Depth 100 -Compress)) }
                catch { $outcome.message += '; Cannot record the final manifest failure: ' + $_.Exception.Message }
            }
        }
    }
}
if ($null -ne $script:closedConsoleLogWriter) { $script:closedConsoleLogWriter.CloseLease() }
foreach ($lease in $script:consoleRecordLeases) { $lease.Dispose() }
# Cancellation can finish while a no-newline interactive prompt is pending.
Write-Host ''
Write-Host ('[OUTCOME] ' + ($outcome | ConvertTo-Json -Depth 100 -Compress))
if ($outcome.status -ceq 'preview') {
    Show-WinBookSplitPlan -Plan $result.plan
}
elseif ($outcome.status -ceq 'success') {
    foreach ($warning in $result.warnings) {
        if ($warning.code -ceq 'output_handle_close_failed') { Write-Host ('[WARNING] ' + (ConvertTo-WinBookSplitDisplayText $warning.message)) -ForegroundColor Yellow }
    }
    Write-Host ('Output: ' + (ConvertTo-WinBookSplitDisplayText $outcome.final_directory)) -ForegroundColor Cyan
    Write-Host ('Chapters written: {0}; physical pages: {1}; every page exactly once.' -f
        $outcome.written_count, $result.execution.coverage.covered_pages) -ForegroundColor Cyan
    Write-Host 'Done.' -ForegroundColor Cyan
    if (-not $NonInteractive -and -not $Preview) {
        do {
            $openFolder = Read-WinBookSplitSessionLine $null 'Open output folder? [O] open, [N] finish'
            if ($null -eq $openFolder -or [string]::IsNullOrWhiteSpace($openFolder)) { $openFolder = 'N' }
            $openFolder = $openFolder.Trim().ToUpperInvariant()
            if ($openFolder -cnotin @('O', 'N')) { Write-Host 'Choose exactly O or N.' }
        } while ($openFolder -cnotin @('O', 'N'))
        if ($openFolder -ceq 'O') {
            try { Invoke-Item -LiteralPath $outcome.final_directory -ErrorAction Stop }
            catch { Write-Host ('Cannot open the output folder: ' + (ConvertTo-WinBookSplitDisplayText $_.Exception.Message)) -ForegroundColor Yellow }
        }
    }
}
else {
    $outcomeColor = 'Red'
    if ($outcome.status -cin @('cancelled', 'timeout')) { $outcomeColor = 'Yellow' }
    Write-Host (ConvertTo-WinBookSplitDisplayText $outcome.message) -ForegroundColor $outcomeColor
    if ($null -ne $outcome.final_directory) { Write-Host ('Completed engine output retained: ' + (ConvertTo-WinBookSplitDisplayText $outcome.final_directory)) -ForegroundColor Yellow }
}
$exitCode = $outcome.exit_code
Write-Host ""
if ($exitCode -ne 130 -and -not $script:abandonedConsoleInput) {
    Wait-WinBookSplitExit
}
exit $exitCode
