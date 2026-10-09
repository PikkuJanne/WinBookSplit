<#
WinBookSplit.ps1
Automated PDF/AZW3/EPUB Chapter Slicer & Organizer

Author: Janne Vuorela
Target OS: Windows 11
PowerShell: Windows PowerShell 5.1 or PowerShell 7+
Dependencies: Regular Windows x64 CPython 3.14.8, pypdf 6.19.0,
              Calibre 9.15.0 (for AZW3/EPUB), .bat wrapper

SYNOPSIS
    A "smart" decomposition tool for technical manuals and textbooks.
    Automates the tedious process of splitting large PDF/AZW3/EPUB files into 
    individual chapters based on internal metadata or manual page ranges.

WHAT THIS IS (AND ISN'T)
    - A logic-driven parser for PDF structures.
    - Designed for CS students/IT Support who need to modularize heavy documentation.
    - Hybrid tool, attempts Auto-Discovery first, falls back to Manual Slicing.
    - Not an OCR engine, it cannot read text on a flat image to find chapters.
    - Not a PDF editor, it creates new files and does not modify the source.

FEATURES
    - Format Conversion:
        Automatically detects AZW3/EPUB files and utilizes Calibre's engine 
        to generate a high-quality PDF source before splitting.
    - Text User Interface (TUI):
        Dynamic console interface providing real-time feedback and file stats.
        Includes a Bookmark Depth selector for Level 1 or Level 2.
    - Intelligent Metadata Extraction:
        Recursively crawls the PDF outline tree to map chapter starts and ends.
    - Smart Manual Fallback:
        Triggers an interactive manual mode if no metadata bookmarks are found.
    - Automated Sanitization:
        Cleans illegal Windows characters from titles to ensure valid filenames.
    - Systematic Naming:
        Prefixes files with at least two digits, widened to the section count.
    - Verbose Logging:
        Generates a detailed execution log tracking every match and action.

MY INTENDED USAGE
    - I keep WinBookSplit in my Tools directory with a shortcut to the .bat on my desktop.
    - When I download a massive manual, AZW3 book, or a textbook:
        1. I drag the PDF or AZW3 onto the .bat launcher.
        2. If it's an AZW3, I let the script convert it to PDF first.
        3. I attempt an "Auto" split for chapters.
        4. If the book is "flat," I switch to manual mode and input page numbers.
        5. I find a clean, numbered folder in my Documents, ready for my reader.

SETUP
    1) Follow docs/codex-v1.0.0/SUPPORT_AND_SETUP.md for explicit isolated setup.
       Install hashed requirements.txt with the exact interpreter's -m pip.
    2) Install Calibre 9.15.0 separately for AZW3/EPUB conversion support.
    3) Create a folder (e.g., C:\Tools\WinBookSplit\).
    4) Place the .ps1, .bat, and engine directory inside.
    5) (Optional) Create a desktop shortcut to WinBookSplit.bat.

USAGE
    A) Drag-and-Drop (Recommended)
        - Drag any PDF, AZW3, or EPUB file onto WinBookSplit.bat.
        - Follow the TUI prompts for conversion and splitting.

    B) Direct PowerShell
        - Run: .\WinBookSplit.ps1 -InputFile "C:\Path\To\Book.azw3"
        - Optional: -PythonPath "C:\Tools\Python\python.exe"
                    -CalibrePath "C:\Calibre\ebook-convert.exe" -KeepConvertedPdf
        - ConversionTimeout limits conversion in seconds (default 1800).
        - ProcessTimeout bounds the engine and its streams (default at least 3600).

NOTES
    - The script uses the shipped engine\winbooksplit_engine.py beside the
      PowerShell entry point, independently of the current working directory.
    - Output uses a new child of Documents or the explicit -OutputDirectory base.

LIMITATIONS
    - Requires a local Python installation with the 'pypdf' library.
    - Requires a validated trusted Calibre executable for AZW3/EPUB conversion.
    - Manual mode assumes every provided page number is the start of a new chapter.

TROUBLESHOOTING
    - "Calibre not found":
        Ensure Calibre is installed to process non-PDF ebook formats.
    - "[!!!] ERROR: No bookmarks found":
        The PDF lacks an internal Outline. Use Manual Mode [M] instead.
    - "Python not found":
        Select regular x64 Python 3.14.8 with -PythonPath or prepare the app .venv.

LICENSE / WARRANTY
    - Personal IT automation tool, provided as-is.
#>

param (
    [string]$InputFile,
    [string]$OutputDirectory,
    [string]$CalibrePath,
    [switch]$KeepConvertedPdf,
    [ValidateRange(1, 86400)][int]$ConversionTimeout = 1800,
    [string]$PythonPath,
    [ValidateRange(1, 172800)][int]$ProcessTimeout = [Math]::Max(3600, $ConversionTimeout + 1800)
)

# --- Configuration ---
$AppName = "WinBookSplit"
$ver = "2.1 (AZW3 Support)"
$documentsPath = [Environment]::GetFolderPath("MyDocuments")

# --- UI Functions ---
function Draw-Header {
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "      $AppName v$ver" -ForegroundColor Yellow
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
}

. (Join-Path $PSScriptRoot 'engine\WinBookSplit.Paths.ps1')
. (Join-Path $PSScriptRoot 'engine\WinBookSplit.Diagnostics.ps1')
. (Join-Path $PSScriptRoot 'engine\WinBookSplit.Runtime.ps1')
. (Join-Path $PSScriptRoot 'engine\WinBookSplit.Process.ps1')

# --- Validation ---
Draw-Header
try {
    $inputItem = Resolve-WinBookSplitInput -Path $InputFile
    $InputFile = $inputItem.FullName
    $requestedBase = $documentsPath
    if (-not [string]::IsNullOrWhiteSpace($OutputDirectory)) { $requestedBase = $OutputDirectory }
    $outputDir = Resolve-WinBookSplitOutputBase -Path $requestedBase
}
catch {
    Write-Host ("[!] Error: " + $_.Exception.Message) -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
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
    elseif ($KeepConvertedPdf) { throw 'KeepConvertedPdf applies only to EPUB or AZW3 conversion.' }
}
catch {
    $preflightFailure = $_.Exception
    Write-Host ('[!] Dependency preflight failed: ' + $preflightFailure.Message) -ForegroundColor Red
    if ($preflightFailure.Data.Contains('Code')) {
        Write-Host ('[DEPENDENCY-ERROR] ' + (@{ Code = $preflightFailure.Data['Code']; Attempts = @($preflightFailure.Data['Attempts']) } | ConvertTo-Json -Depth 12 -Compress))
    }
    Read-Host 'Press Enter to exit'
    exit 1
}

# Book metadata always describes the original source.
$fileName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
$fileSize = "{0:N2} MB" -f ((Get-Item -LiteralPath $InputFile -ErrorAction Stop).Length / 1MB)

# Console diagnostics have their own exclusive run folder; chapters are
# published by the engine only after the complete stage has been validated.
$markerStream = $null
$logStream = $null
$script:consoleLogWriter = $null
$consoleDir = $null
try {
    $consoleReservation = New-WinBookSplitConsoleDirectory -Base $outputDir
    $consoleRunId = $consoleReservation.run_id
    $consoleDir = $consoleReservation.path
    $consoleMarker = [IO.Path]::Combine($consoleDir, '.WinBookSplit-console-owner.json')
    $markerStream = [IO.File]::Open($consoleMarker, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
    try {
        $markerBytes = [Text.Encoding]::UTF8.GetBytes((@{ run_id = $consoleRunId; kind = 'console' } | ConvertTo-Json -Compress))
        $markerStream.Write($markerBytes, 0, $markerBytes.Length)
    }
    finally { if ($null -ne $markerStream) { $markerStream.Dispose(); $markerStream = $null } }
    $logFile = [IO.Path]::Combine($consoleDir, 'console.log')
    $logStream = [IO.File]::Open($logFile, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::Read)
    $script:consoleLogWriter = [IO.StreamWriter]::new($logStream, [Text.UTF8Encoding]::new($false))
    $script:consoleLogWriter.AutoFlush = $true
}
catch {
    $setupMessage = $_.Exception.Message
    foreach ($stream in @($script:consoleLogWriter, $logStream, $markerStream)) {
        if ($null -ne $stream) {
            try { $stream.Dispose() }
            catch { $setupMessage += '; Cannot close a console stream: ' + $_.Exception.Message }
        }
    }
    Write-Host ('[!] Error: Cannot prepare console records: ' + $setupMessage) -ForegroundColor Red
    if ($null -ne $consoleDir) { Write-Host ('Reserved console directory retained: ' + $consoleDir) -ForegroundColor Gray }
    Read-Host 'Press Enter to exit'
    exit 1
}

# --- The Python Engine ---
$enginePath = Join-Path $PSScriptRoot 'engine\winbooksplit_engine.py'

# --- Execution Function ---
function Run-PythonSplitter ($mode, $manualData) {
    # -I ignores PYTHON* environment settings. -X utf8 explicitly applies to
    # isolated Python as well as the UTF-8 environment inherited by Calibre.
    $engineArguments = @('-I', '-B', '-X', 'utf8', $enginePath, $InputFile, $outputDir, $mode, $manualData)
    if ($InputFile -match '\.(epub|azw3)$') {
        $engineArguments += @('--calibre-path', $converterExe, '--conversion-timeout', [string]$ConversionTimeout)
        if ($KeepConvertedPdf) { $engineArguments += '--keep-converted-pdf' }
    }
    $script:consoleLogWriter.WriteLine('[ENGINE] ' + (@{ path = $pythonExe; arguments = @($engineArguments) } | ConvertTo-Json -Compress))
    $script:lastProcessResult = Invoke-WinBookSplitProcess -Path $pythonExe -Arguments $engineArguments `
        -WorkingDirectory ([IO.Path]::GetDirectoryName($enginePath)) -TimeoutSeconds $ProcessTimeout
    $transport = $script:lastProcessResult
    $summary = [ordered]@{}
    foreach ($name in @('Pid', 'ExitCode', 'TimedOut', 'Cancelled', 'ParentStopped', 'DescendantsStopped',
        'StreamsComplete', 'JobAssigned', 'StartError', 'StopError', 'StreamError', 'ResultError',
        'StdoutTotalBytes', 'StderrTotalBytes', 'StdoutTruncated', 'StderrTruncated', 'ElapsedSeconds')) {
        $summary[$name] = $transport.$name
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
        $display = $text
        if ($name -eq 'Stdout') {
            $display = $transport.HumanStdout
        }
        if ($display.Length -gt 0) {
            $color = 'Green'; if ($name -eq 'Stderr') { $color = 'Red' }
            Write-Host $display -ForegroundColor $color
        }
    }
    # A result survives even when subsequent human output evicts it from the
    # bounded stdout tail. Its separate cap fails explicitly on oversized JSON.
    foreach ($record in $transport.ResultRecords) {
        if (-not $transport.Stdout.Contains($record)) { $script:consoleLogWriter.WriteLine($record) }
    }
    $shutdownMessage = 'Its owned process tree was proved stopped; any incomplete output is retained.'
    if (-not $transport.ParentStopped -or -not $transport.DescendantsStopped -or $transport.StopError) {
        $shutdownMessage = 'Shutdown could not be proved complete; any incomplete output is retained. ' + $transport.StopError
    }
    if ($transport.Cancelled) { throw ('Engine processing was cancelled. ' + $shutdownMessage) }
    if ($transport.TimedOut) { throw ('Engine processing exceeded ProcessTimeout (' + $ProcessTimeout + ' seconds), including pipe EOF. ' + $shutdownMessage) }
    foreach ($name in @('StartError', 'StopError', 'StreamError', 'ResultError')) {
        if ($transport.$name) { throw ($name + ': ' + $transport.$name) }
    }
    if (-not $transport.ParentStopped -or -not $transport.DescendantsStopped -or
        -not $transport.StreamsComplete -or $null -eq $transport.ExitCode) { throw 'The engine process tree and streams did not complete safely.' }
    return ConvertFrom-SplitResult -Stdout ($transport.ResultRecords -join "`n") -ExitCode $transport.ExitCode -Mode $mode
}

# --- Main Logic Flow ---

# 1. Initial TUI, shows original book metadata
$script:consoleLogWriter.WriteLine('[DEPENDENCY] ' + (@{ Runtime = $runtime; Converter = $converter } | ConvertTo-Json -Depth 12 -Compress))
$runtimeLine = 'Python: ' + $runtime.Path + ' (' + $runtime.Version + '; ' + $runtime.Source + ')'
$pypdfLine = 'pypdf: ' + $runtime.PypdfVersion + ' (' + $runtime.PypdfPath + ')'
foreach ($line in @($runtimeLine, $pypdfLine)) {
    Write-Host $line -ForegroundColor Gray
    $script:consoleLogWriter.WriteLine($line)
}
if ($null -ne $converter) {
    $converterLine = 'Calibre: ' + $converter.Path + ' (' + $converter.Version + '; ' + $converter.Source + ')'
    Write-Host $converterLine -ForegroundColor Gray
    $script:consoleLogWriter.WriteLine($converterLine)
}
Write-Host "Target Book: $fileName" -ForegroundColor Green
Write-Host "Size:        $fileSize" -ForegroundColor Gray
if ($inputExt -ne ".pdf") {
     Write-Host "Format:      $inputExt (Calibre PDF conversion before splitting)" -ForegroundColor DarkGray
     Write-Host 'Conversion:  tablet profile; validated in a temporary owned workspace' -ForegroundColor DarkGray
     if ($KeepConvertedPdf) { Write-Host 'Full PDF:    retained with successful chapter output' -ForegroundColor DarkGray }
}
Write-Host ""
Write-Host "Select Splitting Method:" -ForegroundColor White
Write-Host " [1] Level 1 Bookmarks (Auto)" -ForegroundColor Cyan
Write-Host " [2] Level 2 Bookmarks (Auto)" -ForegroundColor Cyan
Write-Host " [M] Manual Entry (Page Numbers)" -ForegroundColor Magenta
Write-Host ""

$selection = Read-Host "Enter selection"

if ($selection -match "m|M") { $mode = "manual" } elseif ($selection -eq "1" -or $selection -eq "2") { $mode = $selection } else { $mode = "1" }

# 2. Manual Input Prompt
$manualInput = ""
if ($mode -eq "manual") {
    Write-Host ""
    Write-Host "Enter page numbers where new files should START." -ForegroundColor Yellow
    $manualInput = Read-Host "Pages (comma separated)"
}

# 3. Execution
Draw-Header
Write-Host "Running Processor..." -ForegroundColor Yellow
Write-Host "Log: $logFile" -ForegroundColor Gray
Write-Host ""

try {
    $result = Run-PythonSplitter -mode $mode -manualData $manualInput
    while ($result.status -ceq 'no_plan') {
        $decision = Get-SplitDecision -Result $result
        Write-Host $decision.message -ForegroundColor Yellow
        do {
            $prompt = 'Choose [M] manual or [C] cancel'
            if ($decision.fallback_modes -contains '1') { $prompt = 'Choose [1] Level 1, [M] manual or [C] cancel' }
            $retry = Read-Host $prompt
            $decision = Get-SplitDecision -Result $result -Choice $retry
            $script:consoleLogWriter.WriteLine("[DECISION] $($decision.decision); retry mode: $($decision.retry_mode)")
            if ($decision.decision -in @('pending', 'invalid')) { Write-Host 'Choose one of the displayed options.' }
        } while ($decision.decision -in @('pending', 'invalid'))
        if ($decision.decision -eq 'cancel') { break }
        $mode = $decision.retry_mode
        $manualInput = ''
        if ($mode -eq 'manual') {
            $manualInput = Read-Host 'Pages (comma separated)'
        }
        $result = Run-PythonSplitter -mode $mode -manualData $manualInput
    }
    $exitCode = $result.exit_code
    if ($result.status -ceq 'success') {
        foreach ($warning in $result.warnings) {
            if ($warning.code -ceq 'output_handle_close_failed') { Write-Host ('[WARNING] ' + $warning.message) -ForegroundColor Yellow }
        }
        Write-Host ('Output: ' + $result.execution.final_directory) -ForegroundColor Cyan
        Write-Host 'Done.' -ForegroundColor Cyan
    }
    else { Write-Host (Get-SplitDecision -Result $result).message -ForegroundColor Red }
}
catch {
    Write-Host "[!] Processor failure: $($_.Exception.Message)" -ForegroundColor Red
    $exitCode = 1
}
finally { $script:consoleLogWriter.Dispose() }
Write-Host ""
Pause
exit $exitCode
