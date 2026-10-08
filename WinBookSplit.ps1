<#
WinBookSplit.ps1
Automated PDF/AZW3/EPUB Chapter Slicer & Organizer

Author: Janne Vuorela
Target OS: Windows 11
PowerShell: Windows PowerShell 5.1 or PowerShell 7+
Dependencies: Python 3.x, pypdf library, Calibre (for AZW3/EPUB), .bat wrapper

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
        Prefixes files with two-digit padding to maintain correct chronological order.
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
    1) Install Python engine: Run 'pip install pypdf' in your terminal.
    2) Install Calibre: Required for AZW3/EPUB conversion support.
    3) Create a folder (e.g., C:\Tools\WinBookSplit\).
    4) Place the .ps1, .bat, and engine directory inside.
    5) (Optional) Create a desktop shortcut to WinBookSplit.bat.

USAGE
    A) Drag-and-Drop (Recommended)
        - Drag any PDF, AZW3, or EPUB file onto WinBookSplit.bat.
        - Follow the TUI prompts for conversion and splitting.

    B) Direct PowerShell
        - Run: .\WinBookSplit.ps1 -InputFile "C:\Path\To\Book.azw3"

NOTES
    - The script uses the shipped engine\winbooksplit_engine.py beside the
      PowerShell entry point, independently of the current working directory.
    - Output is always directed to a sub-folder in the 'Documents' directory.

LIMITATIONS
    - Requires a local Python installation with the 'pypdf' library.
    - Requires Calibre installed in standard paths for AZW3/EPUB conversion.
    - Manual mode assumes every provided page number is the start of a new chapter.

TROUBLESHOOTING
    - "Calibre not found":
        Ensure Calibre is installed to process non-PDF ebook formats.
    - "[!!!] ERROR: No bookmarks found":
        The PDF lacks an internal Outline. Use Manual Mode [M] instead.
    - "Python not found":
        Ensure Python is added to your Windows PATH environment variable.

LICENSE / WARRANTY
    - Personal IT automation tool, provided as-is.
#>

param (
    [string]$InputFile
)

# --- Configuration ---
$AppName = "WinBookSplit"
$ver = "2.1 (AZW3 Support)"
$documentsPath = [Environment]::GetFolderPath("MyDocuments")

# Standard locations for Calibre's converter tool
$CalibreSearchPaths = @(
    "C:\Program Files\Calibre2\ebook-convert.exe",
    "C:\Program Files (x86)\Calibre2\ebook-convert.exe",
    "$env:LOCALAPPDATA\Programs\Calibre\ebook-convert.exe"
)

# --- UI Functions ---
function Draw-Header {
    Clear-Host
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host "      $AppName v$ver" -ForegroundColor Yellow
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host ""
}

# --- Validation ---
Draw-Header
# Now accepts AZW3 and EPUB as valid inputs
if (-not (Test-Path $InputFile) -or $InputFile -notmatch "\.pdf$|\.azw3$|\.epub$") {
    Write-Host "[!] Error: Invalid file." -ForegroundColor Red
    Write-Host "Please drop a valid PDF, AZW3, or EPUB file." -ForegroundColor Gray
    Read-Host "Press Enter to exit"
    exit
}

$inputExt = [System.IO.Path]::GetExtension($InputFile).ToLower()

# --- Pre-process: Format Conversion, if needed ---
# If it's not a PDF, we must convert it first.
if ($inputExt -ne ".pdf") {
    Write-Host "Detected $inputExt format." -ForegroundColor Yellow
    Write-Host "Locating conversion engine (Calibre)..." -ForegroundColor Gray

    $converterExe = $CalibreSearchPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $converterExe) {
        Write-Host ""
        Write-Host "[!] CONVERSION ERROR: Calibre not found." -ForegroundColor Red
        Write-Host "To process AZW3/EPUB files, Calibre must be installed." -ForegroundColor White
        Write-Host "Please install Calibre and try again." -ForegroundColor Gray
        Write-Host ""
        Read-Host "Press Enter to exit"
        exit
    }

    Write-Host "Engine found: $converterExe" -ForegroundColor Gray
    
    # Define the path for the new full PDF, saved next to the source file
    $convertedPdfPath = [System.IO.Path]::ChangeExtension($InputFile, ".pdf")
    
    Write-Host ""
    Write-Host "--- STARTING CONVERSION ---" -ForegroundColor Cyan
    Write-Host "Input:  $InputFile" -ForegroundColor White
    Write-Host "Output: $convertedPdfPath" -ForegroundColor White
    Write-Host "This may take a minute depending on book size..." -ForegroundColor Yellow
    Write-Host ""

    # Run calibre converter
    # We use --output-profile tablet for better PDF margins/scaling
    $convertArgs = "`"$InputFile`" `"$convertedPdfPath`" --output-profile tablet"
    
    # Run the process directly so the user sees Calibre's output stream
    $proc = Start-Process -FilePath $converterExe -ArgumentList $convertArgs -Wait -NoNewWindow -PassThru

    if ($proc.ExitCode -ne 0 -or -not (Test-Path $convertedPdfPath)) {
         Write-Host ""
         Write-Host "[!] CALIBRE CONVERSION FAILED." -ForegroundColor Red
         Read-Host "Press Enter to exit"
         exit
    }

    Write-Host ""
    Write-Host "[V] Conversion complete." -ForegroundColor Green
    Start-Sleep -Seconds 1

    # CRITICAL STEP, update the InputFile variable so the rest of the script operates on the newly created PDF.
    $InputFile = $convertedPdfPath
    # Refresh header to clear conversion clutter
    Draw-Header
}


# --- Preparation (Post-Conversion) ---
# We calculate stats here so they reflect the PDF, not the source AZW3
$fileName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
$fileSize = "{0:N2} MB" -f ((Get-Item $InputFile).Length / 1MB)

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputDir = Join-Path $documentsPath "$fileName`_Chapters"
$logFile = Join-Path $outputDir "WinBookSplit_Log_$timestamp.txt"

if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# --- The Python Engine ---
$enginePath = Join-Path $PSScriptRoot 'engine\winbooksplit_engine.py'

# --- Execution Function ---
function Run-PythonSplitter ($mode, $manualData) {
    $pInfo = New-Object System.Diagnostics.ProcessStartInfo
    $pInfo.FileName = "python"
    $argString = "`"$enginePath`" `"$InputFile`" `"$outputDir`" `"$mode`" `"$manualData`""
    $pInfo.Arguments = $argString
    $pInfo.RedirectStandardOutput = $true
    $pInfo.RedirectStandardError = $true
    $pInfo.UseShellExecute = $false
    $pInfo.CreateNoWindow = $true
    $pInfo.StandardOutputEncoding = [System.Text.Encoding]::UTF8

    $p = New-Object System.Diagnostics.Process
    $p.StartInfo = $pInfo
    $p.Start() | Out-Null

    # Stream Output
    while (-not $p.HasExited) {
        while ($line = $p.StandardOutput.ReadLine()) {
            if ($line -eq "[NO_BOOKMARKS_FOUND]") { continue }
            Write-Host $line -ForegroundColor Green
            $line | Out-File -FilePath $logFile -Append -Encoding UTF8
        }
        Start-Sleep -Milliseconds 50
    }
    return $p.ExitCode
}

# --- Main Logic Flow ---

# 1. Initial TUI, shows stats for the PDF, even if converted
Write-Host "Target Book: $fileName" -ForegroundColor Green
Write-Host "Size:        $fileSize" -ForegroundColor Gray
if ($inputExt -ne ".pdf") {
     Write-Host "Format:      Converted from $inputExt" -ForegroundColor DarkGray
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

$exitCode = Run-PythonSplitter -mode $mode -manualData $manualInput

# 4. Handle "No Bookmarks" Failure (Exit Code 55)
if ($exitCode -eq 55) {
    Write-Host ""
    Write-Host "------------------------------------------------" -ForegroundColor Red
    Write-Host "[!] AUTO-SPLIT FAILED: No Bookmarks Found." -ForegroundColor Red
    Write-Host "------------------------------------------------" -ForegroundColor Red
    Write-Host "The PDF has no internal chapter structure."
    if ($inputExt -ne ".pdf") {
         Write-Host "Note: Bookmarks are often lost during AZW3->PDF conversion." -ForegroundColor Gray
    }
    Write-Host ""
    
    $retry = Read-Host "Do you want to switch to MANUAL mode? (Y/N)"
    if ($retry -match "y|Y") {
        Write-Host ""
        Write-Host "Enter page numbers where new files should START." -ForegroundColor Yellow
        $manualInput = Read-Host "Pages (comma separated)"
        
        Write-Host "Retrying in Manual Mode..." -ForegroundColor Yellow
        Run-PythonSplitter -mode "manual" -manualData $manualInput
    }
}

Write-Host ""
Write-Host "Done." -ForegroundColor Cyan
Pause