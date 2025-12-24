<#
WinBookSplit.ps1
Automated PDF Chapter Slicer & Organizer

Author: Janne Vuorela
Target OS: Windows 11
PowerShell: Windows PowerShell 5.1 or PowerShell 7+
Dependencies: Python 3.x, pypdf library (pip install pypdf), .bat wrapper

SYNOPSIS
    A "smart" decomposition tool for technical manuals and textbooks.
    Automates the tedious process of splitting large PDF files into individual 
    chapters based on internal metadata (bookmarks) or manual page ranges.

WHAT THIS IS (AND ISN'T)
    - A logic-driven parser for PDF structures.
    - Designed for CS students/IT Support who need to modularize heavy documentation.
    - Hybrid tool, attempts Auto-Discovery first, falls back to Manual Slicing.
    - Not an OCR engine, it cannot read text on a flat image to find chapters.
    - Not a PDF editor, it creates new files and does not modify the source.

FEATURES
    - Text User Interface (TUI):
        Dynamic console interface providing real-time feedback and file stats.
        Includes a Bookmark Depth selector for Level 1 (Chapters) or Level 2 (Sub-sections).
    - Intelligent Metadata Extraction:
        Recursively crawls the PDF outline tree to map chapter starts and ends.
        Automatically handles page-indexing offsets (0-based vs 1-based).
    - Smart Manual Fallback:
        If a PDF lacks metadata ("No Bookmarks Found"), the TUI triggers an 
        interactive manual mode allowing the user to input custom split points.
    - Automated Sanitization:
        Cleans illegal Windows characters ( \ / : * ? " < > | ) from chapter titles 
        to ensure valid and clean filenames.
    - Systematic Naming:
        Prefixes files with two-digit padding (01, 02...) to maintain 
        correct chronological order in Windows File Explorer.
    - Verbose Logging:
        Generates a detailed execution log in the output folder. 
        Records every matched bookmark, skipped page, and filesystem action.

MY INTENDED USAGE
    - I keep WinBookSplit in my Tools directory with a shortcut to the .bat on my desktop.
    - When I download a massive manual or a textbook:
        1. I drag the PDF onto the .bat launcher.
        2. I attempt a "Level 1" split for main chapters.
        3. If the book is "flat" (no bookmarks), I switch to manual mode and 
           input the page numbers from the printed Table of Contents.
        4. I find a clean, numbered folder in my Documents, ready for my tablet or e-reader.

SETUP
    1) Install the engine: Run 'pip install pypdf' in your terminal.
    2) Create a folder (e.g., C:\Users\user\Tools\WinBookSplit\).
    3) Place these two files inside:
        - WinBookSplit.ps1
        - WinBookSplit.bat
    4) (Optional) Create a desktop shortcut to WinBookSplit.bat.

USAGE
    A) Drag-and-Drop (Recommended)
        - Drag any PDF file onto WinBookSplit.bat.
        - Choose [1], [2], or [M] for manual.

    B) Direct PowerShell
        - Run: .\WinBookSplit.ps1 -InputFile "C:\Path\To\Book.pdf"

NOTES
    - The script utilizes a Python Engine script generated on-the-fly in the 
      $TEMP directory to ensure maximum portability and logic consistency.
    - Output is always directed to a sub-folder in the 'Documents' directory 
      to prevent desktop clutter.

LIMITATIONS
    - Requires a local Python installation with the 'pypdf' library.
    - Manual mode assumes every provided page number is the start of a new chapter.

TROUBLESHOOTING
    - "[!!!] ERROR: No bookmarks found":
        The PDF lacks an internal Outline. Use Manual Mode [M] instead.
    - "Python not found":
        Ensure Python is added to your Windows PATH environment variable.
    - Files not appearing:
        Check the log file created in Documents\<FileName>_Chapters\ for 
        specific traceback errors.

LICENSE / WARRANTY
    - Personal IT automation tool, provided as-is.
#>

param (
    [string]$InputFile
)

# --- Configuration ---
$AppName = "WinBookSplit"
$ver = "2.0"
$documentsPath = [Environment]::GetFolderPath("MyDocuments")

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
if (-not (Test-Path $InputFile) -or $InputFile -notmatch "\.pdf$") {
    Write-Host "[!] Error: Invalid file. Please drop a valid PDF." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit
}

$fileName = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
$fileSize = "{0:N2} MB" -f ((Get-Item $InputFile).Length / 1MB)

# --- Preparation ---
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputDir = Join-Path $documentsPath "$fileName`_Chapters"
$logFile = Join-Path $outputDir "WinBookSplit_Log_$timestamp.txt"

if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# --- THE PYTHON ENGINE ---
$pyScriptContent = @"
import sys
import os
import re
from pypdf import PdfReader, PdfWriter

# Flush stdout for realtime PS logging
sys.stdout.reconfigure(line_buffering=True)

def log(msg): print(msg)

def split_pdf(input_path, output_dir, mode, manual_data=None):
    try:
        reader = PdfReader(input_path)
    except Exception as e:
        log(f"[CRITICAL] Could not read PDF: {e}")
        sys.exit(1)

    total_pages = len(reader.pages)
    
    # --- AUTO MODE, Bookmarks ---
    if mode in ['1', '2']:
        target_level = int(mode)
        if not reader.outline:
            log("[NO_BOOKMARKS_FOUND]") # Signal to PowerShell
            sys.exit(55) # Custom exit code for 'No Bookmarks'

        bookmarks_found = []
        def extract(nodes, depth=1):
            for node in nodes:
                if isinstance(node, list):
                    extract(node, depth + 1)
                else:
                    if depth == target_level:
                        try:
                            pg = reader.get_destination_page_number(node)
                            if pg is not None and pg != -1:
                                bookmarks_found.append({'page': pg, 'title': node.title or "Untitled"})
                        except: pass
        
        extract(reader.outline)
        
        if not bookmarks_found:
            log("[NO_BOOKMARKS_FOUND]")
            sys.exit(55)

        # Process Auto Split
        bookmarks_found.sort(key=lambda x: x['page'])
        unique_map = []
        seen = set()
        for b in bookmarks_found:
            if b['page'] not in seen:
                unique_map.append(b)
                seen.add(b['page'])
        
        log(f"[*] Found {len(unique_map)} chapters based on bookmarks.")
        
        for i, mark in enumerate(unique_map):
            start = mark['page']
            end = unique_map[i+1]['page'] if i+1 < len(unique_map) else total_pages
            
            if start >= end: continue
            
            safe_title = re.sub(r'[<>:""/\\|?*]', '', mark['title']).strip()[:50]
            fname = f"{i+1:02d} - {safe_title}.pdf"
            
            write_slice(reader, start, end, os.path.join(output_dir, fname))

    # --- MANUAL MODE, Page List ---
    elif mode == 'manual':
        # Parse "13, 22, 54"
        try:
            raw_nums = [int(x.strip()) for x in manual_data.split(',') if x.strip().isdigit()]
        except:
            log("[ERROR] Invalid number format.")
            sys.exit(1)
            
        raw_nums.sort()
        
        # Logic: If user says "13", they usually mean "Split HERE".
        
        split_indices = []
        
        # Always start at the beginning (Index 0 / Page 1)
        if 1 not in raw_nums:
            split_indices.append(0)
            
        for p in raw_nums:
            # User input "13" -> Index 12
            idx = p - 1 
            if idx > 0 and idx < total_pages:
                split_indices.append(idx)
        
        # Remove duplicates and sort
        split_indices = sorted(list(set(split_indices)))
        
        log(f"[*] Manual Split Points (Page #): {[x+1 for x in split_indices]}")
        
        for i, start_idx in enumerate(split_indices):
            if i + 1 < len(split_indices):
                end_idx = split_indices[i+1]
            else:
                end_idx = total_pages
            
            # Naming for manual: 01 - Section (Pages X-Y).pdf
            fname = f"{i+1:02d} - Section (Page {start_idx+1}-{end_idx}).pdf"
            write_slice(reader, start_idx, end_idx, os.path.join(output_dir, fname))

def write_slice(reader, start, end, out_path):
    log(f"    [Writing] {os.path.basename(out_path)}")
    writer = PdfWriter()
    for p in range(start, end):
        writer.add_page(reader.pages[p])
    with open(out_path, "wb") as f:
        writer.write(f)

if __name__ == "__main__":
    # Args: Script, InputFile, OutputDir, Mode, ManualData
    # Mode: '1', '2', or 'manual'
    mode = sys.argv[3]
    manual_data = sys.argv[4] if len(sys.argv) > 4 else None
    
    split_pdf(sys.argv[1], sys.argv[2], mode, manual_data)
"@

$tempPyFile = "$env:TEMP\WinBookSplit_Engine_v2.py"
$pyScriptContent | Out-File -FilePath $tempPyFile -Encoding UTF8

# --- Execution Function ---
function Run-PythonSplitter ($mode, $manualData) {
    $pInfo = New-Object System.Diagnostics.ProcessStartInfo
    $pInfo.FileName = "python"
    # Arguments: Script Input Output Mode ManualData
    $argString = "`"$tempPyFile`" `"$InputFile`" `"$outputDir`" `"$mode`" `"$manualData`""
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
            if ($line -eq "[NO_BOOKMARKS_FOUND]") { continue } # Suppress internal flag
            Write-Host $line -ForegroundColor Green
            $line | Out-File -FilePath $logFile -Append -Encoding UTF8
        }
        Start-Sleep -Milliseconds 50
    }
    
    # Check Exit Code
    return $p.ExitCode
}

# --- Main Logic Flow ---

# --- Initial TUI ---
Draw-Header
Write-Host "Target Book: $fileName" -ForegroundColor Green
Write-Host "Size:        $fileSize" -ForegroundColor Gray
Write-Host ""
Write-Host "Select Splitting Method:" -ForegroundColor White
Write-Host " [1] Level 1 Bookmarks (Auto)" -ForegroundColor Cyan
Write-Host " [2] Level 2 Bookmarks (Auto)" -ForegroundColor Cyan
Write-Host " [M] Manual Entry (Page Numbers)" -ForegroundColor Magenta
Write-Host ""

$selection = Read-Host "Enter selection"

if ($selection -match "m|M") {
    $mode = "manual"
} elseif ($selection -eq "1" -or $selection -eq "2") {
    $mode = $selection
} else {
    $mode = "1" # Default
}

# --- Manual Input Prompt, if selected immediately ---
$manualInput = ""
if ($mode -eq "manual") {
    Write-Host ""
    Write-Host "Enter page numbers where new files should START." -ForegroundColor Yellow
    Write-Host "Example: '13, 50, 88' creates files for 1-12, 13-49, 50-87..." -ForegroundColor Gray
    $manualInput = Read-Host "Pages (comma separated)"
}

# --- Execution ---
Draw-Header
Write-Host "Running Processor..." -ForegroundColor Yellow
Write-Host "Log: $logFile" -ForegroundColor Gray
Write-Host ""

$exitCode = Run-PythonSplitter -mode $mode -manualData $manualInput

# --- Handle "No Bookmarks" Failure, Exit Code 55 ---
if ($exitCode -eq 55) {
    Write-Host ""
    Write-Host "------------------------------------------------" -ForegroundColor Red
    Write-Host "[!] AUTO-SPLIT FAILED: No Bookmarks Found." -ForegroundColor Red
    Write-Host "------------------------------------------------" -ForegroundColor Red
    Write-Host "The PDF has no internal chapter structure."
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

# --- Cleanup ---
Remove-Item $tempPyFile -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "Done." -ForegroundColor Cyan
Pause
