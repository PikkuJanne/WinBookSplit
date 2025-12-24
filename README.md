# WinBookSplit — Automated PDF Chapter Slicer & Organizer (PowerShell + Python)

A "drop-and-forget" PDF decomposition tool for technical manuals, textbooks, and heavy documentation. Large PDFs are often unwieldy for e-readers or quick reference, this script treats a PDF like a modular project—dropping a massive book in and getting a clean, organized folder of individual chapters out. It combines recursive metadata parsing with a manual override to ensure every book is manageable. I use it to modularize my Cybersecurity manuals and CS textbooks for focused study sessions.

**Synopsis**
- Hybrid Modes: Auto-Discovery (Level 1/2 Bookmarks) and Manual Slicing (Custom page ranges).
- Recursive Extraction: Deep-crawls the PDF outline tree to find chapter starts and titles often missed by basic splitters.
- Smart Fallback: Automatically detects flat PDFs without bookmarks and triggers a TUI prompt for manual page entry.
- Organized Output: Cleans illegal characters from titles and applies two-digit padding (01, 02...) to keep files perfectly ordered.
- Detailed Logging: Generates a verbose report in the output folder tracking every bookmark match, skipped page, and extraction range.
- Non-Destructive: Always creates a new folder in Documents, never modifies the original source file.

**Requirements**
- Windows 10 or 11
- Windows PowerShell 5.1 (built-in) or PowerShell 7+
- Python 3.x (Must be installed and added to Windows PATH)
- pypdf library (pip install pypdf)

**Nice to have**
- A basic understanding of your PDF's internal structure so you can decide between splitting by main chapters (Level 1) or sub-sections (Level 2) :)

**Files**
Place these together (e.g. C:\Tools\WinBookSplit\):
- WinBookSplit.ps1
  - Main script: handles the TUI, calculates page ranges, manages the hybrid logic, and generates the Python engine on-the-fly.
- WinBookSplit.bat
  - Simple launcher: enables drag-and-drop functionality for PDF files.
- pypdf
  - The engine: This Python library handles the heavy lifting of PDF binary manipulation. Install via pip.

**Installation**
1. Copy the script files to a folder of your choice, e.g.: C:\Tools\WinBookSplit\
2. Open your terminal and run: pip install pypdf
3. (Optional) Create a desktop shortcut to WinBookSplit.bat and name it something friendly: "Book Splitter"

**Usage**

**Recommended: Drag-and-Drop**
1. Drag a PDF file (Textbook, Manual, or Report) onto the WinBookSplit.bat icon.
2. A window will open showing the book details and asking for a Mode:
   - Type 1 for Level 1 (Main Chapters).
   - Type 2 for Level 2 (Sub-chapters/sections).
   - Type M for Manual Mode (if you already know your page cuts).
3. Press Enter.
4. If Auto-Split fails due to no bookmarks, type Y to switch to manual mode and enter your page numbers (e.g., 15, 34, 72).
5. Find your organized chapter folder in your Documents folder.

**Command line**
Run from a PowerShell prompt:
.\WinBookSplit.ps1 -InputFile "C:\Path\To\MyBook.pdf"
You will see the same interactive menu and the same verbose log output.

**What it actually does (step-by-step)**
1. Checks
   - Verifies the input file is a valid PDF.
   - Ensures the Python environment and pypdf are accessible.
2. Analysis (TUI)
   - Probes the PDF metadata for an internal Outline (bookmarks).
   - Shows the user the book title and total page count before processing.
3. Construct Logic
   - Auto-Mode: Recursively maps chapter titles to page indices.
   - Manual Mode: Parses user-provided page numbers into start/end indices.
   - Sanitization: Replaces illegal Windows characters in titles (e.g., "?" or ":") with underscores.
4. Processing
   - Generates a temporary Python Engine script.
   - Executes pypdf to create new, optimized PDF slices for every detected section.
   - Shows real-time "Writing..." progress in the console.
5. Logging
   - Generates a WinBookSplit_Log.txt in the output folder.
   - Records the exact page ranges, titles found, and any errors encountered during binary extraction.

**Limitations / When not to use**
   - Scanned Images: If the PDF is just photos of pages with no OCR or metadata, "Auto-Mode" will fail. Use Manual Mode instead.
   - Encrypted PDFs: Files with strict Owner Passwords may prevent the script from extracting pages.
   - Complex Outlines: Some PDFs have broken bookmark links, the script skips these to prevent creating corrupt or empty output files.

**Troubleshooting**
- "AUTO-SPLIT FAILED: No Bookmarks Found"
  - The PDF has no internal Table of Contents metadata. Type "Y" when prompted to enter page numbers manually.
- "Python not found"
  - Ensure Python is installed and the "Add to PATH" checkbox was ticked during installation.
- "ModuleNotFoundError: No module named 'pypdf'"
  - The required library is missing. Run 'pip install pypdf' in your command prompt.

**Intent & License**
Personal helper for modularizing heavy technical documentation and CS textbooks. "I just want to read Chapter 5 on my tablet without loading a 500MB PDF." Provided as-is, without warranty. Use at your own risk. Feel free to modify the logic to fit your specific study workflow.