# WinBookSplit — Automated PDF/AZW3/EPUB Chapter Slicer & Organizer (PowerShell + Python)
A "drop-and-forget" PDF/AZW3/EPUB decomposition tool for technical manuals, textbooks, and heavy documentation. Large PDFs are often unwieldy for e-readers or quick reference, this script treats a PDF like a modular project—dropping a massive book in and getting a clean, organized folder of individual chapters out. It combines recursive metadata parsing with a manual override to ensure every book is manageable. I use it to modularize my Cybersecurity manuals and CS textbooks for focused study sessions.

**Synopsis**
- Hybrid Modes: Auto-Discovery (Level 1/2 Bookmarks) and Manual Slicing (Custom page ranges).
- Multi-Format Support: Automatically detects AZW3/EPUB files and converts them to PDF using Calibre before processing.
- Recursive Extraction: Deep-crawls the PDF outline tree to find chapter starts and titles often missed by basic splitters.
- Smart Fallback: Automatically detects flat PDFs without bookmarks and triggers a TUI prompt for manual page entry.
- Organized Output: Cleans Windows-invalid title characters, keeps Unicode, and widens numbering to the section count (01, 02... or 001...120).
- Detailed Logging: Generates a verbose report in the output folder tracking every bookmark match, skipped page, and extraction range.
- PDF Output: Publishes a new validated run folder under Documents or the explicit output base; earlier runs and source PDFs remain unchanged.

**Requirements**
- Windows 11 x64; Windows 10 remains unverified
- Windows PowerShell 5.1 or the tested PowerShell 7.6.5
- Regular x64 CPython 3.14.8 and plain pypdf 6.19.0
- Calibre 9.15.0 for AZW3/EPUB; PDF-only processing does not require Calibre

**Nice to have**
- A basic understanding of your PDF's internal structure so you can decide between splitting by main chapters (Level 1) or sub-sections (Level 2) :)

**Files**
Place these together (e.g. C:\Tools\WinBookSplit\):
- WinBookSplit.ps1
  - Main script: handles the console menu and invokes the shipped Python engine.
- engine directory
  - Keep the shipped Python and PowerShell helpers beside the launchers.
- WinBookSplit.bat
  - Simple launcher: enables drag-and-drop functionality for PDF, AZW3, and EPUB files.
- pypdf
  - The engine: Install the hashed requirements with the selected interpreter's `-m pip` during explicit setup.
- Calibre (ebook-convert.exe) 
  - External engine: Required for ebooks; select a trusted executable with `-CalibrePath`, or use a validated standard/PATH installation.

**Installation**
1. Copy the script files to a folder of your choice, e.g.: C:\Tools\WinBookSplit\
2. Follow [the isolated setup instructions](docs/codex-v1.0.0/SUPPORT_AND_SETUP.md) to prepare a fresh application `.venv` with the exact Python/pypdf pins. Installation is an explicit setup action; processing never installs packages.
3. Install Calibre 9.15.0 separately if you process ebooks.
4. (Optional) Create a desktop shortcut to WinBookSplit.bat and name it something friendly: "Book Splitter"

**Usage**

**Recommended: Drag-and-Drop**
1. Drag a file (PDF, AZW3, or EPUB) onto the WinBookSplit.bat icon.
2. For AZW3/EPUB, Calibre generates and validates a PDF in an owned temporary workspace after you choose a split mode.
3. A window will open showing the book details and asking for a Mode:
   - Type 1 for Level 1 (Main Chapters).
   - Type 2 for Level 2 (Sub-chapters/sections).
   - Type M for Manual Mode (if you already know your page cuts).
4. Press Enter.
5. If Auto-Split fails due to no bookmarks, type Y to switch to manual mode and enter your page numbers (e.g., 15, 34, 72).
6. Find your organized chapter folder in your Documents folder.

**Command line**
Run from a PowerShell prompt:
.\WinBookSplit.ps1 -InputFile "C:\Path\To\MyBook.azw3"
You will see the same interactive menu and the same verbose log output.

For PDF runs, an existing directory can be selected as the output base:

```powershell
.\WinBookSplit.ps1 -InputFile "C:\Books\Book [1].PDF" -OutputDirectory "D:\Reading"
```

Use `-PythonPath` to select a trusted absolute interpreter path explicitly:

```powershell
.\WinBookSplit.ps1 -InputFile "C:\Books\Book.pdf" -PythonPath "C:\Tools\WinBookSplit\.venv\Scripts\python.exe"
```

Before creating console records, converting or splitting, the application checks
regular Windows x64 CPython 3.14.8 and imports pypdf 6.19.0 with that exact
interpreter. Selection prefers `-PythonPath`, the application `.venv`, then
validated launcher/PATH candidates. An invalid explicit selection or
existing broken `.venv` fails with setup guidance; it never silently substitutes
another runtime. The engine uses the same absolute path with `-I -B`, excluding
book/current-directory modules, user site packages and Python environment hooks.
The console and log report the selected executable, versions and pypdf origin.

Each engine attempt runs in an owned Windows job. Both UTF-8 streams are drained
to EOF; console diagnostics retain the last 64 KiB per stream and report any
truncation. The structured result is kept separately. `-ProcessTimeout` sets the
whole attempt deadline in seconds (1..172800); its default is at least one hour
and 30 minutes longer than `-ConversionTimeout`. Timeout or cancellation stops
only this attempt's process tree and returns a failure. Interrupted staging
directories are retained for inspection because the console does not own the
engine's cleanup ledger.

For EPUB/AZW3, select an existing converter explicitly when needed:

```powershell
.\WinBookSplit.ps1 -InputFile "C:\Books\Book.epub" -OutputDirectory "D:\Reading" -CalibrePath "C:\Calibre\ebook-convert.exe" -KeepConvertedPdf
```

Each conversion uses `--output-profile tablet`, the compatibility option tested
with Calibre 9.15.0. The original ebook and any neighboring PDF remain unchanged.
The engine requires a fresh, nonempty, readable PDF with physical pages before
planning the split. Both converter streams are drained; the log keeps up to the
last 64 KiB of each stream and reports truncation. `-ConversionTimeout` limits
each attempt in seconds (default 1800, range 1–86400); timeout or cancellation
stops the owned converter process tree before cleanup.

`-KeepConvertedPdf` stores `WinBookSplit_Converted.pdf` inside the successful
chapter folder. Its hash and page count are recorded separately from chapter
outputs in `WinBookSplit_Manifest.json`. Without the switch, the temporary full
PDF is removed after its validated bytes are captured in memory; no full PDF is
published. The manifest distinguishes the original ebook from that PDF snapshot.
Conversion errors return a nonzero status. Normal failure cleanup removes only
known owned files and leaves a separate failure record; unexpected/replaced files
are preserved with a reported cleanup failure. A hard stop can leave a marked
stage; do not delete other runs or input files when inspecting it.

Calibre is resolved and version-checked only for ebooks. A trusted absolute
`-CalibrePath` can select a nonstandard installation; automatic discovery checks
standard locations and safe absolute PATH entries. Empty/relative entries and
executables next to the book or in an unrelated current directory are excluded.
Discovery also refuses reparse paths, short-path aliases and automatic hard-linked
executables. Use an ordinary long absolute executable path or the prepared venv.
The batch launcher can use the application `.venv` and trusted converter PATH.
Explicit path and retention options are available through PowerShell above.
The acceptance books contain only original local resources. Calibre and pypdf
run with your privileges; these checks do not provide a document sandbox or
establish behavior for ebooks with remote resources. Manual preview/confirmation
and the final release workflow remain pending.

Input and output paths are handled literally. Directories used as input files,
non-filesystem providers and unreadable files are rejected with a nonzero exit.
Titles with no usable text get stable `Section 1`-style names. Destination-aware
preparation shortens titles and the run-folder stem before writing; preview and
execution retain the same filenames. A destination that cannot hold chapters
and diagnostic records is rejected with a request for a shorter output base.
The current budget is 259 UTF-16 units for a complete file path and 247 for a
created directory; this does not claim arbitrary long-path or UNC support.

**What it actually does (step-by-step)**
1. Checks & Conversion
   - Verifies the input file and checks for Python/pypdf.
   - If the input is AZW3/EPUB, it invokes Calibre’s ebook-convert to generate a source PDF.
2. Analysis (TUI)
   - Probes the PDF metadata for an internal Outline (bookmarks).
   - Shows the user the book title and total page count before processing.
3. Construct Logic
   - Auto-Mode: Recursively maps chapter titles to page indices.
   - Manual Mode: Parses user-provided page numbers into start/end indices.
   - Sanitization: Replaces illegal Windows characters in titles (e.g., "?" or ":") with underscores.
4. Processing
   - Runs the shipped Python engine and its validated split plan.
   - Executes pypdf to create new, optimized PDF slices for every detected section.
   - Shows real-time "Writing..." progress in the console.
5. Logging
   - Generates a WinBookSplit_Log.txt in the output folder.
   - Records the exact page ranges, titles found, and any errors encountered during binary extraction.

**Limitations / When not to use**
- Scanned Images: If the PDF is just photos of pages with no OCR or metadata, "Auto-Mode" will fail. Use Manual Mode instead.
- Encrypted PDFs: Files with strict Owner Passwords may prevent the script from extracting pages.
- Complex Outlines: Some PDFs have broken bookmark links, the script skips these to prevent creating corrupt or empty output files.
- Conversion Artifacts: Bookmarks can sometimes be lost or altered during AZW3 to PDF conversion. Use Manual Mode if the outline is missing after conversion.

**Troubleshooting**
- "Calibre not found" 
  - Select trusted Calibre 9.15.0 `ebook-convert.exe` with `-CalibrePath`, or provide a validated standard/PATH installation. PDF-only processing skips this check.
- "AUTO-SPLIT FAILED: No Bookmarks Found"
  - The PDF has no internal Table of Contents metadata. Type "Y" when prompted to enter page numbers manually.
- "Python not found"
  - Select regular Windows x64 CPython 3.14.8 with `-PythonPath`, or follow the isolated setup instructions for the application `.venv`.
- "ModuleNotFoundError: No module named 'pypdf'"
  - Use the exact selected interpreter's `-m pip install --require-hashes --only-binary=:all: -r requirements.txt` during explicit setup; avoid an unrelated `pip` command.

**Intent & License**
Personal helper for modularizing heavy technical documentation and CS textbooks. "I just want to read Chapter 5 on my tablet without loading a 500MB PDF." Provided as-is, without warranty. Use at your own risk. Feel free to modify the logic to fit your specific study workflow.
