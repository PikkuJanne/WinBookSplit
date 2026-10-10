@echo off
setlocal DisableDelayedExpansion
:: WinBookSplit Launcher
:: Drop one PDF, EPUB or AZW3, or double-click to enter its literal path.

if not "%~2"=="" goto WinBookSplitMultipleFiles
:: Only absent or quoted-empty argument 2 can reach this raw-presence check.
if not "%2"=="" goto WinBookSplitMultipleFiles

:: Use the inbox shell; a book/CWD/PATH must not supply powershell.cmd/exe.
set "WinBookSplitPowerShell=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%WinBookSplitPowerShell%" (
    echo Windows PowerShell was not found in its Windows system location.
    exit /b 3
)

:: Paths stay quoted data; the script prompts when no file was dropped.
if "%~1"=="" (
    "%WinBookSplitPowerShell%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0WinBookSplit.ps1"
) else (
    "%WinBookSplitPowerShell%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0WinBookSplit.ps1" "%~1"
)
set "WinBookSplitExitCode=%ERRORLEVEL%"
exit /b %WinBookSplitExitCode%

:WinBookSplitMultipleFiles
echo Please select one PDF, EPUB or AZW3 file at a time. Multiple files are not supported.
pause
exit /b 2
