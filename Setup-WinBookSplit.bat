@echo off
setlocal DisableDelayedExpansion
:: Explicit local setup; forward script arguments once without CALL or shell code.

set "WinBookSplitSetupPowerShell=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
:: Sysnative reaches the inbox x64 host from a 32-bit command processor.
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" set "WinBookSplitSetupPowerShell=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%WinBookSplitSetupPowerShell%" (
    echo Windows PowerShell was not found in its Windows system location. 1>&2
    exit /b 3
)

"%WinBookSplitSetupPowerShell%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Setup-WinBookSplit.ps1" %*
set "WinBookSplitSetupExitCode=%ERRORLEVEL%"
if "%WinBookSplitSetupExitCode%"=="1" echo PowerShell could not complete setup; review the error above. Enforced script policies still apply. 1>&2
exit /b %WinBookSplitSetupExitCode%
