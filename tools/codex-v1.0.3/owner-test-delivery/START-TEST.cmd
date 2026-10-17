@echo off
setlocal DisableDelayedExpansion
title WinBookSplit Windows 10 test
echo WinBookSplit Windows 10 test
echo Keep this window open while the checks run.
echo.
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0Start-OwnerTest.ps1"
set "owner_test_exit=%errorlevel%"
echo.
echo The test process has finished. Follow the report instructions above.
echo If no report was created, send the short error message shown above.
echo.
pause
exit /b %owner_test_exit%
