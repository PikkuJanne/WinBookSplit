@echo off
:: WinBookSplit Launcher
:: Drag and Drop a PDF onto this file to start.

if "%~1"=="" (
    echo No file dropped. Please drag a PDF file onto this batch file.
    pause
    exit /b
)

:: Launch PowerShell script with the dropped file path
PowerShell -NoProfile -ExecutionPolicy Bypass -File "%~dp0WinBookSplit.ps1" "%~1"