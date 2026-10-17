# Confirmed incident and evidence boundary

## Source material

The owner supplied `WinBookSplit-discovery-20261017-122739.txt` (v1.0.2).
`fixtures/discovery-report.sanitized.txt` preserves its text except for replacing
the local account name with `WBS_USER` and normalizing line endings. The original
file is NOT included. `fixtures/incident.json` is a structured derived record.
Source line references below refer to the original displayed diagnostic.

## Observed, not hypothetical

- No application `.venv` existed (lines 1-3).
- `python` resolves to `%LOCALAPPDATA%\Python\bin\python.exe` in PS7 and the
  PS5.1 command listing (33-46, 65-78).
- The successful PS7 probe launches the concrete runtime under
  `%LOCALAPPDATA%\Python\pythoncore-3.15-64\python.exe`, reports regular x64
  CPython 3.15.0, and exits 0 (49-58).
- Current-user manager registration has family
  `PythonSoftwareFoundation.PythonManager_qbz5n2kfra8p0`, publisher
  `CN=4975D53F-AA7E-49A5-8B49-EA4FDC1BB66B`, signature kind `Store`,
  architecture `X64`, status `Ok`; observed version is 26.4.240.0 (113-121).
- In the ORIGINAL failed run, automatic PATH discovery rejects the hard-linked
  Python entry point without a probe (130-134). The publisher-filtered manager
  query exits 0 with `packages: []` (137-170). PATH pymanager is then excluded as
  unverified (176-180), and PATH py is rejected as a reparse point (183-187).
- The application returns `runtime_not_found` (190-191).

## Confirmed code defects / causal connection

`engine/WinBookSplit.Runtime.ps1` in v1.0.2 filters `Get-AppxPackage` to the
python.org publisher. The Store package therefore never enters validation.
`ConvertFrom-WinBookSplitManagerRegistration` independently requires the
python.org publisher/family `_3847v3x7pw1km` and `SignatureKind=Developer`.
Changing only the query still fails at validation.

PATH aliases are deliberately rejected as ordinary runtimes, but there is no
working official Store-manager route to resolve them. The real interpreter is
never reached. `Get-WinBookSplitDependencySummary` in `WinBookSplit.ps1` can
collapse uncategorized discovery rejections to `missing_python`. That is not an
accurate account of the recorded candidate paths.

These are discovery and diagnostics defects. The v1.0.2 Python policy already
accepts regular stable x64 3.15.0. No new version-policy broadening is required.

## What is NOT established

The report does not establish a pypdf installation in the concrete base runtime.
It does not prove the final Win10 application will work after these edits; that
requires the native gate. The PS5.1 nested Python probe in the old collector
failed with PowerShell error-stream/XML handling messages (83-109); do NOT count
that probe as a pass or as proof Python cannot run in PS5.1. The four local file
hashes are recorded fingerprints, not a proven match to the published archive.
Codex must compare exact Git/release bytes, explaining any line-ending changes.

The original Windows 10 release test selected a reviewed private runtime after
owner discovery failed. Source records `V102-11-windows10-review.json` and
`V102-11-windows10.json` record this. That was execution evidence for a supplied
runtime, not a fix of this owner's installation route. Preserve those records;
add a corrected gate, do not rewrite old passes to mean something they did not.

## Repair target

Recognize the installed manager using paired official identities; invoke a
supported, verified manager entry point; list only installed runtimes; validate
the returned concrete interpreter; then use normal setup and the existing engine.
Keep the desktop installation unchanged as the regression environment.
