# Windows 10/11 implementation audit

Target Windows 10 **22H2 x64** (build 19045 with recorded servicing revision) and supported/tested Windows 11 x64 builds. Keep inbox Windows PowerShell 5.1 as the BAT baseline. PowerShell 7 remains a separately tested direct-entry option. Do not replace the launcher with a PowerShell-7-only command.

Read actual `WinBookSplit.ps1`, BAT, Runtime/Process/Paths/Logging/Diagnostics helpers, Python job/process helpers and packaging/setup. Audit every relevant P/Invoke, process job assignment/cancellation, file-handle/reparse check, console API, encoding operation and PowerShell/.NET method against the oldest target. Avoid changes unless a real compatibility issue is found. Do not disable process-tree ownership, cleanup validation, quoting or Unicode handling to accommodate Windows 10.

Do not gate on a registry ProductName string alone. Both OS families use version 10.0 and registry names can be stale. For test evidence capture native OS caption, build, architecture and ProductType; ProductType must identify a workstation. Treat Windows Server, compatibility-mode strings and mocked OS probes as distinct from native Windows 10. No production hard block on servicing revision unless necessary and documented.

Run full native smoke from a fresh extraction in a standard-user-writable ordinary local path, including spaces/brackets/non-ASCII characters and an unrelated current directory. The app's existing network-path/long-path exclusions stay unchanged. Verify `-Version`/help do not require installed Python or pypdf.

Keep current Calibre 9.15.0 as the starting converter target and validate actual EPUB/AZW3 conversion on Windows 10/11. New converter version acceptance is not a mandatory feature; fix only a demonstrated blocker. PDF-only use never requires Calibre. Keep source files and neighboring same-name PDF bytes unchanged.

Upstream compatibility and Microsoft servicing are separate topics. Document the end of ordinary Windows 10 Home/Pro support (14 October 2025, [S9]) without claiming that the application supplies OS security updates or refusing solely because Windows 10 is old. No forced OS upgrade, reinstall or paid ESU prerequisite imposed by this application patch.
