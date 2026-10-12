WinBookSplit v1.0.1 fixes Python discovery and setup while preserving the Windows BAT shortcut, console, single-file drag-and-drop, PDF splitting and optional EPUB/AZW3 conversion.

- Supports regular GIL Windows x64 CPython **3.13.16 or later stable 3.13 patches**, and **3.14.8 or later stable 3.14 patches**. These are numeric security floors, not exact-patch requirements. The exact tested Python versions are 3.13.16 and 3.14.8; every later patch has not been tested. Python 3.11/3.12, other minors, x86, ARM, free-threaded Python and PyPy are outside this release policy.
- Keeps **pypdf 6.20.0** separately hash-locked, with isolated imports and dependency-origin checks. Calibre **9.15.0** is separately installed and required only for EPUB or unencrypted AZW3 input.
- Adds explicit application-local setup, read-only readiness checks, offline reviewed-wheel support and repair that preserves the old environment. Processing never installs packages or replaces a global Python installation.
- A failed optional launcher listing preserves discovery of a valid trusted PATH interpreter. Automatic selection is deterministic. Explicit Python selections and an existing application `.venv` remain authoritative.
- Dependency failures give a concise reason and a concrete setup action, with full local technical diagnostics retained.

Extract the complete ZIP to a short writable per-user directory. Install compatible regular x64 Python separately, then run this once from the application directory:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Setup-WinBookSplit.ps1 -NonInteractive
```

The execution-policy switch affects only that PowerShell process. No elevation, global PATH change, environment activation or global package installation is needed. Organization policy still applies. PDF input needs no Calibre. See the included [setup guide](https://github.com/PikkuJanne/WinBookSplit/blob/v1.0.1/docs/SETUP.md) for explicit `-PythonPath`, `-Check`, offline wheels and `-Repair`.

The same frozen ZIP was exercised on native Windows10 22H2 x64 with Python3.14.8 and inbox PowerShell5.1.19041.7417 through a USER_LAUNCHED automated kit, including actual EPUB/AZW3 conversion; Windows11 x64 with Python3.13.16 and3.14.8/inbox5.1; and four supplemental same-extraction cases under PowerShell7.6.6. Source CI runs on Windows Server and is separate from workstation acceptance. The unchanged Explorer route was not repeated; BAT argument forwarding was automated. No repeated manual fidelity programme or clean-OS claim is made.

Windows 10 Home/Pro 22H2 standard servicing ended on 14 October 2025. OS security servicing is separate from application compatibility; this tool does not enroll a machine in servicing or require a purchase.

Calibre 9.15.0 retains older embedded libraries. Its reviewed EPUB/AZW3 conversion path and published advisory ranges do not establish that every converter/plugin path is vulnerability-free. Process deadlines are not an OS/network sandbox; see the included SECURITY.md for limits.

Scripts and the ZIP are **unsigned**. SHA256SUMS detects changes against the published values; a hash is not a publisher signature. Keep accepted Python branches serviced with current security updates. Existing document, source-preservation, process and output-safety policies remain in force. Encrypted/DRM content and the documented unsupported PDF structures remain excluded; there is no new OCR, GUI, telemetry, uploader or ebook-output feature.

Release source: `6f2c2e5d4d5fc484bcb1530474f09957e09775ef`. ZIP SHA256: `747d578209b3b3e84734620c3e071af828dbe038b55af779ace8508dbb7a6b3a`. The v1.0.0 tag and published assets remain unchanged. Detailed source/native/security scope is recorded in the [v1.0.1 evidence](https://github.com/PikkuJanne/WinBookSplit/tree/main/docs/codex-v1.0.1/evidence).