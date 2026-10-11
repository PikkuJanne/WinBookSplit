# Verified starting observations

Observed 11 October 2026 through the GitHub connector. Refresh before implementation; this file is neither an instruction to reset nor a claim about a later checkout.

- Current observed main: `be1fa1e271aa2bd031d907e9dbc9e7bfc32bbb62`, message `Icon and poster`. Preserve these newer assets. [R1]
- Published v1.0.0: source `73c90d10dc16b1c321737afbdfbe959d61a24573`, publication 2026-10-11T08:02:26Z. Its ZIP SHA-256 is in BASELINE.json. No v1.0.1 was in the observed release list. Old release bytes/tag must not be modified. [R2]
- Root AGENTS still routes sessions to the completed v1.0.0 plan and prohibits new versions without authority. This user request supplies authority for v1.0.1; update its routing and narrow scope deliberately. [R3]
- `engine/WinBookSplit.Runtime.ps1` checks CPython 3.14.8 exactly inside its embedded Python probe **and** again in PowerShell's record validation. It separately checks pypdf 6.20.0 and Calibre 9.15.0 exactly. [R4]
- Runtime order is explicit path, authoritative application .venv, `py.exe -0p`, then python.exe on PATH. Failed launcher listing already falls through; the user's final rejection was the exact-version check. Do not describe that fallback as absent. [R4]
- The reported Python 3.14.2 failed before pypdf import. No evidence here establishes whether that machine has pypdf. The launcher error alone does not establish whether py.exe is the legacy launcher, manager, shim or a misconfigured executable. [User log; R4]
- The README declares Windows 11 only and exact runtimes. Its candidate label was frozen before publication; the actual release API is the authority for publication status. [R2, R5]
- PDF policy uses a `PdfReader` subclass/private `_handle_encryption` hook and `root_object_recovery_limit`. Do not broaden pypdf blindly. [R6]
- Existing CI uses `windows-2025`, a single exact Python and current hash-pinned developer packages. `tests/run_tests.py` already supplies extensive synthetic/evidence machinery. Extend it instead of replacing it. Developer wheels and validators need separate matrix review. [R7, R8]
- Existing packaging lives under `tools/release/`, with `build_package.py`, `payload.json` and README. Inspect all embedded pins, manifests and evidence digests before extending the allowlist. [R9]

All source identifiers resolve in SOURCES.md. This handoff contains no production patch, private user document, credentials or fabricated runtime evidence.
