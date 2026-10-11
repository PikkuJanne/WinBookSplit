# Changelog

## 1.0.1

- Adds the Windows 10 22H2 x64 compatibility target alongside Windows 11, while
  keeping the BAT/PowerShell console and existing splitter/converter workflow.
- Replaces exact Python-patch equality with a shared reviewed policy for regular
  GIL Windows x64 CPython 3.13.16+ and 3.14.8+, within those two minor branches.
  The exact source-test targets are 3.13.16 and 3.14.8; later stable patches are
  accepted without claiming each one was tested. The security floors include
  the bundled OpenSSL fix used by explicit HTTPS setup.
- Keeps pypdf 6.20.0 and its wheel hash controlled separately, with package origin,
  isolated imports and existing document/process safety guards retained.
- Improves bounded read-only discovery, preserves a valid PATH fallback after
  failed launcher listing, and selects accepted automatic candidates consistently.
- Adds explicit application-local setup, read-only readiness checks, reviewed
  offline wheel use and repair with preserved backups and safe rollback.
- Gives ordinary dependency failures a concise reason and setup action while
  preserving local technical diagnostics and native nonzero outcomes.
- Expands automated source regression and CI across two actual Python minors.
  [Release evidence](https://github.com/PikkuJanne/WinBookSplit/tree/main/docs/codex-v1.0.1/evidence)
  identifies the native final-package scope separately from source/Server checks.

## 1.0.0 — 2026-10-11

Published as [WinBookSplit v1.0.0](https://github.com/PikkuJanne/WinBookSplit/releases/tag/v1.0.0).
The canonical application version is stored in `VERSION`. Its historical release
evidence, tag and published assets remain unchanged.

- Retains the Windows BAT/PowerShell console and single-file drag-and-drop flow,
  with explicit plan confirmation, cancellation and optional opening requests.
- Uses one validated physical-page plan for manual starts and Level 1/Level 2
  bookmarks, including front matter and parent opening pages.
- Converts EPUB and genuine unencrypted AZW3 to PDF through separately installed
  Calibre; optionally retains the full converted PDF.
- Preserves immutable source files and publishes unique output directories after
  checking chapter files. Output cleanup is restricted to this run's owned files.
- Adds complete scripted parameters, no-prompt execution, plan-only preview,
  bounded process supervision and meaningful native exit statuses.
- Adds bounded local console/run diagnostics and an explicit local redacted
  support export, without automatic uploads or telemetry.
- Documents tested page content, limited metadata/navigation/static annotations,
  conservative unsupported-PDF rejection and the exact dependency pins.
- Adds authored regression fixtures, strict source-bound receipts and scoped
  Windows CI. Those checks are distinct from release-package certification.
- Resolves the BAT shell from the Windows system location and fails clearly
  when it is missing; CWD/PATH shell decoys cannot replace it. Current security
  pins use pypdf 6.20.0 and PowerShell 7.6.6.

## Historical development naming

The earlier `2.1` banner predates the public semantic-version plan. It identifies
legacy development code, not a previous public 2.1 release. No earlier public
release or tag is invented by this changelog.

See [README](README.md) for the current support scope and
[third-party notices](THIRD_PARTY_NOTICES.md) for separately installed components.
