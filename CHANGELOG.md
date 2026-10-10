# Changelog

## 1.0.0 — release candidate

This version is a candidate under development, not a published release. The
canonical application version is stored in `VERSION`. Release-package checks,
publication and anonymous asset verification have their own later gates.

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
