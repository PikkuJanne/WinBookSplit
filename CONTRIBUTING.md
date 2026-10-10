# Contributing

Open a GitHub issue for a non-sensitive bug report or a focused improvement, then
keep the change within this small Windows PDF/ebook-to-PDF splitter. Include the
application version, tested Windows/PowerShell versions, outcome code and steps
using a generated or redistributable fixture. Do not upload private books,
unredacted logs, source paths, credentials or workplace documents. Use the
[redacted support export](docs/SUPPORT_DIAGNOSTICS.md) when appropriate and review
the summary before sharing it. Follow [SECURITY.md](SECURITY.md) for security issues.

Preserve the BAT/PowerShell workflow, Python/pypdf engine and optional Calibre.
All modes must share an ordered, nonempty partition covering each physical page
exactly once. Keep source inputs immutable, never overwrite prior output, and
clean up only files owned by the current run. Preserve both process streams and
nonzero statuses. Paths and document text must not become shell commands.
Reproduce a bug before fixing it; keep mechanical refactors separate from behavior.

Use the existing tests and pinned isolated developer setup described in
[tests/README.md](tests/README.md). Run checks affected by the change and retain
actual commands, versions, source identity, outcomes and skipped scopes.
Unit/syntax results are not end-to-end Windows, Calibre, GUI or package results.
Generated fixtures belong in the test recipe, not private-document check-ins.

Update user guidance when behavior changes. Avoid broader compatibility claims,
new frameworks, dependency auto-installation, uploads/telemetry, OCR inference,
DRM removal, installers or a website. Dependency updates need primary-source
review and affected regression evidence. Keep license/attribution notices intact.
