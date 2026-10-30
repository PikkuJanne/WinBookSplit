# Changelog

## Unreleased functional repair

- Interactive processing offers existing application-local setup or explicit
  backup-preserving repair once, proceeds only after consent, then rechecks the
  app environment and resumes the original request. Preview, noninteractive
  processing and readiness checks neither prompt nor install.
- Accepts PDFs through ordinary empty-password authentication; distinguishes
  nonempty-password requirements, unsupported encryption, missing crypto backend
  and malformed input. No password-entry or DRM-removal feature is added.
- Ignores bounded validated local initial views with one categorized warning;
  chapter PDFs are unencrypted without the opening action. Active/malformed
  actions, source/output/process safety and dependency locks/floors remain guarded.
- These changes belong to the repair branch and unpublished test packages.
  Published v1.0.3 and earlier tags/assets are unchanged. Ordinary Store first-run,
  repeat/network, Windows 10 and private-book acceptance remain NOT_RUN; no new
  release or publication is authorized. See the
  [repair status](https://github.com/PikkuJanne/WinBookSplit/blob/codex/functional-repair/docs/functional-repair/STATUS_NEXT.md).

## 1.0.3

- Repairs automatic discovery of official Microsoft Store Python Install Manager
  registrations, validates each official channel's complete identity tuple, and
  resolves its verified package-bound entry point to existing concrete runtimes.
- Keeps regular x64 Python 3.13.16+, 3.14.8+ and 3.15.0+ support, reviewed security
  floors, isolated package-origin checks and authoritative explicit/app-venv setup.
- Updates the exact pypdf lock to 6.21.3 and its reviewed wheel SHA256 for
  upstream parser/resource fixes, including the document-wide cross-reference
  stream budget, retaining existing package-origin, filter and process guards.
- Reports rejected manager, listing, interpreter and package routes distinctly,
  retaining bounded local technical records and complete setup/repair actions.
- Adds an existing-Store native verification route that blocks before fallback
  acquisition and rejects runtime substitution or contradictory process evidence.
  Retains frozen-distribution bindings and a separate checker for native receipts.
- Preserves the accepted earlier frozen Windows 10 existing-Store result for
  its original bytes. Under the owner's 27 October decision, a fresh Windows10
  Store run of the final ZIP and real company-policy testing remain untested;
  the release does not claim those passes. See the
  [v1.0.3 evidence](https://github.com/PikkuJanne/WinBookSplit/tree/codex/v1.0.3-store-discovery/docs/codex-v1.0.3/evidence).

## 1.0.2

- Adds tested regular GIL Windows x64 CPython 3.15.0 and numeric acceptance of
  later stable patches in 3.15, while retaining serviced 3.13.16+ and 3.14.8+
  branches and their reviewed security floors. Unknown minors remain unvalidated,
  without being labeled insecure. pypdf 6.20.0 and its wheel hash remain unchanged.
- Adds `Setup-WinBookSplit.bat`, using the OS-known inbox PowerShell host and a
  process-scoped invocation for Internet-marked unsigned setup scripts where
  organization controls permit. Persistent policies and Internet marks remain
  unchanged; enforced signing and application controls still apply.
- Generates complete, literally quoted PowerShell setup and repair actions and
  reports the authoritative selected runtime, observed version and specific
  reason. A rejected or broken environment cannot silently fall back to another
  interpreter, and processing never installs dependencies.
- Preserves application-local setup, read-only Check, reviewed offline wheel use,
  explicit Repair backups, safe output ownership and bounded process/stream checks.
  Setup output is UTF-8, including direct CMD invocation at Unicode paths.
- Extends the maintained CI matrix and finite native acceptance kit for Python
  3.15, marked setup BAT/remedy round-trips and observed host policy scopes.
  [Release evidence](https://github.com/PikkuJanne/WinBookSplit/tree/main/docs/codex-v1.0.2/evidence)
  distinguishes actual source/Server checks from exact-package Windows 10/11
  acceptance and records any unavailable gates. A source ZIP or candidate alone
  is not the accepted release distribution.

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
