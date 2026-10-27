# Existing Store-installation release gate

## Current owner decision

OWNER_DECISION_20261027.md closes further owner work for v1.0.3. The fresh final-ZIP
owner Store and genuine company-policy scenarios below remain NOT_RUN/owner-waived
publication prerequisites. Preserve their original technical contract/checker and
old T07 acceptance; do not turn developer results into a Store PASS. Complete all
feasible agent/native developer checks and disclose unavailable manual scenarios.
Do not send the owner another test kit or request another run.

## Human work: one launch and one returned report, when needed

Prefer authorized agent access to the actual Windows 10 desktop. Otherwise
finish and self-test the kit, then give the owner ONE precise launch command for
that kit and the expected result location. No broad data recollection is needed;
the incident is already captured. Do not ask the owner to install another Python,
change aliases/PATH, select a runtime path, inspect every PDF, or take screenshots.
If a candidate fails, preserve its report and fix the evidenced failure; a rerun
is allowed but must be a consolidated verification, not another manual programme.

Use the maintained native kit with a PowerShell5.1/BAT controller; it must start
without needing an application Python to bootstrap itself. The exact launch
command is generated and executed by Codex before delivery. Do not publish a
speculative PowerShell command in the handover. Respect enforced organization
controls. No claim that automatic UI testing is physical Explorer drag-and-drop.

## Before application setup

- Use a NEW extraction of the exact candidate distribution ZIP in a short writable
  per-user directory. Do not overwrite the owner's checkout or remove its venv.
- Verify expected ZIP hash and package members from tested source R. Record host
  OS/build/architecture and actual inbox PS5.1 version. Keep terminal/Explorer
  environment differences visible. Do not modify PATH, LOCALAPPDATA or VIRTUAL_ENV
  for this scenario or activate a dev environment.
- Inventory current-user Store registration and its existing managed runtime(s)
  using an independent installed-only observation. This is evidence, not hidden
  input to application discovery. Confirm the eventual chosen interpreter existed
  before testing and record its version, executable hash and normalized identity.
- Fresh application `.venv` must be absent. The owner/global pypdf state may be
  unknown or already installed; do not alter it to manufacture a missing-package
  case. Readiness must correctly distinguish ready vs setup-required.

## Mandatory chain on the SAME candidate bytes

A. Run the application's native discovery/readiness route with no Python override.
   It must recognize the Store family and reach a concrete existing runtime. If
   pypdf is missing, record compatible-base/setup-required, not missing-Python.
B. Run normal `Setup-WinBookSplit.bat -NonInteractive`, no explicit runtime or
   controller-generated selection. Explicit package installation into the new app
   venv is allowed. An offline reviewed pypdf wheel may be provided as a package
   source when required; it is not permission to provide a Python interpreter.
C. Verify the venv base identity equals the automatically discovered existing
   runtime. Run setup again; confirm healthy no-op. Read-only `-Check` must not
   install or replace anything.
D. Run the ordinary WinBookSplit BAT with an authored PDF through scripted console
   input, from an unrelated working directory, without bypassing production
   selection. Also run manual/Auto1/Auto2 functional checks. Reopen outputs, compare
   source page identities/content and coverage, verify no empty output and no
   source/neighbor overwrite. Read actual engine receipts.
E. Record original installed-runtime hash and manager/PATH state afterward; confirm
   preserved state. Retain created application outputs/venv and the redacted
   results for the owner, not just a transient test environment that is discarded.
   Explain that a desktop shortcut still targeting the old v1.0.2 directory must
   target the verified v1.0.3 launcher; do not silently modify unrelated shortcuts.

The same launch then records a separate local policy component on that owned
extraction: authored ZoneId=3 streams on the processing/setup entry scripts,
actual direct inbox RemoteSigned denial, setup BAT read-only success and exact
automatic formatter command execution as a healthy no-op. It retains the marks,
package data bytes and venv, compares the same inbox policy scopes, and includes
all observations in the existing preservation envelope. It selects no runtime
and installs nothing. This is an authored local MOTW check; actual enforced
organization-policy denial remains `NOT_RUN` until independently demonstrated.
A policy refusal remains a meaningful failure and is never retried through
another launch/bypass route.

## Forbidden substitutions

No interpreter installation/download/copy, private runtime fallback, `-PythonPath`,
test-only interpreter environment setting, child PATH injection, private cloned
baseline, fabricated registration, or prebuilt application venv. A fixture manager
or explicit-runtime run is useful elsewhere but disqualifies THIS gate. Do not
silently downgrade an unavailable Store gate to a standard installation gate.

The existing native kit's `owner_regular315_usable=false` +
`reviewed_private3150_archive` path must fail the new gate, even when all PDF tests
pass. Audit both kit and validator. The controller must not pass the independently
inventoried runtime path as a hidden parameter to normal setup or processing.

## Report and review

Emit machine-readable plain UTF-8 JSON and a small bounded set of sanitized native
receipts; capture stdout/stderr as text rather than PowerShell CLI XML forwarding.
Preserve source/ZIP identifiers, run ID, registration tuple, listing/probe sequence,
base-to-venv relationship, exact argument vectors, exit codes, no-substitution
checks, PDF validation and before/after observations. A source hash is not the
package hash. Hash the RETURNED redacted bytes, and keep local raw hashes separate.
Use consistent path tokens so correlation survives redaction. No credentials or
private book text in logs/fixtures/CI/artifacts.

Adapt the maintained report validator and map its output to
`templates/native-store-report.json`; the supplied checker is an additional
no-substitution gate, not a replacement for existing PDF/native checks. Every
receipt file is relative to the returned report directory, bounded, ordinary,
non-symlink, and hash-checked. Require a complete artifact set for registration,
discovery, setup, BAT, PDF validation and preservation. Review actual receipts,
not only a top-level status field. Old reports stay old-schema evidence.

## Blocking

If the real desktop is unavailable, finish code, CI, package and kit, mark this
gate BLOCKED and give the one launch instruction. Do not publish, tag, call the
application fixed, or replace the owner install merely to clear the gate. Native
results on an equivalent pre-existing Store install are useful, but this incident
is not closed until the owner's installation succeeds with the verified package.
