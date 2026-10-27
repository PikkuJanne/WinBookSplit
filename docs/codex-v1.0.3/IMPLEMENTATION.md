# Focused implementation specification

Current 27 October follow-up selects exact official pypdf6.21.3/policy5 for the
reviewed xref-stream budget and parser resource fixes. Shared engine/discovery/
setup/origin/process/file implementations and Python/Calibre floors remain intact.
OWNER_DECISION_20261027.md closes further owner work; unavailable manual final-ZIP
Store/company-policy checks remain NOT_RUN/owner-waived. Old accepted T07 applies
only to its original bytes. Complete all feasible agent checks and truthful public
release verification; do not request another owner run or change validators.

This is a repair specification, not a prevalidated Windows patch. Reuse the
existing functions and supervisor. Inspect all call sites and package allowlists
before editing. Avoid adding a second independent runtime resolver or test-only
production path. Primary source links are in SOURCES.md.

Current acceptance: T07's unchanged actual owner Windows10 Store return passed
for frozen R/exact ZIP; see evidence/T07-owner-return.md and JSON. Separate actual
enforced-policy STORE-024 remains NOT_RUN and T08 publication stays BLOCKED.
The T02/T03 checkpoint observations below retain their original component scope.

The owner explicitly authorized a separate pypdf repair on 18 October. The current
implementation changes the exact dependency to 6.21.1 with wheel SHA256
553e2d42f5bb3bb2a785a7d3de8ef90bb7df74b5413eff5a7b4388007ea9254b, policy4,
necessary maintained pins/assertions and three small parser guard regressions.
Runtime/Calibre policy and engine/guards remain intact. This creates new candidate
bytes: original R/T07 remain preserved and cannot satisfy the changed candidate's
native acceptance. See tasks/08-T08-security-repair.md and current evidence.

T02 checkpoint: the exact-name current-user query and paired identities are
implemented. `Get-WinBookSplitRegisteredManagerPaths -WithRegistration` retains
validated package data for `Invoke-WinBookSplitManagerListing`. The documented
family alias must have ordinary trusted ancestors and the AppExecutionAlias tag;
Windows resolves it into a suspended child whose package full name and physical
image must match registration before resume. The existing job, deadlines,
cancellation and stream handling remain in force. Real python.org component
listing passed on Windows 11/PS5.1/PS7; Store-channel native evidence is NOT_RUN.
T03 wires these descriptors and this adapter into the existing automatic resolver.
Manager path-only listings and legacy `-0p` listings remain separate. Concrete
runtime candidates still pass unchanged ordinary/automatic path guards and the
independent isolated identity/policy/package probe; no hardlink exception was
needed for the observed concrete Windows 11 runtime. Query completion/count and
empty/rejected/failed outcomes are retained in attempts. See T03 evidence for
real below-floor rejection versus supported synthetic-listing/actual-child tests.
Store-channel/owner Windows 10 exact-ZIP acceptance remains NOT_RUN at T07.

## 1. Current-user manager registration

Read `Get-WinBookSplitRegisteredManagerPaths` and
`ConvertFrom-WinBookSplitManagerRegistration` in `engine/WinBookSplit.Runtime.ps1`.
Query the exact package NAME without excluding Store packages by the old publisher
filter. Keep current-user scope, bounded count/output/time, inbox PS5.1 Appx
module and complete stream capture. Do not query all users or elevate.

Validate complete identity tuples for both official distribution channels:

| Channel | Family suffix | Publisher | Signature kind |
| --- | --- | --- | --- |
| Microsoft Store (observed) | qbz5n2kfra8p0 | CN=4975D53F-AA7E-49A5-8B49-EA4FDC1BB66B | Store |
| python.org (existing code) | 3847v3x7pw1km | CN=Python Software Foundation, O=Python Software Foundation, L=Beaverton, S=Oregon, C=US | Developer |

Both use name `PythonSoftwareFoundation.PythonManager`. Reconfirm these pairings
against current registered packages/official metadata before implementation.
Do not accept cross-paired publisher/family/signature data, merely matching names,
development/staged/resource/framework packages, unregistered decoys, or arbitrary
redirects. Retain schema checks, architecture/status checks and registration
consistency. Installed versions vary: do not freeze 26.4.240.0. Robust version
parsing must not require the old owner version to pass.

The original report did not capture all staging/development flags: synthetic
fixtures provide explicit values labelled SYNTHETIC. Never label them observed.
Registration identity is input to manager resolution, not proof every returned
runtime/package is safe.

## 2. Supported entry point; no global weakening of path checks

Python documentation gives package-family-specific entry points under
`%LOCALAPPDATA%\Microsoft\WindowsApps\<package-family>` and discourages launching
the physical executable inside Program Files. Prefer the appropriate documented
`pymanager.exe` route; validate the actual current-user registration and bind the
chosen alias/activation to that package. Do not rely on arbitrary PATH names.

A generic `Get-WinBookSplitTrustedPath` relaxation is NOT the solution. Implement
a narrow manager-entry validation/launch adapter if necessary. Verify AppExecutionAlias
binding with supported Windows/package mechanisms and real tests. If OS activation
is used, account for the actual process/descendants and argument handling; do not
escape bounded supervision by marking the alias process as irrelevant.

Probe the real route early: packaged aliases may behave differently from ordinary
executables under suspended CreateProcess/Job Object code. Do not assume the
existing physical-executable supervisor is already compatible. Reuse it where
possible; make only the smallest tested adaptation where necessary. Preserve
finite deadlines, cancellation, descendants/stream completion and meaningful exit
codes. No unbounded shell/GUI activation, no launch through an unverified decoy.

Hard links have no unique original filename. Do not guess the concrete runtime
with GetFinalPathNameByHandle alone or simply run any hard-linked PATH entry. A
real runtime can itself have multiple links; determine whether current checks
wrongly reject it and, if necessary, use a bounded registered-runtime provenance
rule rather than a blanket link-count ban/permission. Retain book/CWD decoy guards.

## 3. Installed-only enumeration, then independent runtime verification

Use an explicit documented manager command such as `list --format=exe` through
the verified entry point, after checking actual supported syntax. Keep the legacy
launcher parser separate; `-0p` remains an option for genuine legacy routes.
Never use `exec`, `install`, `--online`, `--source`, configuration refresh, or a
bare launcher call to discover a runtime. Disable automatic install in the child
as defense in depth, without changing parent environment or user configuration.
List operations may produce bounded local manager logs; do not describe them as
absolutely zero filesystem writes. No runtime/package installation or network
index retrieval is allowed in discovery.

Parse bounded complete output. Preserve Unicode/spaces and validate absolute paths;
reject malformed/truncated/duplicate-key JSON or invalid line formats explicitly.
Do not parse friendly table output or localized stderr as paths. A successful
query with zero packages is NOT a manager acceptance; listing zero runtimes is a
separate outcome. A failed optional route must not discard a valid independent
candidate.

Pass concrete installed interpreter candidates into the existing compatibility
probe. Verify the independently returned `sys.executable`, x64 regular CPython,
isolation flags and policy identity. Keep pypdf lock/origin checks for processing.
Do not insist that `sys.executable` equal the PATH forwarding alias: compare it
to the correctly resolved concrete candidate. Keep deterministic branch/patch
ranking and explicit-path/existing-venv authority.

Setup's interpreter-only discovery must not require pypdf beforehand. Successful
base discovery followed by missing pypdf means setup is needed, not absent Python.
Processing may use its healthy app `.venv` afterward. Do not pass a controller's
base selection back as hidden input to setup; setup must discover it itself.

## 4. Diagnostics and setup UX

Update `Get-WinBookSplitDependencySummary` and shared helpers, not only console
strings. Give discovery rejections categorized Reason/stage/source/probe records.
Distinguish no candidates, recognized manager without runtimes, rejected entry
point, rejected manager identity, listing failure, interpreter failure, unsupported
version/flavor, missing pypdf, and invalid app venv. Prefer the deepest relevant
failure without hiding failed alternatives. Sanitize displayed text; retain
technical details locally. Mark a query completed separately from a manager or
interpreter accepted. Keep outcome exit semantics compatible.

Use existing `Get-WinBookSplitSetupCommand`/setup BAT for exact executable remedies.
Do not recommend reinstalling Python when a compatible base was found. Run every
emitted remediation command tested here verbatim (substituting only sandbox-owned
paths where fixtures require), rather than reconstructing a working variant.
Check no-Python/readiness guidance without accidental installs. Preserve
RemoteSigned/MOTW behavior, process-only policy, native exit code and no global
changes. Enforced policy denial remains a meaningful failure.

## 5. Remove acceptance's private-runtime escape route

Inspect `tools/compatibility/Invoke-NativeAcceptance.ps1`,
`tools/compatibility/native_smoke.py`, report validators and acquisition helpers.
Add an explicit Store-existing-installation scenario to the maintained kit and
version its report additions deliberately. Historical reports retain their old
contract; they cannot satisfy the new gate.

The new scenario must fail BEFORE any fallback provisioning if no pre-existing
Store runtime is discovered. No controller-injected PythonPath/PATH/LOCALAPPDATA,
prebuilt app venv, copied runtime or existing-runtime-to-private-baseline cloning
is permitted. Explicit-runtime and private-runtime tests may still run separately
for other coverage, labeled as such. Check the gate validator rejects those reports.

Inventory the existing registered runtime independently for verification only;
it must not become application configuration. Correlate native registration,
listing, probe, setup and engine receipts. Hash the chosen runtime before/after;
verify the created app venv's base is that same installed interpreter. Do not use
the broken PS5.1 nested collector pattern; capture stdout/stderr safely and write
plain UTF-8 JSON files with an explicit schema and exit status.

## Likely edit surface (inspect actual checkout)

`engine/WinBookSplit.Runtime.ps1`, `WinBookSplit.ps1`, shared diagnostic/setup
helpers as needed; existing runtime/Pester tests; native compatibility kit and
its validator; docs/SETUP.md, README, CHANGELOG, package allowlist and VERSION.
Avoid changes to the PDF engine, document policy or dependency versions unless a
reproduced regression necessitates a separate scoped change.
