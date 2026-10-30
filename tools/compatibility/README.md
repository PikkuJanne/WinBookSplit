# Finite native package acceptance kit

For the active [functional repair](../../docs/functional-repair/STATUS_NEXT.md),
preserve the single pretested `WinBookSplit-Functional-Native-Checkpoint4.zip` in
Downloads. The latest direct owner order removes all human testing, owner sign-off
and return-report requests. No owner action is required. Reuse is limited to an
agent-accessible qualifying machine; unavailable native cases stay NOT_RUN and
do not become a human-testing blocker. Its `START-HERE.bat` and numbered guide
return one redacted
`SEND-THIS-REPORT.zip`, including failures/missing prerequisites. It binds the
checkpoint-4 source/archive and verifies ordinary BAT consent/setup/resume plus
repeat, without provisioning Python, supplying a selected runtime, injecting PATH
or prebuilding the ordinary app environment. Network activity stays NOT_RUN
without a trace. A new documentation-only archive does not change that kit's
binding or establish new Store acceptance. Checkpoint 6 performs no publication;
eventual publication uses the later direct authority and a future source/package
freeze.

The remainder describes the historical v1.0.3 kit and its separate scenarios,
not a requirement to restart T07 or an authorization to acquire a private runtime
for the functional repair. Its human-testing requirements are superseded, not
pending tasks. Historical results/specs keep their original scope.

This developer tool builds a self-contained kit for one exact distribution ZIP
on regular Windows 10 22H2 x64. It carries the frozen release assets, authored
fixtures and authenticated helpers. GitHub's automatic source archive does not
substitute for the distribution. Owner startup uses inbox PowerShell without
application Python or Git. The v1.0.3 T06 evidence records the frozen candidate
and exact delivered launch-command checks; T05 source/component checks alone
do not certify those archive bytes. T07 held the original Windows 10 Store gate.

The builder defaults to `--scenario store-existing`. Its generated
`Run-Windows10.cmd` explicitly selects that scenario without provisioning. After
authenticating the kit, this route refuses supplied runtime paths, activated or
injected Python configuration and prebuilt application environments. Independent
Store registration/listing/runtime inventory is verification and helper-bootstrap
evidence only. Fresh application discovery and normal setup/BAT receive no
selected interpreter or environment overrides. An unavailable existing Store
runtime blocks before Python acquisition, copying or cloning. Ordinary setup may
create the new app-local venv and install the locked pypdf after discovering the
existing interpreter. A reviewed wheelhouse is a package source only.

The Store report and its six receipt envelopes use schema 2. The offline
`tools/codex-v1.0.3/check_native_store_report.py` additionally rejects attempted
fallback even when unused, supplied/copied/cloned/injected runtimes, inventory
used as application selection, and old schemas. These consistency checks do not
attest native execution. The original contract required actual Windows 10 owner
acceptance; the latest repair order removes that human-testing requirement.
Store-existing returns the printed `native-store-result.zip`, including sanitized
completed/failure receipts. Its raw working extraction, venv and outputs remain
local. Retain a blocked return for diagnosis; never edit it into a passing report.

After the automatic Store setup and PDF chain, the same launch runs a separate
local policy component against that owned extraction and its healthy venv. It
authors ZoneId=3 alternate streams on the two entry scripts while preserving
their package data bytes, proves direct inbox `RemoteSigned -File ... -Version`
denial, runs the actual setup BAT read-only check, and executes the exact shared
formatter's automatic setup command as a healthy no-op. It supplies no Python
path, clones no environment and performs no package operations in this component.
The same inbox policy scopes, script bytes, retained marks and venv are compared
before/after. These command/observation records enter the existing preservation
receipt and its hash-bound `policy_component` summary. A refusal is retained and
never retried through another policy route. These authored MOTW checks leave
real enforced organization-policy denial explicitly `NOT_RUN`; they do not clear
that separate T07 publication blocker.

The separate `--scenario legacy-explicit` retains the established explicit and
private-runtime regression recipes below. Its schema 1 reports and private
baseline can never satisfy the v1.0.3 Store-existing gate. Historical reports and
published release bytes remain unchanged.

The generated launcher applies Bypass only to its inbox child process where
organization policy permits. Enforced controls remain meaningful failures. In
the **legacy-explicit scenario only**, a usable regular x64 Python 3.15 may be
selected or the reviewed private 3.15.0 archive provisioned. Its missing-package
baseline may use an owned `--without-pip` fixture. These are separate regression
inputs, never Store-existing acceptance.

The legacy suite returns the printed `native-result.zip`; raw local diagnostics,
fixtures and outputs remain local. A prerequisite or process failure produces
FAIL/BLOCKED/NOT_RUN evidence and never counts as a pass. Enforced organization
signing rules require an approved route and are never retried through encoded code.

The legacy finite cases cover Version/help, installation, Check/no-op, actual setup BAT on Internet-marked scripts, exact emitted PowerShell Action execution, preserved broken-environment Repair, RemoteSigned direct-script denial, observed child/parent policy scopes, manual and both Auto modes, a later first start, repeat publication, Preview, invalid input, missing bookmarks, cancellation, missing Python/pypdf, genuine EPUB/AZW3 keep/discard, and the byte-identical processing BAT's one/two-source forwarding. The maintained remedy regression uses authenticated copies of the exact extracted ZIP members and original authored PDFs. Its noninteractive processing probe preserves the dependency failure without offering setup. A separate actual `Setup-WinBookSplit.ps1 -Check -NonInteractive` emits the structured Action; the harness requires no changes or setup operations, validates every literal command element, executes that exact text, and retries the unchanged original split request with interactive plan confirmation. Public Check has no document argument; the processing probe retains the original input path and the expected Action base is required to be outside its directory. These are explicit developer fixtures, not ordinary discovery or consent/resume acceptance. No owner Python or PowerShell installation is requested. Optional reviewed PS7 is supplied only for agent workstation coverage. Timeout, cancellation, failed-legacy-launcher and mocked organization-policy controls retain explicit controlled/simulation labels. A mocked organization refusal
is a guard control, not a claim that real organization enforcement was tested. Physical pages, exact source/fixture hashes, selected interpreter/package origins, console ownership and complete process stop/drain are validated before PASS.

The remedy harness reads its expected pypdf version through the authenticated
application's strict compatibility policy; historical dependency pins are not
substituted for the selected package source.

The older `kit_builder.py` acquisition recipes retain their historical pypdf
6.21.3 and Calibre 9.15.0 pins, as do their already built kits. Those recipes are
outside current release acceptance and are not used to provision its test tools.
Current checks use the newly frozen application and separately authenticated
6.21.4/9.17.0 tools; no archived kit is rebuilt or relabelled.

BAT cases retain its actual default Documents directory and authenticate only emitted owned paths. The Explorer path is unchanged; no repeated manual sign-off is requested. The harness retains all raw work and performs no recursive acceptance-work cleanup.

## Developer build

Prepare original fixtures once with the existing fixture generators and a developer environment. ReportLab/Pillow are producer dependencies and are excluded from the owner kit. Keep the bounded raw preparation receipt outside the six-file fixture/catalog closure; only that frozen closure enters the kit, including genuine Calibre-generated AZW3.

```powershell
& $DeveloperPython -I -B tools/compatibility/prepare_fixtures.py `
  --repository-root $Repository --output-root $FreshFixtures `
  --calibre-path $Calibre --report $FreshPreparationReport

& $BarePython -I -B tools/compatibility/kit_builder.py build `
  --scenario store-existing `
  --repository-root $Repository --assets-root $FrozenAssets `
  --fixture-root $FreshFixtures `
  --core-path "$Repository/tools/compatibility/Invoke-NativeAcceptance.ps1" `
  --oracle-path "$Repository/tools/compatibility/native_smoke.py" `
  --verifier-path "$Repository/tools/release/frozen_assets.py" `
  --output $FreshKit
```

Use absolute ordinary paths and a new output name. The builder emits the kit directory, sibling ZIP and build receipt through exclusive writes. Its default COMMITTED mode requires every helper's raw bytes to match the recorded Git commit; Git queries reuse the existing bounded, sanitized package-input helper. `--draft-rehearsal` instead records null kit source identities and a separate repository observation. Such a rehearsal never certifies different public or final release bytes. Keep the independent kit ZIP hash with the delivery.

Legacy preflight may supply explicit ordinary `-BarePythonPath`,
`-BaselinePythonPath` (regular 3.15.0 without pypdf), reviewed `-CalibrePath`, exact
`-Wheelhouse` and optional reviewed `-PowerShell7`. Store-existing refuses runtime
parameters and `-ProvisionPrerequisites`; it may receive only the wheelhouse as
a package source. `-AllowWindows11DryRun` labels Windows 11 separately and supplies
no Windows 10 evidence; `-Execution AGENT_LAUNCHED` is explicit for agent runs.

## Developer controls

`tests/python/test_native_kit.py` exercises strict JSON, authentication, guarded exclusive output, preserved collisions, sanitized Git environment, held directory identity and refusal to adopt an unlinked/recreated archive. Its temporary work is limited to authored fixtures. Run it through the existing Python test runner.

`tests/compatibility/Probe-NativeKit.ps1` accepts explicit RepositoryRoot, CoreRoot, TestPython, a fresh WorkRoot and ReportPath. Its thirteen controls exercise bounded reads, strict receipts, absent/empty environment restoration, malformed proof refusal, fixed OS redaction, the vendor's UTF-16 path bound and actual owned descendant timeout/cancellation under the selected native shell. These controls are tooling checks, not final-ZIP or Windows 10 acceptance. The normal scoped shell syntax and analyzer checks remain required.

`Invoke-NativeDiscoveryCollector.ps1` reuses the production job supervisor to
capture both UTF-8 streams as text into bounded, exclusive plain JSON receipts.
The collector never forwards nested PS5.1 stderr through a PowerShell pipeline.
Its no-Python probe covers positive Unicode streams, nonzero failure, launch
failure, timeout, inherited hung streams/descendants, truncation and unchanged
environment. The maintained Python suite runs this probe under actual PS5.1/PS7
controller/child combinations; unavailable hosts are reported as skips.

## Supplemental PowerShell 7 checks

After complete Windows 11 inbox 5.1 acceptance of the exact candidate ZIP,
run `Invoke-FinalZipPs7Acceptance` against that same application extraction and
its authenticated prior local report. Use the reviewed kit/bindings hash, prior
report hash, adapter hash and an explicit ordinary PowerShell 7.6.6 executable.
Do not derive expected pins from an untrusted manifest or substitute another ZIP.

The calling session must authenticate and hold the SDK source files before
loading Logging, Diagnostics, Process, Support and Invoke-NativeAcceptance.
Call the existing `Initialize-WinBookSplitLogTypes`,
`Initialize-WinBookSplitSupport` and `Initialize-WinBookSplitProcess` functions
in that order. The invocation below assumes this guarded preparation has
completed in the current session.

```powershell
. "$Repository/tools/compatibility/Invoke-FinalZipPs7.ps1"
Invoke-FinalZipPs7Acceptance -KitRoot $KitRoot -BindingsSha256 $ReviewedBindingsSha256 `
  -PriorReport $PriorReport -PriorReportSha256 $ReviewedPriorReportSha256 `
  -NativePlan $NativePlan -ExtractionReceipt $ExtractionReceipt `
  -PowerShell7 $PowerShell7 -ExpectedPowerShell7Version '7.6.6' `
  -AdapterPython "$Repository/tools/compatibility/final_zip_ps7.py" `
  -AdapterPythonSha256 $ReviewedAdapterPythonSha256 -WorkRoot $FreshPs7Work
```

The four direct `-File` cases are Version, healthy read-only Setup Check,
Manual `1,4,7` on the authored ten-page PDF, and Auto2 on the nested twelve-page
PDF. The adapter verifies the real PS7 path/version and native workstation facts,
reuses the existing owned Process/Session transport and console/page oracles,
and requires the application environment, source fixtures and prior output
marker to remain unchanged. It installs nothing and retains all raw work.
The completed `ps7-result.json` uses fixed redacted fields; raw process reports
remain private. These supplemental checks provide no Windows 10 evidence and do
not repeat the complete inbox 5.1, ebook, BAT or Explorer acceptance programme.
Syntax, scoped analyzer and synthetic binding controls are preparation checks;
only an actual run on the authenticated final ZIP can certify these four cases.
