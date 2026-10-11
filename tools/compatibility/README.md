# Finite native package acceptance kit

This developer tool builds a self-contained kit for testing one exact release ZIP on regular Windows 10 22H2 x64. The kit carries the three frozen release assets, original authored fixtures, and a small authenticated helper closure. It verifies the external asset hashes and committed kit helper bytes before extracting or running application code. Owner verification and extraction require no Git checkout.

Run the delivered `Run-Windows10.cmd` once. This explicitly downloads hash-pinned ordinary CPython 3.14.8, pypdf 6.20.0 and portable Calibre 9.15 into fresh private work, then runs the shipped setup and acceptance cases. It creates no global packages, PATH changes or elevation requirement. Return only the printed `native-result.zip`; raw local diagnostics, source fixtures and outputs remain local. The returned ZIP contains one allowlisted JSON report, excluding personal paths and raw console text. A prerequisite or process failure produces FAIL/BLOCKED/NOT_RUN evidence; it never counts as a pass.

The finite cases cover Version/help, installation, read-only Check, healthy setup no-op, manual and both Auto modes, a later first start, repeat publication, Preview, invalid input, missing bookmarks, cancellation, missing Python/pypdf, genuine EPUB/AZW3 keep/discard, and the byte-identical BAT's one/two-source forwarding. Authored timeout, cancellation and failed-legacy-launcher controls use the existing owned process transport. Their CONTROLLED_NATIVE label does not claim an installed classic launcher. Physical pages, exact source/fixture hashes, selected interpreter and package origins, console ownership, and process stop/drain proof are validated before PASS.

BAT cases retain its actual default Documents directory and authenticate only emitted owned paths. The Explorer path is unchanged; no repeated manual sign-off is requested. The harness retains all raw work and performs no recursive acceptance-work cleanup.

## Developer build

Prepare original fixtures once with the existing fixture generators and a developer environment. ReportLab/Pillow are producer dependencies and are excluded from the owner kit. Keep the bounded raw preparation receipt outside the six-file fixture/catalog closure; only that frozen closure enters the kit, including genuine Calibre-generated AZW3.

```powershell
& $DeveloperPython -I -B tools/compatibility/prepare_fixtures.py `
  --repository-root $Repository --output-root $FreshFixtures `
  --calibre-path $Calibre --report $FreshPreparationReport

& $BarePython -I -B tools/compatibility/kit_builder.py build `
  --repository-root $Repository --assets-root $FrozenAssets `
  --fixture-root $FreshFixtures `
  --core-path "$Repository/tools/compatibility/Invoke-NativeAcceptance.ps1" `
  --oracle-path "$Repository/tools/compatibility/native_smoke.py" `
  --verifier-path "$Repository/tools/release/frozen_assets.py" `
  --output $FreshKit
```

Use absolute ordinary paths and a new output name. The builder emits the kit directory, sibling ZIP and build receipt through exclusive writes. Its default COMMITTED mode requires every helper's raw bytes to match the recorded Git commit; Git queries reuse the existing bounded, sanitized package-input helper. `--draft-rehearsal` instead records null kit source identities and a separate repository observation. Such a rehearsal never certifies different public or final release bytes. Keep the independent kit ZIP hash with the delivery.

On a supported machine, the generated PowerShell wrapper can also accept explicit ordinary `-BarePythonPath`, reviewed `-CalibrePath` and exact `-Wheelhouse`. This avoids downloads while retaining all verification and setup stages. Private Calibre acquisition uses a fresh short directory under the OS-known user profile, with its final portable path checked against the vendor's 58 UTF-16-unit limit before launch. A profile that cannot fit this bound records BLOCKED; it never opens an automatic folder chooser. `-AllowWindows11DryRun` labels the actual Windows 11 scope separately and provides no Windows 10 evidence. `-Execution AGENT_LAUNCHED` must be stated for an agent-launched run; the owner command defaults to USER_LAUNCHED.

## Developer controls

`tests/python/test_native_kit.py` exercises strict JSON, authentication, guarded exclusive output, preserved collisions, sanitized Git environment, held directory identity and refusal to adopt an unlinked/recreated archive. Its temporary work is limited to authored fixtures. Run it through the existing Python test runner.

`tests/compatibility/Probe-NativeKit.ps1` accepts explicit RepositoryRoot, CoreRoot, TestPython, a fresh WorkRoot and ReportPath. Its thirteen controls exercise bounded reads, strict receipts, absent/empty environment restoration, malformed proof refusal, fixed OS redaction, the vendor's UTF-16 path bound and actual owned descendant timeout/cancellation under the selected native shell. These controls are tooling checks, not final-ZIP or Windows 10 acceptance. The normal scoped shell syntax and analyzer checks remain required.

## Supplemental PowerShell 7 checks

After complete Windows 11 inbox 5.1 acceptance of the exact frozen v1.0.1 ZIP,
run `Invoke-FinalZipPs7Acceptance` against that same application extraction and
its authenticated prior local report. Use the reviewed kit/bindings hash, prior
report hash, adapter hash and an explicit ordinary PowerShell 7.6.6 executable.
Do not derive expected pins from an untrusted manifest or substitute another ZIP.

```powershell
. "$KitRoot/tools/compatibility/Invoke-NativeAcceptance.ps1"
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
