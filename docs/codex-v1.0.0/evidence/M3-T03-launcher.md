# M3-T03 — Launcher and menu evidence

**AC-057/058/059: PASS.** The final frozen full gate and independent actual
Explorer audits passed. The implementation is C4 below; the documentation-only
evidence checkpoint receives its actual clean/live receipt after normal push.
[Machine record](M3-T03-launcher.json) includes the complete96-path source map,
actual commands and raw external receipt names/hashes/sizes.

## Change and preserved behavior

The existing BAT/PowerShell console now supports an interactive no-file launch,
all three PDF/EPUB/AZW3 formats, and explicit refusal of extra dropped arguments.
BAT disables delayed expansion, preserves quoted literal paths, invokes its
adjacent script from unrelated directories, and returns the saved native exit.
Quoted-empty argument 2 with a third argument is also rejected before PS starts.

Missing interactive input prompts for one literal path; blank or exact C cancels
before dependency probes/log/output setup with 130. Paired outer quotes are removed
without expansion or evaluation. NonInteractive/Preview still reject absent input.
Initial choices are exact trimmed 1/2/M/C; invalid words and blank answers reprompt.
Existing fallback M/Y, offered 1 and N/C remain exact; closed fallback input cancels.
UTF-8 input preserves Unicode. Cancellation uses neutral text without fake success.
The shared validated plan, writer, literal argv, concurrent bounded streams,
owned-process supervision, immutable inputs and authenticated cleanup remain intact.

## Source, environment and synchronization

- Base main: `56d9d4a9556520b8b92426f44f74fe7faceb4d7b` (PR18 merge).
- Implementation C4: `302b063322b7ef341cfc7b7a1f7aa3cb9d5fcea8`.
- C4 tree: `5de6ef93ffc04e33ac45de24c2ea946c335c51e7`.
- All 96 runtime/test source-manifest paths equal exact C4 Git blobs; raw digest
  `63bed7c22f00e655016b018dde3394b8535fb1ad2af06d8181d4ec0e1b5e9eec`.
- Clean local C4 equals the fresh live M3 branch at `2026-10-09T18:47:27.008501Z`.
  This sync receipt certifies branch equality, not CI, release or formal approval.
- The UI kit's 12 application files exactly equal C4, with no PS instrumentation
  or changed parameter defaults. Its retained manifest records earlier C9fd4477;
  the later two commits change test tooling only. Manifest SHA-256:
  `def345d8cc0fb29d769a8b1eef68de685248df3e47e8fd44de5eeaed45f4c2c6`.

Current Windows x64 workstation: OS identification Windows 11, build 26300.9457,
26H2. Registry ProductName retains the legacy label Windows 10 Pro; this is not a
Windows 10 support claim. Actual hosts are PS5.1.26100.9444/.NET4.0.30319.42000 and
PS7.6.5/.NET10.0.11. Regular x64 CPython 3.14.8, pypdf 6.19.0, real Calibre 9.15.0;
dev-only ReportLab5.0.1/Pillow12.3.0/charset-normalizer3.5.2, Pester6.2.0 and
PSScriptAnalyzer1.25.0. Fresh isolated developer and UI runtime venvs are recorded.

The exact-byte UI kit uses ordinary Documents output and application-local runtime
discovery. A task-owned copy of the supported portable Calibre temporarily occupied
an initially absent standard per-user location; ebook logs record known_location, version 9.15.0.
After UI checks, the sealed copy was moved on the same volume to an absent external
task-owned target. The original portable runtime was preserved. Independent audit
matched all 1,345 files, 83 directories, 662,902,308 bytes, hashes and native identities;
standard discovery is absent again. No overwrite, deletion or global settings change.

## Actual Windows interactions (assistant-operated)

| Acceptance | Actual observed result | Independent review state |
| --- | --- | --- |
| AC-057: PDF drop | Explorer parent and exact literal BAT/PS argv; native PS/BAT 0; 3 PDFs, all 6 physical pages in order. | PASS: independent v9 drop audit. |
| AC-057: EPUB drop | Genuine authored EPUB, real conversion; native PS/BAT 0; 3 PDFs, all 3 generated physical pages. | PASS: independent v9 drop audit. |
| AC-057: AZW3 drop | Genuine Calibre-generated AZW3, real conversion; native PS/BAT 0; 3 PDFs, all 4 generated physical pages. | PASS: independent v9 drop audit. |
| AC-058: double-click choose | Exact BAT opened the literal source prompt; chosen PDF completed with native PS/BAT 0, 3 PDFs/6 physical pages. | Verified with stated capture limits. |
| AC-058: repeat and C | Complete cancelled/130 frame, zero writes/no engine; native PS/BAT 130, no error/success or extra pause. | Verified. |
| AC-059: two-file drop | Actual two-file Explorer selection/drop; explicit all-format refusal, native BAT 2 before PS/processing. | PASS: independent v9 drop audit. |
| AC-059: invalid choices | maybe/Q reprompted at initial/fallback menus; explicit Auto then Manual retry, native PS/BAT 0, 3 PDFs/6 pages. | Verified; saved decisions invalid, invalid, retry/manual. |

Inputs are authored local fixtures with spaces, Unicode, brackets, ampersand,
percent and exclamation marks. Known inputs/prior outputs retained their recorded
identities. Positive runs have reopened output/content/range/manifest/hash checks;
ebook conversion metadata and absence of known generated/staging files are recorded.
No private document was used. Canonical acceptance IDs retain `manual_windows`;
the operator is the assistant, and these new events are not relabeled human tests.

Sky observed real Explorer double-clicks and two-file selection. Its drag calls
returned rapidly without launching the app despite observed source/target icons.
A narrowly guarded external SendInput helper completed the actual Explorer drags:
known Explorer identity/foreground/window geometry/DPI and source hashes, fixed
points, owned left-button hold/move/hover/release, no modifiers or arbitrary input.
Windows Terminal-hosted app windows were not exposed as targetable Sky windows.
App answers therefore used external guarded native console events: held known
BAT/PS handles and creation identities, frozen bytes, exact current app prompt,
no pending keys, one allowed answer/Enter, then an observed transition. No shell
commands, generic window control, instrumentation, settings changes or elevation.
This task-specific skill-guidance override addresses technical targeting/timing
limits, not an automatic approval rejection. Native events are not human keystrokes.

Native observers held process handles before input and captured actual exit codes.
Successful UI captures retained only the final 32 console rows; the OUTCOME prefix
scrolled out. Joint native exits 0, Done, saved operation outcome, manifests and
reopened outputs support the observed successes; complete authoritative stdout-frame
equality is **not** claimed. Source cancellation's whole final frame was captured.
The first mode/exit captures failed on CJK double-width cell counts; later CHAR_INFO
capture resolved that helper defect. No mode answer was resent after the initial write.

## Automated supplement and final gate

Frozen clean-C full03 **PASS**: native aggregate0, all16 stages native0, 235 Python
tests,71 Pester per actual PS5.1/PS7 host,72 CLI and70 launcher controls, zero skips.
Source-before/after and all96 committed blob hashes match; synthetic source,
settings, prior/neighbor identities and known owned suite cleanup pass. Existing
manual/Level1/Level2/shared-plan/diagnostic/output/path/conversion/runtime/process/
outcome stages passed, including real Calibre EPUB/genuine AZW3. A separate read-only
audit revalidated all15 child reports and all captured stream hashes.

Actual full command from unrelated `<evidence-root>`:

```powershell
& '<dev-venv>\Scripts\python.exe' -I -B '<repo>\tests\run_tests.py' --layer full --tool-root '<tool-root>' --shell-path '<PS51>' --shell-path '<PS7>' --calibre-path '<calibre-original>\ebook-convert.exe' --report '<evidence-root>\full-final03.json'
```

Full raw report SHA256 `b59ecc8d89829b10846e86fed8cab5a253497756dd7bb9d5a18f73a7de206704`;
strict audit SHA256 `b19a6a6477e3ba45c15e8baca84cc1ceb57443bed65c1a1ca4eb0a0379872eff`.
Full report recorded UTC `2026-10-09T19:02:43.730609+00:00`.
Launcher partition21PS51/21PS7/28BAT is actual native redirected-input control
coverage. Targeted helper Pester passed13/13 per host; its Read-Host mocks alone
do not establish native EOF or actual Explorer acceptance. Ordinary unit/syntax
success is separate from these actual application/Calibre/UI observations.

Direct PS copies are exact. Successful native BAT controls declare only copied
OutputDirectory/PythonPath/CalibrePath/NoPause defaults; refusal controls use a
copied PS launch sentinel. BAT stays exact. These piped controls supplement the
12-file exact-byte UI kit and do not substitute for its required Explorer checks.

## Preserved failures and evidence limits

Baseline receipts reproduce no-input failure, ignored extra input, arbitrary-word
implicit menu choices, fallback EOF timeout and quoted-empty argument bypass.
Failed launcher-first/second/third receipts retain 0/10/14 completed rows and their
owning workspaces: prompt oracle, pre-menu log discovery and fallback-label defects.
Earlier Pester/source-encoding and native-helper inspect/capture failures are retained.

Full02 is FAILED: 12/16 zero-exit stages, four failed stages, source changed, retained
work. Conversion/runtime missing-converter controls found the temporary UI Calibre;
CLI completed 72 rows then rejected changed source; launcher completed 42 direct-host
rows then BAT returned9009 because its controlled PATH omitted the PS5.1 directory.
235 Python and 71 Pester per host passed individually there, not as a frozen full pass.
The sentinel LF/write_text byte defect was reproduced and fixed with write_bytes;
C4 then added the explicit native host directory to the test PATH. App bytes stayed equal.
Failed drop audits v1–v8 remain failed; final v9 parses saved raw records and
accounts for the exact ten UI entries plus one contemporaneous full02-authored
run9250 allocation. It reads only child-name metadata for that extra directory,
never its contents or private prior data. Five UI successes total15PDFs/25pages;
the known choose/flat and all five saved UI run files remain unchanged. Process-only
PS51 module-path isolation repairs the reviewer query without altering stored policy.
Targeted development aggregates failed source stability despite native success.
Interrupted full01 returned1 with no final report; workspace/shutdown is UNKNOWN,
with no cleanup/pass claimed. A scope addendum rechecks all35 known UI files against
their v9 hashes/sizes/native identities; earlier saved identity comparison covered
choose/flat, while the final v9 drop review checked content/hashes and recorded identities.

Historical M3-T01 user testing remains complete on E9983191/Cb8c6ed0, including four
original PS controls and both successful fresh BAT retries. Its Explorer builder
changed only copied PS OutputDirectory/CalibrePath defaults; original/copy PS hashes
were ea6be0eba97d1915b7f96e7cce83ffe6eab2d8e311d04e3522b17b83373d2378 and
f6e6f59d8dc26f2f07b1e7880551e8297753a1cb8b9e075d6f6431acb780614f. BAT/engine/support
bytes were unchanged. Preserve original BAT255/unlaunched evidence and attestation;
no human check is reopened or transferred to C4, a package or a clean OS.

Independent C4 source review found no blocker, with disclosed authorship of its
Launcher.Tests.ps1 contributor test. It is not a submitted GitHub approval or CI.
Raw commands, process metadata, captures/screenshots and owning workspaces remain
external; public evidence uses placeholders and receipt names/hashes only.
No clean OS, extracted package, rendered-fidelity, CI, tag or public release is certified.

The evidence/status/handoff files mark only M3-T03 done after actual native/UI
review. Commit/push only intended documentation and record clean HEAD/live equality
externally; the final public-evidence review and actual PR receipt remain external.
[Draft PR19](https://github.com/PikkuJanne/WinBookSplit/pull/19) was actually created
at `2026-10-09T19:22:40Z`; CI status rollup is empty. The evidence commit's own
hash is not self-referential. Next task is **M3-T04 —
Show plans and converted-PDF page guidance**; first public v1.0.0 remains in progress.
