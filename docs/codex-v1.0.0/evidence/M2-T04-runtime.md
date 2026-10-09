# M2-T04 — Interpreter and converter preflight

M2-T04; AC-043/044/045 **PASS**. Date: 9 October 2026; timezone Europe/Berlin.
[Machine record](M2-T04-runtime.json) records actual source/runtime identities,
commands, native outcomes, reduced observations and external raw receipt hashes.
Base `c7af98e89b9b9d96acd047e24d6ee2eba199ba77`, tree `3b207b02b112ea8017c4d78f25e7ab55cf1fa4de`.
C `e9bcef45d826a5f94679c9e24a0cc67dd6017e43` is clean/live **SYNCED** at `2026-10-09T12:19:12.811580+00:00` on `codex/winbooksplit-v1-m2`.
Final unrelated-CWD twelve-stage gate passed 186 Python tests, 28 Pester tests per actual supported host and all nine inherited/current regression stages. Tested 67-path raw digest: `8ea3467230d88e2438b2daa7da6b61ea79465ad3e177e083a57d1a00e4f41e66`. All tested raw bytes match clean C; no clean-C full rerun is claimed.

## Changes and reproduced behavior

Preserves BAT, console/drag-drop, Python/pypdf, optional Calibre and existing plans.
The new shipped PowerShell helper selects explicit PythonPath, application .venv,
read-only validated py.exe -0p listing, then ordinary trusted PATH python.exe.
Explicit or existing broken .venv is fatal without fallback. Literal absolute
ordinary files/ancestors are required; reparse and short-name aliases reject;
automatic hard-link executables and book/CWD descendants reject before execution.
Empty/relative PATH entries are excluded. No book-directory dependency shadowing.

The exact selected interpreter probes regular Windows x64 CPython 3.14.8 and
pypdf 6.19.0 distribution/import origin with -I -B, then launches the existing
engine with -I -B. Chosen paths, versions, source and pypdf origin are printed/logged.
Dependency failures preserve structured attempts/native diagnostics and exact
interpreter hash-pinned setup guidance before console/conversion/output reservation.
Installation remains an explicit separate action; no auto-install, registration,
elevation, global PATH/policy/settings change. PythonPath is appended as the last
parameter, preserving earlier parameter order. Existing 0/1/55 codes remain.

Only ebooks resolve CalibrePath, trusted PATH or known installations and validate
the exact 9.15.0 banner. Real Calibre's optional exact creator line is supported;
arbitrary extra output/version mismatches/native failures reject. PDF never probes
Calibre, including a supplied invalid converter path. Existing owned-job conversion,
immutable captured snapshot and optional full-PDF retention remain unchanged.

Before edits, 14 actual frozen whole-launcher cases (seven/host) reproduced real
registered 3.14.7 selection, missing pypdf after console creation, executable/module
book/CWD/PYTHONPATH shadows, ignored application venv and portable PATH. Separate
scoped original-function probes expose bare-python/import behavior; these are not
full application passes. All reproduction source snapshots remain immutable.

## Actual commands and acceptance

Windows 11 Pro 26H2 build 26300.9457 x64 workstation; not clean OS/VM/Windows 10 evidence.
Registry ProductName Windows 10 Pro is a compatibility label. Fresh official regular
GIL CPython 3.14.8 hash-pinned developer venv: pypdf 6.19.0, ReportLab 5.0.1, Pillow 12.3.0,
charset-normalizer 3.5.2; isolated imports/pip check 0. All 66 saved shell-tool hashes/
sizes match; Pester 6.2.0 and PSScriptAnalyzer 1.25.0. Actual separate hosts
PS 5.1.26100.9444/PS 7.6.5. Already explicitly installed portable Calibre 9.15.0,
converter SHA256 f46a01c9b8cd392e1190ce2e50a2a5a48b326c8c3eab670bbcd8b298bede501d.
No new converter install or full install-tree reproducibility claim.

Final actual command from unrelated CWD (resolved argv/times/native status in JSON):

```powershell
& <M2-T04-root>\dev-venv\Scripts\python.exe -I -B <repo>\tests\run_tests.py `
  --layer full --tool-root <shell-tool-root> --calibre-path <Calibre915>\ebook-convert.exe `
  --shell-path <WindowsPowerShell-exe> --shell-path <PowerShell7-exe> `
  --report <M2-T04-root>\full-final-02.json
```

All 12 stage statuses 0, no unit/Pester skips, syntax/scaffold findings. Legacy
application analyzer observations remain separate limitations. Actual policy
reads before/after match; process-only RemoteSigned requested in test children.
LongPathsEnabled=1 was observed, never enabled. Raw streams/encoded commands and
private profile paths remain external; hashes/byte sizes and aliases are public.

| Acceptance | Actual result |
| --- | --- |
| AC-043 | Ten whole-host refusals: missing executable, native non-Python stub exiting 7, real unsupported 3.14.7, missing pypdf, wrong copied pypdf per host. Eight selections: explicit literal, .venv, multiple PATH, authored py listing per host. Two executable/module shadow controls. Exact isolated selected engine/import, early nonzero and source/neighbor checks pass. |
| AC-044 | Two actual PDF splits with Calibre absent and an owned version-spy supplied; no version call/conversion, complete three physical pages/content. |
| AC-045 | Six actual missing/wrong/book/CWD decoy converter refusals. Five real portable EPUB splits: explicit/PATH per host plus unchanged BAT PATH; exact selected version/path, three physical pages/content and manifest checked. |

Runtime 33 = 16 native 1 refusals before console/conversion/output, 17 native 0 splits,
51 reopened chapter PDFs. Every physical page 1..3 is present once with independent
content vectors, actual result frame/manifest/writer parity and exact known cleanup.
Read-only inputs/neighbor SHA/size/device/inode/attributes and decoy bytes remain
unchanged; positive controls prove decoy markers work, and acceptance leaves them
absent. Final runtime/full owned temp cleanup passes. Inherited conversion repeats
eight real EPUB/genuine AZW3 host/retention launches, two captured-reader controls,
eight native zero invalid PDF refusals and actual BAT early missing converter refusal.
Inherited mode/plan/output/path/diagnostic guards, fixtures and original snapshot
are preserved; full-stage counts/hashes are in the machine record.

Earlier 19 focused runtime Pester tests pass separately per host before the later
test CWD correction; final full 28 includes all 19 runtime and existing nine tests.
Four synthetic validator tests reject 47 evidence contradictions
plus missing/duplicate exact cases; these are structural tests, not Windows passes.
The strict independent runtime validator requires exact 33 IDs and meaningful
native/selection/page/content/identity/policy/cleanup observations. ENGINE log rows
show launch intent; actual successful result plus reopened outputs prove execution.

## Defects, review and remaining limits

Initial focused PS5.1 root-array test failed because ConvertFrom-Json unwraps a
singleton array; lexical object guard corrected it. Actual compatible Calibre's
two-line banner initially rejected under a one-line anchor; exact optional creator
line then passed both hosts and focused 19. First full gate failed typed Pester
container += before Pester and an obsolete inherited BAT missing dependency scope;
ordinary array assignment and strict new early-refusal fields corrected them.
The next focused Pester run exposed the PATH test's candidate below encoded-runner
CWD: correct shadow exclusion prevented the intended incompatible probe. Setting
and restoring the test's trusted repository location preserves exact actual-probe
assertions and separate shadow guards; before/after receipts remain external.
Earlier failed/partial/passing sources, argv typo and corrected read-only audit
errors remain distinct; no failed run is relabeled final acceptance.

Independent production/schema/validator/harness review across agents and actual
full raw/source/public evidence review found no blocker. Harness authorship limits
are explicit in the machine record; local review is not GitHub approval/CI. [Draft continuation PR](https://github.com/PikkuJanne/WinBookSplit/pull/14).
C `e9bcef45d826a5f94679c9e24a0cc67dd6017e43` is clean/live **SYNCED** at `2026-10-09T12:19:12.811580+00:00` on `codex/winbooksplit-v1-m2`.
E records observed C; E's own push/live receipt stays external/PR/thread.
At start PR13 was already merged at 11:23:44Z at c7af98e8; normal fetch/fast-forward
reconciled local main/M2 and normal M2 push, no reset or cumulative M2 acceptance inferred.
No milestone merge/tag/release created here; fresh tags/releases/workflows inspected.

Authored py listing selects the real interpreter but does not test real manager
registration/installation. Actual manager read-only listing exposed only 3.14.7.
Wrong pypdf 9.99 is a declaration changed only in owned copied package, not a released
version. Real 3.14.7 rejection is tested; broader unsupported-runtime compatibility
is untested. Trusted installation probes are compatibility checks, not executable
authentication or a same-account sandbox. Automatic short/reparse/hardlink checks
do not promise arbitrary UNC/long-path support.

Preflight's narrow native helper concurrently bounds both 64 KiB tails, deadline
includes parent exit/pipe EOF, and timeout kills/waits only its exact parent. It
records DescendantsStopped=null; no descendant-stop or cleanup authority follows.
Existing Calibre conversion still uses owned job proved-tree-stop before cleanup.
Inherited pipe threads may outlive a timed-out probe until EOF; engine process
remaining bounds/deadlines/tree/UTF-8 work and cumulative M2 review belong to M2-T05.
Existing source check/create/lock-close/crash limitations persist.

Automatic approval review rejected removal of two earlier generated partial sample
workspaces with “blocked by policy” and no detailed reason. No deletion occurred
and no alternative-shell retry. These copied samples remain outside Git; all six
sample application outputs already passed checked known cleanup. This residual is
separate from final 33/full-suite owned temp cleanup, which actually passed.

M3 owns standalone version/noninteractive CLI/menus/preview. Human Explorer/GUI/
Ctrl+C, rendered fidelity/PDF features, expanded/remote-resource/DRM rejection,
CI/exact package/public release gates remain open. Stop here; next M2-T05 only.
