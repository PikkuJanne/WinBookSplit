# M1-T01 — Mechanical engine extraction and Calibre installation

Observed: 8 October 2026, Europe/Berlin. Final full run recorded at
`2026-10-08T21:06:51.400571+02:00` (19:06:51 UTC).
Author: Codex primary; independent read-only source/receipt/privacy review by
the review agent, with separate extraction and test agents.
Base: freshly synchronized main `0de84f367f9bd5ddfa3f408a9c29505d7a39633f`,
tree `d16b98e509cc3b71464ad9c988270b915cdf545c`.
Implementation C: `88c2149b3b3034fbd0d7ef23c2382f4b01648e4c`, tree
`b52ef0e7a4407a888cb06c47e7b99b8814431716`.
Tested actual-path digest:
`2a456dfc50d0ae62e93eaa52c10717ed9a90a3d819ab80560b5a5006f6387054`.
The [machine record](M1-T01-extraction.json) contains the exact source map,
output page identities/PDF hashes, child evidence hashes and historical results.

## Change and scope

The existing Python body now ships as `engine/winbooksplit_engine.py`.
`log`, `split_pdf` and `write_slice` retain identical ASTs. The guarded `main`
retains the original CLI logging/stdout configuration and argv dispatch;
importing the module performs no processing or configuration. PowerShell
resolves the file from `$PSScriptRoot`, replaces the embedded/shared TEMP engine
path and removes its generation/deletion. The setup header mentions the engine
directory. Batch, README, license and all three artwork files still match their
original raw Git blobs. No parser, bookmark, output, discovery or application
process-supervision fix is mixed into this extraction.

The explicit extraction harness reads immutable original Git blobs rather than
weakening `expected_original.json` or the original harness. The central full
route now selects extraction equivalence; the historical baseline route still
refuses changed launcher hashes. A test-only timeout defect was reproduced and
fixed so actual launcher tests cannot clean an uncertain child process's outputs.
This does not fix the application's legacy process handling.

## Actual environment and commands

Current Windows 11 Pro 26H2 x64, build 26300.9457; isolated workstation setup,
not a clean OS. Fresh external regular GIL CPython 3.14.8 x64 dev venv,
pypdf 6.19.0, ReportLab 5.0.1, Pillow 12.3.0 and charset-normalizer 3.5.2.
Hash-required binary-only installation, `pip check`, exact version/GIL/platform
and all four import-origin assertions passed. Actual shell hosts were Windows
PowerShell 5.1.26100.9444 and PowerShell 7.6.5. Pester 6.2.0 and
PSScriptAnalyzer 1.25.0 imported from the previously audited external absolute
manifests in both hosts. Children used their own built-in module directory and
process-only RemoteSigned; stored machine/user policies were unchanged.

Variables below are symbolic redactions of actual absolute external paths:
`$DevPython` is this fresh venv, `$Repo` is the checkout, `$ToolRoot` is the
isolated versioned shell-tool root, `$PS51`/`$PS7` are the actual executables,
and `$M1ExternalRoot` is this run's new owned external evidence directory.
Reports are preserved externally; no private document was an input.

| Actual command/test | Actual outcome | Evidence |
| --- | --- | --- |
| Explicit Python 3.14.8 `-I -m venv`; `$DevPython -I -m pip install --require-hashes --only-binary=:all: --index-url https://pypi.org/simple -r $Repo\requirements-dev.txt`; `pip check` | Exit 0; exact versions/import origins/GIL/x64 assertions passed | Environment receipt hash in JSON |
| `$DevPython -I -B $Repo\tests\run_tests.py --layer full --tool-root $ToolRoot --shell-path $PS51 --shell-path $PS7 --report $M1ExternalRoot\full-final.json` from unrelated external CWD | Exit 0, all four stages passed; 20 Python tests, 9 Pester tests per actual host, zero skips | Raw full SHA-256 `742732ac1ecb00e44289d6d25c5cc7de3ebf6b22e19183652b4169929499ddb7` |
| Explicit extraction child, 14 pre/post cases and safe-import probe, AC-011 | Original/extracted complete observations equal: stdout/stderr, exits, filenames, exact PDF bytes and every page identity/order; all three processing ASTs equal | Child SHA-256 `01729072d5a5a835df66fd3cd42f7847ad8ecae9d0ed2229c9f9d7aecd5e9063` |
| Actual PS5.1 `-File`, PS7 `-File`, batch via actual `cmd /d /c`; same three concurrently, AC-012 | Six exit-0 probes from unrelated CWD with matching reference outputs; counterfeit CWD engine and shared TEMP old-engine-name sentinel unchanged | Per-probe host, output hashes/IDs and stderr observations in JSON |
| Both-host syntax and scaffold analysis | Zero syntax errors/new-scaffold findings; 60 retained launcher analyzer observations per host remain limitations | Separate legacy rule counts/child hashes in JSON |
| Actual owned Windows Python parent/descendant timeout and simulated uncertain-tree/close-error tests | Timeout remains failure; real tracked redirector/interpreter/descendant PIDs terminated; simulated uncertain tree preserves marked synthetic output and never closes uncertain pipes | Included in final 20-test suite |
| `$DevPython -B $Repo\tools\codex-handoff\validate_plan.py --plan-root $Repo\docs\codex-v1.0.0` | VALID/exit 0: 36 tasks, 107 acceptance requirements, 30 oracles; structural check only | Tool transcript |
| `git -c core.whitespace=cr-at-eol diff --check` and staged equivalent | Exit 0; original PowerShell raw CRLF/no-final-newline preserved | One-shot option only; no Git config mutation |

AC-011 and AC-012 passed their specified narrow obligations. Three real
parallel launches shared owned TEMP and used distinct input/output GUIDs.
Actual launcher outputs used the observed Documents folder: each new directory
was exclusively reserved with an exact ownership marker and synthetic neighbor.
No private Documents entries were listed or read. Source/input/neighbor/decoy
hashes remained unchanged; all observed owned output directories and run temps
were removed. Cleanup validates containment, GUID/marker, member allowlist and
all reparse attributes before direct-file removal. Unknown ownership/member or
uncertain process termination preserves artifacts; no unconditional cleanup
after every possible initialization/error state is claimed.

Both PS7 piped probes emitted the existing RawUI.CursorPosition invalid-handle
stderr while producing matching PDFs/exit 0. AC-012 verifies engine resolution;
it does not pass console UX or application stderr/error handling. The raw stderr
is retained externally and its digest/description is in the public JSON.

## User-requested Calibre installation

Installed the [official portable 9.15.0 build](https://calibre-ebook.com/download_portable)
per-user at `$UserProfile\Apps\Calibre915\Calibre Portable`, without elevation
or persistent PATH/policy changes. The installer was downloaded from the
[official version directory](https://download.calibre-ebook.com/9.15.0/).
Its 205,319,944 bytes matched the
[GitHub release asset SHA-256](https://github.com/kovidgoyal/calibre/releases/tag/v9.15.0)
`55a8c89bc0739a2dc6d496742ea625fccc6dfdec1413eb805805a28e7227536c`
and the [vendor SHA-512](https://calibre-ebook.com/signatures/calibre-portable-installer-9.15.0.exe.sha512)
recorded in JSON. Windows Authenticode was Valid with signer Kovid Goyal.
GPG verification was NOT RUN.

Actual automated install command:

```powershell
$CalibreBase = Join-Path $env:USERPROFILE 'Apps\Calibre915'
$InstallerProcess = Start-Process -FilePath $DownloadPath -ArgumentList ('"{0}"' -f $CalibreBase) -WindowStyle Hidden -Wait -PassThru
$Converter = Join-Path $CalibreBase 'Calibre Portable\Calibre\ebook-convert.exe'
& $Converter --version
```

New-target/path-length and verified digest/signature preflights passed before
execution. Installer exit 0; converter exit 0 reported `calibre 9.15.0`.
Installed converter SHA-256:
`f46a01c9b8cd392e1190ce2e50a2a5a48b326c8c3eab670bbcd8b298bede501d`;
its Windows signature is Valid. GUI launcher existence was checked, but GUI
interaction and actual EPUB/AZW3 conversion are NOT RUN. No Calibre binaries or
libraries were committed. The trusted portable location is outside the three
unchanged application search paths: discovery integration remains M2-T04,
real format conversion remains M4-T04 and later package/release gates.

## Earlier failures and retained defects

An early checksum receipt incorrectly asserted SHA-512 equality after
nonterminating octet-stream decoding errors. It was invalidated; a fresh
terminating-error rerun explicitly decoded and compared both hashes and checked
the signature before any installer execution. Only the verified receipt is used.
One preliminary full run correctly failed source preservation because its
README was edited during execution. The following stable 17-test run passed,
but represents earlier harness bytes. Review then reproduced a direct-parent
timeout leaving its descendant alive (owned log grew 531 to 951 bytes); that
writer was explicitly stopped/waited before any cleanup. Test-only tracked-tree
termination/fail-closed handling and the uncertain-pipe-close classifier were
fixed and verified in the final 20-test full run. Earlier report hashes and
corrected new-test assumptions are preserved in JSON. An initial root helper
invocation with `-I` failed its sibling import; documented `-B` rerun passed.

Manual `1` still creates zero files with exit 0; `1,4,7` still loses pages 1-3;
invalid tokens are still filtered. Level 1 still omits front matter; nested
Level 2 A2 still crosses parent boundaries. These are understood preserved
defects, not repaired-engine acceptance. Explorer, all launcher/error/discovery
paths, actual conversion, fidelity/rendering, CI, clean OS/extracted-package and
release/tag/public-download gates are NOT RUN. Windows10/ARM/UNC/other Python
remain unclaimed. CI's canonical task is M4-T05; earlier M0 prose calling it
M5-T01 was a documentation error, not a CI result.

## Git, review and next thread

Branch `codex/winbooksplit-v1-m1`, [draft PR #5](https://github.com/PikkuJanne/WinBookSplit/pull/5).
C was normally pushed; clean local HEAD equaled the fresh live feature SHA at
`2026-10-08T19:08:08.466690+00:00` (21:08:08 Berlin). Every tested raw-path
hash independently matched clean C. Tests ran on the pre-commit worktree;
no clean-C full rerun is claimed. Independent source/raw-evidence/privacy review
found no remaining blocker. No GitHub submitted review or CI run is claimed.
The live audit found main `0de84f3`, no protection/rulesets/required checks,
workflows/runs, tags or releases; settings were untouched.

This following documentation-only evidence/status checkpoint E references C.
Its final normal-push/live receipt belongs in the thread/PR, not inside itself.
Only intended source/test/docs paths were committed; raw reports, fixtures,
dependency binaries and private artifacts remain outside Git. No tag/release,
merge, force push, reset, stash or branch/history deletion occurred. The M1 cumulative merge
gate remains M1-T06.

Next: **M1-T02 — Fix manual start-page parsing**, after a fresh live check.
Use the strict contract/oracles and start with the reproduced first-page/token
defects. Preserve historical original/M1 receipts and guards; deliberately
transition affected current tests from mechanical equivalence to corrected
behavior when the engine changes. Do not call fixed code the unchanged original
or silently rewrite historical known-bad expectations.
