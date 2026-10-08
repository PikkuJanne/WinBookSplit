# M0-T03 — support/dependency selection evidence

Task: M0-T03, AC-007 / AC-008.
Date: 8 October 2026, Europe/Berlin; machine observation
`2026-10-08T17:30:03.208436+00:00` (19:30:03 local).
Author: Codex; independent read-only dependency research and support review.
Base commit after reconciliation: `a66c8f95f922c36c58b47b6cbfba82be399a552a`.
Base tree: `edef8ef19e24c6620eecf8fb4522d93a9a60071c`.
Tested pre-commit paths digest:
`25820fd2997867e091c88237be105db0aeed1ac34bd0c525f8aacca17c24adc0`.
Seven unchanged original application paths digest:
`6ec53a82c7013b28d9ffd4fa301190ae5e6d0e0d9ffb670ccd2863fea3d06d89`.
Individual raw SHA-256 values and exact dependency wheels are in
`M0-T03-environment.json`. This describes the tested worktree paths, not an
invented clean-commit test run or the hash of its own future evidence commit.

## Changes and decisions

`SUPPORT_AND_SETUP.md` freezes regular GIL x64 CPython 3.14.8 only, plain
pypdf 6.19.0 and the Calibre 9.15.0 conversion target, with Windows 11 and both
specified PowerShell hosts. Minimum and selected Python are the same; older
helper/runtime observations do not become a broad supported Python range.
The root runtime/dev requirements pin and hash only the necessary wheels.
ReportLab 5.0.1, Pillow 12.3.0 and charset-normalizer 3.5.2 are developer-only.
Stdlib unittest, Pester 6.2.0 and PSScriptAnalyzer 1.25.0 are selected for
M0-T04; the external shell tools remain uninstalled/untested.

Current official Python/pypdf/ReportLab security information justified versions
newer than the historical baseline. SOURCES.md retains the exact primary links
and review boundaries. A manual declared-range review and vendor notes are not
an automated scan. No source splitting, launcher, conversion logic, framework,
public release or global dependency/policy setting changed. All seven original
Git blobs and raw working bytes still match the recorded baseline; Git
`hash-object --no-filters` is necessary for the raw-byte comparison with this
checkout's `core.autocrlf=true`. An initial filtered-hash assertion failed;
the corrected raw comparison passed without changing source files.

## Actual environment and commands

Windows 11 Pro 26H2 build 26300.9457 x64. Windows PowerShell 5.1.26100.9444,
CLR 4.0.30319.42000; PowerShell Core 7.6.5, both 64-bit. CIM/Python identify
Windows 11; the registry's stale Windows 10 product label is recorded, not used
to claim a Windows 10 test. Existing registered Python is 3.14.7 with no pypdf;
bundled historical Python is 3.12.14/pypdf 6.10.0/ReportLab 4.4.9. Calibre was
absent from PATH, all three original configured locations and the additionally
inspected D: Program Files location. Existing Poppler is 26.07.0, inbox Pester
3.4.0 only, no PSScriptAnalyzer. New Python was explicitly extracted using the
existing Python manager 26.3; venvs bootstrap pip 26.2.1.

`$ProbeRoot` below denotes the uniquely created external TEMP directory for
this task. Raw reports/venvs remain there; public evidence omits workstation
usernames/absolute temporary paths. No activation, global install, elevation
or execution-policy change was performed.

| Actual command / check | Exit and observed result |
| --- | --- |
| `git status --short`, branch/upstream/origin/worktree/log/`ls-remote`; `gh auth status`, live PR/main/tags/releases reads | Initial clean feature HEAD/live SHA `869218c3cc4c8b7f6b47d057a5e8e4a414924173`; authentication usable; tags/releases empty. PR #2 had already merged at 17:16:55Z; live main `a66c8f9...`. |
| `python tools/codex-handoff/check_sync.py --repo .` at start | Exit 0, SYNCED at `2026-10-08T17:18:14.713451+00:00`; this was the starting feature checkpoint, not final task synchronization. |
| `git fetch origin`; `git merge --ff-only origin/main` | Both exit 0; existing feature branch incorporated the actual PR #2 merge, no rewrite/conflict/runtime diff. |
| `python tools/codex-handoff/validate_plan.py --plan-root docs/codex-v1.0.0` | Exit 0; 36 tasks, 107 acceptance cases, 30 oracles structurally valid. Not application correctness. |
| Windows CIM/registry, both `$PSVersionTable`/64-bit probes, `Get-Command`, `py list`, isolated interpreter imports, `Get-Module -ListAvailable Pester,PSScriptAnalyzer`, `pdftoppm -v`, literal Calibre file checks | Actual versions/absence above; no inferred Calibre, Windows10, external-module or renderer upgrade pass. |
| `py help install`; `py install --target=<new $ProbeRoot/python3148> 3.14.8` | Exit 0; manager reported official index signature verified and extracted 3.14.8. `py list` afterwards still lists only registered 3.14.7. |
| Explicit extracted `python.exe -I -m venv <new runtime-venv/dev-final-venv>` | Exit 0, separate fresh environments outside Git. |
| Runtime Python `-I -m pip install --require-hashes --only-binary=:all: --index-url https://pypi.org/simple --report <runtime-install.json> -r requirements.txt`; `-m pip check` | Exit 0 for both; pypdf 6.19.0 only beyond bootstrap pip. Assertions confirm regular GIL x64 3.14.8, venv prefix/import origin and absence of ReportLab/Pillow/charset-normalizer. |
| Final dev Python same install flags/report with `requirements-dev.txt`; `-m pip check`; package/version/import-origin checks | Exit 0; pypdf 6.19.0, ReportLab 5.0.1, Pillow 12.3.0, charset-normalizer 3.5.2. Every mandatory package has an exact reviewed wheel hash; no optional extras. |
| Final dev Python `-I tests/baseline/characterize_original.py --launcher-probes --report <baseline-final3148.json>` | Exit 0; all 14 original engine expectations and two actual unchanged early launcher error paths reproduced. Reopened all slices with exact page IDs; source/input/neighbor preservation and repeated fixture equality passed. This is known-bad characterization, not fixed-engine acceptance. |
| Complete runtime setup block copied from `SUPPORT_AND_SETUP.md`, with only trusted interpreter/application paths substituted, executed separately by `powershell -NoProfile -EncodedCommand` and `pwsh -NoProfile -EncodedCommand` | Both exit 0: fresh venv, hash-verified runtime install, pip check, exact interpreter/import assertions. Applications used new external paths with spaces and ü. Initial capture decoded shell console bytes as UTF-8 and replaced some path characters; no Unicode console-fidelity claim. PS5.1 stderr contained module-initialization progress CLIXML; PS7 stderr was empty. |
| Corrected documented `-c` assertions and final developer version checks under both shell hosts | Both exit 0. These are dependency/setup checks, not full launcher/Windows splitting tests. |
| Primary-source reads; `gh api 'repos/py-pdf/pypdf/security-advisories?per_page=100'`, same for Calibre | API reads exit 0; 54 / 21 records, no next page observed. No returned declared affected range includes selected 6.19.0 / 9.15.0 in the independent manual review. API snapshot digests retained. |
| Raw seven original blob/hash comparison and `git diff --check` | Passed raw original identity and whitespace check; application unchanged. |

Preliminary failures are preserved in the JSON. A temporary PS5.1 `-File` probe
was blocked by the workstation's execution policy before running; policy was
left unchanged and the documented interactive commands were checked directly.
The initial setup quoting failed under PS5.1 because native argument handling
stripped Python double-quoted literals; the independent review reproduced it.
The documentation now uses compatible quoting, and the exact complete setup
block passed in both hosts. An exploratory 4.4.9 developer environment passed
characterization, but its selection was superseded after reading the vendor
security fix; the final 5.0.1 environment was fresh and reran all affected checks.

## Acceptance and limits

AC-007 satisfied for support selection: actual tools and current primary
requirements/advisories were inspected; exact targets and unverified platforms
are recorded. Calibre absence is a real probe outcome, not a conversion pass.
AC-008 satisfied: separate hash-pinned runtime/developer sets and explicit
isolated setup are executable; the runtime-only venv has no developer packages;
Calibre is required only when users process ebooks. The application's future
discovery/preflight enforcement is M2-T04, not implemented here.

New fixture hashes on the final selected dependencies: simple10
`e5331ca3852d0581aece3b72f590537b344961840e553c9a6e6ce298f1023ac5`, nested12
`a4cd8731da7ff93e25bec5181481a42ed7f1bd6c2dfe6818e2c4f061df35d245`.
Both repeat byte-for-byte within this environment and retain the known page
IDs/outlines. Cross-version byte equality was never promised; M0-T02 hashes
and actual render receipts remain historical. No new visual renderer pass is
claimed for these dependency-generated bytes.

No M0-T03 acceptance ID was skipped. Full PS5.1/7 success splitting, Explorer
drag/drop, EPUB/AZW3 conversions, current renderer validation, external shell
tools, repaired-engine and release-package checks are NOT RUN. Windows10/ARM/
UNC/other Python configurations remain unsupported/unverified. Human checks
and an automated vulnerability scanner: NOT RUN. No private documents,
generated PDFs or binary dependencies were committed or uploaded.

## Git and continuation

Branch: `codex/winbooksplit-v1-m0`. PR #2 was already merged; create a new draft
PR for this remaining M0 support/scaffolding work. Independent read-only review
confirmed wheel hashes/dependency separation/scope and found the now-corrected
PS5.1 quoting issue. No GitHub review submission or CI pass is claimed.
M0-T04 retains the cumulative M0 review/test-scaffolding/merge gate.

This implementation checkpoint will be committed as C, normally pushed and
live-verified. The following evidence/status checkpoint records that actual C
receipt. Its own final receipt belongs in the thread/PR, avoiding a self-SHA
loop. No release or tag was created.

Next exact task: **M0-T04 — Add local test and evidence scaffolding**. Read its
task/TESTING/GITHUB_WORKFLOW and the frozen support plan; implement the small
local runner, isolate/validate the selected shell tools under both hosts,
connect acceptance/evidence IDs, and perform the cumulative M0 review/merge.
Do not call the original known-bad harness repaired-engine acceptance.
