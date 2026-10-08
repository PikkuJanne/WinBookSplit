# M0-T04: local harness and reviewed M0 merge

Task: M0-T04, AC-009 and AC-010. Date: 8 October 2026, Europe/Berlin.
Executor: Codex for repository owner PikkuJanne; separate read-only agent
reviewed cumulative M0 source, receipts, privacy and the final scaffold.
Machine-readable results, commands and individual source/tool hashes:
[`M0-T04-harness.json`](M0-T04-harness.json).

## Source and scope

Initial clean feature checkpoint `fc540e9a6e504dfe62a16c48dcd9a6f34addee63`
was freshly SYNCED at `2026-10-08T17:45:25.758703+00:00` (19:45:25 Berlin).
PR #3 had already merged. A normal fetch/fast-forward incorporated live main
`644ddc101e104dbe2867cd53f20f5a0548e893b1`; recorded older baselines were not
reset targets. No unrelated edits existed or were discarded.

Implementation C: `01ab01bccbbe16ea2bc03eeca0153b18b7c5bfee`.
The full/probe tests ran on the pre-commit worktree at base `644ddc1`, with
actual tested-path digest
`fc8e5fd5ebb1001beba6d1a4f072b9172fc444410336c029e84841834c1bd536`.
Recipe: SHA-256 of UTF-8 canonical JSON mapping relative paths to raw byte
SHA-256, sorted keys and comma/colon separators. The JSON retains the map and
base Git tree. Clean C and merged main actual bytes independently matched
that map; a clean-C full rerun is not claimed. Core autocrlf remains enabled;
raw hashes and `git hash-object --no-filters` were used where appropriate.

Added one `tests/run_tests.py` entry point with stdlib unittest, targeted/full
layers, safely drained native streams, explicit exit aggregation, exclusive
external reports, validated child evidence, source/evidence hashes and owned
temporary directories. Shell checks import exact external manifests, gate
syntax and new-scaffold security/defect/compatibility rules, and run Pester.
Ignore rules protect generated documents, tools, environments and raw outputs.
No runtime refactor, dependency-pin change or behavior fix occurred. All seven
original application files still match the original raw hashes/Git blobs.

## Actual environment and commands

Windows 11 Pro 26H2 build 26300.9457 x64, isolated current workstation setup;
not a clean OS. A fresh external dev venv used the existing official extracted
regular GIL CPython 3.14.8 x64. Hash-required/binary-only install of
`requirements-dev.txt`, `pip check`, exact interpreter/package-version and
venv import-origin assertions passed. Versions: pypdf 6.19.0, ReportLab 5.0.1,
Pillow 12.3.0 and charset-normalizer 3.5.2. No global install occurred.

`Save-Module` saved Pester 6.2.0 and PSScriptAnalyzer 1.25.0 to a new external
directory. Both actual hosts imported absolute versioned manifests and
verified versions/origins: Windows PowerShell 5.1.26100.9444 (CLR
4.0.30319.42000) and PowerShell 7.6.5 (CLR 10.0.11). The JSON records all 66
saved file sizes/hashes, independently verified against actual bytes. Only
extracted module receipts exist; no package-archive hash is invented.

Fresh PS5.1 effective policy was Restricted (all stored scopes Undefined).
The developer children explicitly request process-only RemoteSigned and use
their selected host's own built-in module directory, with vendor tools imported
absolutely. Stored machine/user policies before and after are unchanged;
PS7 LocalMachine remains RemoteSigned. Managed policies take precedence.
Original early launcher characterization retains its existing process Bypass
invocation; no application success-path/policy acceptance is implied.

Variables below denote the actual absolute fresh interpreter, runner, external
tools, two hosts and unique external report files. Exact redacted argv and raw
receipt hashes are in the JSON; `tests/README.md` supplies reproducible setup
and invocation examples.

| Actual command/layer | Result | Scope |
| --- | --- | --- |
| `$DevPython -I -B $Runner --layer python --report <new external>` | Exit 0; 11 tests, no skips | Final targeted runner regressions |
| `$DevPython -I -B $Runner --layer shell --tool-root $ShellToolRoot --shell-path $PS51 --shell-path $PS7 --report <new external>` | Exit 0; 9 Pester tests per host, no skips; zero syntax/new-scaffold findings | Actual targeted both-host harness |
| Same explicit paths with `--layer full`, twice from different unrelated directories | Exit 0 each; 11 Python tests, 9 Pester tests per host, 14 engine cases plus 2 early launcher boundary probes | Full relevant M0 layer; original behavior is known bad |
| `--layer python --failure-probe native` | Native 23, later success 0; runner 1 | AC-009 native propagation |
| `--layer python --failure-probe python` | Failed unittest 1, later success 0; runner 1 | AC-009 Python propagation |
| `--layer shell --failure-probe pester` with both hosts/tools | Each host has one failed Pester test, then successful native command; final central success 0; runner 1 | AC-009 actual both-host Pester propagation |
| External `$DevPython -B -m unittest discover -s bundle_tests -v` | Exit 0; 57 tests in 99.138 seconds: 56 passed, one symlink-creation skip | Historical handoff importer/helper fixtures; three shipped helpers byte-identical |
| Plan validator, staged whitespace checks and original SHA/blob comparison | Passed | 36 tasks/107 acceptance requirements/30 oracles; structural consistency only |

AC-009 PASS: all real failing fixtures retain nonzero status through later
success and report writing. AC-010 PASS: two complete runs produce the same
deterministic layer outcomes and source digest, use distinct owned directories
that are removed, and preserve every checkout/input/neighbor byte. Unrelated
directories contain synthetic sentinels and hostile same-name Python modules;
isolated trusted imports do not load them. All five reports verify owned cleanup
and unchanged sources. The outer snapshot covers all non-Git checkout bytes.
Raw acceptance summary SHA-256:
`95d2da6824a03102ff0ce4d3c375d2613c16a21cde5c622dc46d146749f18188`.

Review reproduced an invalid shell preflight that nevertheless wrote a report
from finally. Validated canonical destination gating fixed it; two permanent
Pester regressions pass in both hosts. Empty child evidence could also conceal
exit-0 omissions; schema/status/nonempty-result checks and Python regressions
now fail it. Initial PS5.1 policy/module autoload failures led to the explicit
allowed process policy and trusted host module path. Preliminary failed probes
remain observations rather than passes.

The original application has 68 analyzer findings per host: 51 WriteHost
warnings, 15 trailing-whitespace information findings and two approved-verb
warnings. These are separately recorded legacy observations, not a runtime
static pass. Original page loss, invalid-token filtering, zero-file success,
front-matter omission and parent crossing remain reproduced defects.

## Review, synchronization and merge

Independent cumulative review found no remaining M0 merge blocker and checked
the actual repeated/probe receipts, module hashes, all seven original files,
source/path/process/cleanup/privacy rules and intended staged paths. Exactly
12 implementation files were committed, with no document/dependency binaries.
No GitHub review submission or approval is claimed.

Normal push of C succeeded. Clean local C equaled fresh live feature C at
`2026-10-08T17:59:55.797912+00:00` (19:59:55 Berlin), SYNCED.
[PR #4](https://github.com/PikkuJanne/WinBookSplit/pull/4) was non-draft,
MERGEABLE/CLEAN at exact C. Main was unprotected, rulesets/required checks and
workflow/run counts were zero. CI execution is NOT RUN; M5-T01 owns it.
The reviewed merge used `--match-head-commit C`, with no admin bypass or branch
deletion, at `2026-10-08T18:01:01Z` (20:01:01 Berlin).
Merge commit: `ae9ddd0444e61578ed59625d6d3df4038e6fb09e`.

`git fetch origin main:main` normally fast-forwarded local main, then
`git switch main` retained identical tested bytes and the feature history.
Clean local/live main equality was freshly SYNCED at
`2026-10-08T18:01:05.818622+00:00`. This following documentation-only closure
records these already observed C/merge receipts. Its own final push/live SHA
belongs in the thread, not inside this commit. A failed final push means an
incomplete checkpoint regardless of the task's recorded status.

## Limits and next thread

No M0-T04 acceptance ID is skipped. The external handoff symlink test is the
one truthful non-acceptance skip. Full repaired splitting under both shells,
human Explorer, actual EPUB/AZW3 Calibre conversion, current renderer QA, CI,
automated vulnerability scanning and exact extracted-release/package checks
remain NOT RUN. Calibre is still unavailable. Windows10/ARM/UNC/other Python
support is unclaimed. Signing remains optional with unsigned disclosure;
live tags/releases were empty and v1.0.0 is not published.

Next exact task: **M1-T01 — Extract the existing Python engine narrowly**.
Use the stable known-bad results for mechanical-equivalence checks before any
fix. Follow NEXT_SESSION.md and the task/specs; do not start M1 in this thread.
Raw reports, module files and fresh venv remain outside Git. No private input
was processed, committed or uploaded.
