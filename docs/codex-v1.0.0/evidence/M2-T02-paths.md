# M2-T02 — Literal paths and destination-aware filenames

Task: M2-T02; acceptance AC-036 through AC-039 **PASS**.
Date: 9 October 2026; user timezone Europe/Berlin; actual UTC timestamps below.
Base: `8a21e49ca45d67e12edad6266a010978533579f1`, tree
`361f2dfeddfee3348134d312b3a5b30010d2f7f0`.
Tested implementation C: `a38d0e248cecd71513a2b06d6595c055e2712b95`; source digest `9da651cc0c13c961cf9e0a0d96d939792771ced264a7f37c0e0107636123f5d9`
(55 actual paths). All tested raw bytes match clean C. Tests ran
before C; no clean-C full rerun is claimed. [Machine record](M2-T02-paths.json)
contains raw report hashes, source map, actual commands, cases and limitations.

## Changes and reproduction

PowerShell resolves literal FileSystem input/provider paths, requires a readable
file and supported extension, and uses literal metadata. `.pdf` directories and
non-filesystem providers fail before processing with exit 1. `-OutputDirectory`
selects an existing literal base; otherwise actual Documents is used. The new
Paths helper rejects observed reparse ancestors, budgets console records and
reserves a console directory with native exclusive creation. Marker/log/writer
setup errors stop before processing, close opened streams and name any retained
directory. BAT, the native ownership helper and existing process-drain logic
remain unchanged.

The existing Python planners still produce one complete immutable reader-bound
plan. Filename preparation removes Windows-invalid/control/format characters,
keeps usable Unicode including emoji, trims terminal dots/spaces and supplies
stable `Section n` fallbacks for empty/punctuation-only titles. Number width is
`max(2, digits(section_count))`; shared validation rejects casefold collisions,
unsafe components and reserved names, including superscript COM/LPT variants and
CONIN$/CONOUT$. Source names/paths are never repaired or written.

`prepare_split(..., output_base=...)` freezes destination-aware names and a
shortened run stem before preview. Execution consumes those exact entries and
rejects a changed resolved base. Legacy unbound plans retain their names and
reject an impossible destination before allocation. Current budgets are 259 UTF-16 units for files,
247 for created directories and 255 for components, including all stage/final/
owner/manifest/failure paths; no arbitrary long-path or UNC support is claimed.
These are conservative application limits, informed by Microsoft's
[Windows naming rules](https://learn.microsoft.com/en-us/windows/win32/fileio/naming-a-file)
and [path-limit documentation](https://learn.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation).
No registry, stored execution-policy or elevation change was made.

Before changing behavior, exact Git-blob snapshots reproduced wildcard metadata
selecting a different PDF on both hosts, directory/provider arguments entering
processing, control-title rejection, empty/punctuation names, wrong lexical order
for 120 outputs and a full path beyond the chosen conservative budget. Raw reserved
basenames were accepted by the validator; this is defensive-validation evidence,
not a claim that prefixed baseline titles wrote device files. Unicode already
worked in the corrected UTF-8 probe. Initial cp1252/locale probe confounds remain
separate from production defects. Exact snapshot blobs have mixed line endings (README/PS already CRLF, engine/
Diagnostics LF); current raw hashes separately record normal autocrlf materialization.

## Actual execution

Environment: actual Windows 11 Pro 26H2 build 26300.9457 x64. Registry/API
`ProductName=Windows 10 Pro` is a compatibility label, not a Windows 10 pass.
Fresh hash-pinned regular-GIL CPython 3.14.8 dev venv; pypdf 6.19.0,
ReportLab 5.0.1/Pillow 12.3.0/charset-normalizer 3.5.2, isolated imports and
pip check passed. All 66 saved external shell-tool sizes/hashes match;
Pester 6.2.0/PSScriptAnalyzer 1.25.0. Actual hosts:
PS 5.1.26100.9444 and PS 7.6.5. The tools verifier's separately recorded interpreter
was the retained T01 venv; all current tests used the fresh T02 venv.

The final unrelated-CWD command was:

```powershell
& <M2-T02-root>\dev-venv\Scripts\python.exe -I -B <repo>\tests\run_tests.py `
  --layer full --tool-root <shell-tool-root> `
  --shell-path <WindowsPowerShell-exe> --shell-path <PowerShell7-exe> `
  --report <M2-T02-root>\full-final-02.json
```

Actual argv/CWD are aliased in the machine record, rather than inferred from this
presentation. Final run: `2026-10-09T09:47:40.312937+00:00` to `2026-10-09T09:49:46.354823+00:00`;
all 10 stage/native exits 0, 162 Python tests, nine Pester tests per host,
zero unit/Pester skips and no syntax/scaffold findings. The path stage separately
parses the new Paths helper on both hosts; inherited shell static checks have their
recorded narrower scope. Legacy analyzer observations remain non-acceptance results.

| Acceptance | Actual check and result |
|---|---|
| AC-036 | Three actual PS5.1/PS7/BAT runs with relative/literal special-character mixed `.PdF` input and observed read-only attribute. A wildcard decoy has different bytes/size. Metadata, source binding, output manifest and reopened page identities/content match the literal input. Eight actual PS rejects cover directories, providers, corrupt parser inputs and held-unreadable files. |
| AC-037 | Nine real title outputs match independent filename expectations; 14 focused filename units also pass reserved/control/Unicode/casefold and UTF-16 cases. |
| AC-038 | Each manual/Level1/Level2 writer produces 120 real one-page files, 001..120 in lexical chronology, complete identities and unique names. |
| AC-039 | A 175-unit existing base produces a final file exactly 259 UTF-16 units long. Preview entries remain unchanged; direct-API traps prove no replan/source reopen. Five failures cover impossible bound/unbound budgets, changed bound base, actual sharing-denied base and ENOSPC injection after a real first slice. No successful final output or source write; owned cleanup passes. |

Inherited manual/Level1/Level2/shared-plan/diagnostic/output regressions all pass.
They retain 22/4/4 targets, 8/5/6 aggregates, 250/150/150 seeded mode plans,
25/10/10 mode writers, shared preview/source controls and 300 seeded common plans,
six additional controlled success launchers (three concurrent), actual diagnostic
host/choice/stream cases and repeat/concurrent/failure/reparse ownership controls.
Sources, wildcard decoys, synthetic neighbors and earlier runs remain unchanged;
only authenticated known owned outputs/temp are removed. No private Documents
content is enumerated or processed.

Console setup has separate actual-host controlled trusted-source statement-block
receipts: four before cases continued to the engine boundary with exit 0; four after
cases stop with exit 1 and preserve existing members. This is **not AST extraction,
whole-launcher or ACL evidence**. Two actual native console reservation/held-cleanup
checks also pass. No extra whole-launcher console-setup pass is inferred.

Retained development receipts include the first runner NameError, initial Unicode
print failures, long-directory test-setup arithmetic failure, first PS5.1 read-only
policy-capture denial and later autoload/capture attempts. Corrected attempts have
new receipts; earlier passing targeted/full snapshots are distinct from final proof.
The old numeric policy capture has no host identity and matches PS7 only. Final
command-only reads identify both hosts; the path stage independently compares each
host's stored policies before/after and finds no change. Existing `LongPathsEnabled=1`
is observed, never enabled by this task. Encoded command tokens in public summaries
are represented by hashes; exact raw commands remain external.

## Review, synchronization and limits

Independent production/harness/raw evidence review found no remaining blocker.
No GitHub submitted approval or CI pass is claimed.
Branch: `codex/winbooksplit-v1-m2`; [draft continuation PR](https://github.com/PikkuJanne/WinBookSplit/pull/12).
C was normally pushed; fresh clean/live receipt: **SYNCED** at
`2026-10-09T09:52:45.121089+00:00`, local/live `a38d0e248cecd71513a2b06d6595c055e2712b95`.

At start, live PR11 was already merged at 09:08:41Z with merge/main 8a21e49;
a normal fetch/fast-forward reconciled main/M2 and the branch was normally pushed.
This observation does not establish cumulative M2 acceptance. Historical PR11 and
M2-T01 receipts remain historical; the new draft continues the remaining tasks.

Held sharing-denial controls are not ACL tests; ENOSPC is a simulation, not a filled
physical volume. Check/create and Windows close/publication intervals do not sandbox
an arbitrary same-account process. Unexpected owned members are retained. No human
Explorer/full interactive UX, UNC, arbitrary long-path, Windows 10/ARM, real Calibre,
PDF fidelity/features, CI, release-package or release pass is claimed. Ordinary
discovery, conversion, stream/cancel/remaining encoding and final CLI work stay in
their scheduled tasks. Compatibility exits 0/1/55 remain until M3.

Documentation-only E references C's observed evidence; its own synchronization
receipt stays outside its commit and in the PR/thread. No merge/tag/release here.
Next: **M2-T03 — Move ebook conversion into owned workspace**.
