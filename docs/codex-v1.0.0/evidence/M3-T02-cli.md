# M3-T02 — small non-interactive CLI

AC-053/054/055/056 PASS. The frozen fifteen-stage full gate passed
230 Python tests, 58 Pester tests per actual PS5.1/PS7 host, all inherited
regressions and 72 native CLI controls, including real Calibre EPUB/genuine AZW3.
See [machine record](M3-T02-cli.json) for source/receipt hashes, acceptance details,
runtime versions, exact redacted command and the observed C synchronization receipt.

## Source and environment

Implementation C `710ea8f92036b110918fb255f222814e6c4eeb4e`, tree `a6d0f02241c6c6b5f997301e90b295fe0214b06b`. All 91 tested raw paths match
committed C; tested digest `73872edccb4f380de3b97666a8084983b56c8bedd6d0565f4c7aa94b790265b7`. The full gate ran on materialized C
from an unrelated CWD. Evidence/status
files are excluded from the runtime digest.

Fresh regular x64 CPython 3.14.8 developer venv used hash-required binary-only
requirements-dev installation and pip check: pypdf 6.19.0, ReportLab 5.0.1,
Pillow 12.3.0, charset-normalizer 3.5.2. All 66 previously reviewed shell-tool
file hashes match; three extra metadata files are recorded separately. Actual hosts
were Windows PowerShell 5.1.26100.9444 and PowerShell 7.6.5, using Pester 6.2.0,
PSScriptAnalyzer 1.25.0 and real portable Calibre 9.15.0 on the current Windows
11 x64 workstation, build 26300.9457, DisplayVersion 26H2. The registry
retains the legacy ProductName Windows 10 Pro; Python reports Windows 11.
This was not a clean OS installation. OS metadata was reobserved after full04
on the same workstation and is recorded literally in the machine record.

```powershell
& <M3-T02-root>/dev-venv/Scripts/python.exe -I -B <repo>/tests/run_tests.py --layer full --tool-root <shell-tools> --shell-path <PS51> --shell-path <PS7> --calibre-path <Calibre915> --report <M3-T02-root>/full-final04.json
```

Native exit 0; all fifteen stage exits 0. Source, stored execution policies,
read-only synthetic inputs, neighbors and prior output remained unchanged.
Known owned suite cleanup passed. Zero skips, syntax errors or new scaffold
analyzer findings; legacy application findings remain separately reported.
Placeholders redact external/profile paths; hashed external receipts preserve
actual argv/CWD, streams and complete child reports.

## Behavior and acceptance

Explicit Auto requires level 1/2 and rejects start pages; Manual requires strict
ASCII decimal physical starts and rejects a bookmark level. Missing/contradictory
choices and unknown arguments return 2 before discovery/writes. Missing explicit
dependencies return 3; invalid physical ranges return 2; requested no-plan returns
5 without a prompt or fallback. The 72 controls run under both actual hosts with
DEVNULL stdin, validate native/final outcomes and independently reopen every
published physical page, chapter, manifest and retained full converted PDF.

NonInteractive suppresses prompts, pauses and cosmetic clearing. Independent
NoPause execution and Preview without application NonInteractive also pass.
Version alone works from an unrelated CWD in an entrypoint/shared-JSON-only copy
with Python/Calibre unavailable, prints `WinBookSplit 1.0.0-dev`, and exits 0.
Full comment-based help and documented scripted examples were exercised.
Native invocation/type-binding failures before script entry retain host status.

PDF preview publishes no files. EPUB/genuine AZW3 preview uses real temporary
conversion, discloses its generated source and proves safe cleanup. Preview shares
the immutable validated planner, reports physical inclusive ranges, names,
warnings and exact coverage, and returns preview_complete 0 with zero writer
count and null execution. Preview rejects converted-PDF retention and execution
results; execution rejects preview results.

Actual hostile bookmark titles and legal source filenames containing Unicode
line separators reproduced extra standalone console OUTCOME lines. Fixed both-host
receipts verify one final record, native 0, no output writes, unchanged source and
raw metadata. Human display escapes terminal/line controls; raw plans, JSON paths
and saved process streams retain their values. These reproduce/fix receipts and
meaningful Pester controls accompany the final full gate.

## Earlier attempts and limits

Earlier mutable-source Python/shell/CLI attempts and their actual failures are
retained, including four AST-loaded function failures fixed by explicit PlanOnly.
The first failed CLI child workspace was lost by outer cleanup; its 70 raw records
remain, and no retrospective cleanup pass is claimed. Corrected failed/timeout
retention preserves the workspace and its owning ancestor when descendant shutdown
is unproved. Earlier full gates
passed prior snapshots; the later display reproductions exposed their coverage gap.
An exact C comparison found 34 newline-only CRLF/LF differences in the checkout.
Full03 was interrupted before a final report; no pass or unknown partial-workspace
cleanup is claimed. The named paths were materialized from exact C blobs after
proving newline-only differences, and full04 tested all 91 committed raw paths.
An initial PS5.1 smoke was rejected under stored Restricted policy; the process-only
RemoteSigned retry passed. No global policy changed. See the machine record for
the original outcomes and receipt hashes.

Independent source, strict receipt and public evidence review passed. Current-build
human checks, including both BAT retries, are complete in M3-T01-human.md/json and
were not rerun or relabeled here. No CI, final extracted-package, new rendered
fidelity, GitHub approval or public-release pass is claimed. BAT bytes remain
unchanged. M3-T03 owns remaining launcher/menu behavior; v1.0.0 publication remains
subject to the release runbook.

C was normally pushed and clean/live SYNCED at `2026-10-09T17:55:24.447853+00:00`. E records
that observed result. Its own post-push receipt stays external/thread; this file
does not claim its containing commit was already synchronized.
