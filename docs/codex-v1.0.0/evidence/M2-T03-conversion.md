# M2-T03 — Owned ebook conversion

Task M2-T03; AC-040/041/042 **PASS**. Date: 9 October 2026;
user timezone Europe/Berlin. [Machine record](M2-T03-conversion.json) contains
actual commands, runtime/source identities, raw receipt hashes and outcomes.
Base `8a622dca7f37f50ed3c467ff98a26b8503d3ad1f`, tree `719edac70cf3168c0ec199b5b25d97062b191ae2`.
C `0128377099e47f51bfad7887f421b52999d522ad` is clean/live **SYNCED** at `2026-10-09T11:06:07.454548+00:00` on `codex/winbooksplit-v1-m2`.
Final unrelated-CWD eleven-stage gate passed 182 Python tests, nine Pester tests per actual PowerShell host, and all inherited regressions. Tested 61-path raw digest: `2dedae126955a871fff7812145a7bc78d38d258aa0c8e1e3918e87d225ffb4f9`. All raw tested bytes match clean C; no clean-C full rerun is claimed.

## Changes and reproduction

The existing BAT/console/Python/pypdf architecture and all logical planners remain.
PowerShell keeps the original ebook selected and passes CalibrePath, optional
KeepConvertedPdf and ConversionTimeout to the shipped engine. An explicit converter
or the three earlier standard locations are available; broader discovery and exact
runtime-version/import preflight remain M2-T04. PDF-only processing is converter-free.

A separate unique flat OutputRun workspace owns a parent-created, exclusively
registered WinBookSplit_Converted.pdf. The converter must preserve that identity.
The original ebook is held read-only and hashed before/after. After the tracked
child tree stops, exact membership and a fresh nonempty readable unencrypted PDF
with positive physical-page count are validated. Bytes and page-content hashes
are captured, then only known owned files are removed before preparation returns.
prepare_ebook uses the existing immutable reader-bound plan/preview/writer.
Original ebook identity stays separate from the generated PDF snapshot; filenames
use the original book stem. The snapshot remains valid after its workspace path
is removed. No chapter planner or source reopen runs during execution.

Optional retention writes exact captured bytes as WinBookSplit_Converted.pdf only
inside a successful chapter run. Its separate manifest filename/hash/size/page
count does not increase chapter outputs/written_count. Publication rechecks those
bytes. Default publishes no full PDF. PDF-only schemas and historical guards stay strict.

The converter starts suspended, joins an owned kill-on-close Windows job, then
its verified primary thread resumes. Both unbuffered streams drain concurrently
into 64 KiB tails with total counts/truncation flags. Prefixed logs avoid extra
protocol frames; bounded failure records preserve actual native status/context.
Timeout/controlled cancellation stops and waits only this owned tree before cleanup.
Unproved shutdown, unknown/replaced/reparse members preserve staging and the primary
cause in an explicit cleanup failure. No arbitrary PID kill or recursive adoption.

Before edits, exact source snapshots and real Calibre 9.15.0 reproduced a neighboring
authored one-page Book.pdf being overwritten by a three-page conversion on both
hosts. Stale neighboring output was accepted when a zero-exit child wrote nothing;
empty/corrupt/zero-page output was accepted. Missing output reported failure yet
returned zero. The final 14-control reproduction uses a genuine unrelated PDF.
The first authored AST selector failed before conversion; the next corrected
attempt used sentinel neighbor bytes. These remain separate receipts. Reviewer
backend/placeholder probes are not whole-launcher evidence.

## Actual commands and outcomes

Actual Windows 11 Pro 26H2 build 26300.9457 x64 workstation, not a clean OS/VM.
Registry ProductName Windows 10 Pro is a compatibility label, not a Windows 10 pass.
Fresh hash-pinned regular-GIL CPython 3.14.8 dev venv, pypdf 6.19.0, ReportLab 5.0.1,
Pillow 12.3.0 and charset-normalizer 3.5.2: isolated imports and pip check pass.
All 66 saved shell-tool sizes/hashes match; Pester 6.2.0/PSScriptAnalyzer 1.25.0.
Actual separate hosts: PS5.1.26100.9444 and PS7.6.5. Existing explicitly installed
portable Calibre is actually 9.15.0; converter executable SHA256
f46a01c9b8cd392e1190ce2e50a2a5a48b326c8c3eab670bbcd8b298bede501d.
No full Calibre installation-tree reproducibility claim.

Final command, from unrelated CWD (actual argv/timestamps are aliased in JSON):

```powershell
& <M2-T03-root>\dev-venv\Scripts\python.exe -I -B <repo>\tests\run_tests.py `
  --layer full --tool-root <shell-tool-root> --calibre-path <Calibre915>\ebook-convert.exe `
  --shell-path <WindowsPowerShell-exe> --shell-path <PowerShell7-exe> `
  --report <M2-T03-root>\full-final-02.json
```

All eleven native stage statuses are zero; no unit/Pester skips, syntax/scaffold
findings or fabricated CI passes. Stored policies before/after are unchanged;
children request process-only RemoteSigned. Existing LongPathsEnabled=1 was
observed, never enabled. Sources/neighbors/prior outputs and known cleanup pass.

| Acceptance | Actual result |
|---|---|
| AC-040 | Eight real PS5.1/PS7 format/retention runs and two direct API controls preserve read-only original EPUB/genuine AZW3 and neighboring same-name genuine PDF hashes/attributes. Generated target is owned workspace only. |
| AC-041 | Eight actual whole-PS-script runs: no converter write, empty, corrupt and zero-page PDF on each host. Actual parent observes native zero and final return control; engine/launcher reject nonzero, write zero chapters and clean known owned files. |
| AC-042 | EPUB/AZW3 × default/keep × both hosts using real Calibre. Default publishes no full PDF; keep independently hashes/reopens exact captured full PDF in completed run only. Three EPUB/four AZW3 physical pages, including contents, survive exactly once. |

Independent real reference conversions match exact page-content vectors. Default
launcher vectors are engine metadata crosschecked with references and reopened
slices, not independent observation of deleted workspace bytes. Direct API controls
independently inspect captured bytes/content and immutable preview/no-replan parity;
keep cases independently reopen/hash actual full PDFs. Three original chapter
markers appear once. Original fixtures have no remote resources and are MIT test
material; real Calibre generates AZW3 rather than renaming an EPUB.

Seventeen focused source-stable fake-native units pass: literal hostile arguments,
large dual streams/final no-newline UTF-8 tails, invalid/encrypted zero-exit output,
nonzero despite valid PDF, start/timeout/cancel proof, unknown/replaced members,
cleanup/close failures and exact retained-snapshot/tamper publication checks.
Separate actual parent/grandchild probes observe job count five to zero, parent
status one and exact cleanup/read-only source/neighbor preservation on timeout
and injected KeyboardInterrupt. Extra member images were not enumerated. These
are scoped helper controls, not human Ctrl+C or actual Calibre timeout evidence.
Unchanged BAT separately returns one for missing standard-location Calibre;
successful portable converter discovery remains M2-T04.

Development receipts retain initial native probe count/breakaway assumptions,
reproduced BufferedReader shutdown hang and fix, startup-cancel stop-proof fix,
partial external receipt serialization, collector typo and two policy setup errors.
A first runner fixture lacked the conversion-only manifest adapter. First targeted
acceptance exposed a test helper truthiness error: nonempty ContentStream dictionaries
evaluated false and produced empty hashes. Scoped diagnostics show reference/captured
streams actually match. The fixed is-not-None check has a meaningful regression.
Earlier failed/partial/passing attempts remain separate hashes, never final proof.

The first full gate passed ten stages and failed the existing diagnostic quoting
guard: a new format check parsed a synthetic raw argument as an OS path before
marshaling it. The one-line correction uses a fixed suffix data match; main path
validation and the existing strict quoting/stream test remain unchanged. The
second full gate is the final passing source, with its own digest and receipt.

## Review and checkpoint

Independent source/native/harness/raw/public review found no remaining blocker;
receipt source hashes and scopes are recorded. No submitted GitHub approval or CI claim.
[Draft continuation PR](https://github.com/PikkuJanne/WinBookSplit/pull/13). C `0128377099e47f51bfad7887f421b52999d522ad` is clean/live **SYNCED** at `2026-10-09T11:06:07.454548+00:00` on `codex/winbooksplit-v1-m2`.
E references this already observed C receipt. E's own push/live receipt remains
external/PR/thread without a self-hash loop. At start PR12 was already merged at
10:19:22Z with main/merge 8a622dca. Normal fetch/fast-forward reconciled main/M2;
no cumulative M2 acceptance inferred. No milestone merge/tag/release here.
Only intended files are staged; no private book, generated PDF/ebook binary,
raw converter log, private profile path or opaque encoded command is published.

Remaining: M2-T04 preflight, M2-T05 process/encoding/cumulative review, M3 preview/
confirmation/CLI, rendered fidelity/features, human Explorer, expanded/remote-resource/
DRM rejection cases, CI, package and public v1.0.0. Calibre may use external temporary
data; we do not recursively adopt it. Same-account check/create/lock-close intervals
are not sandbox guarantees; no crash durability/arbitrary long-path/UNC claim.
Current 0/1/55 compatibility remains. Next exact task: M2-T04 only.
