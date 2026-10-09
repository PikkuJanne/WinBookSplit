# M2-T05 — Owned subprocesses and UTF-8 diagnostics

M2-T05; AC-046/047/048/049 **PASS on the tested worktree**. Date: 9 October 2026;
timezone Europe/Berlin. [Machine record](M2-T05-process.json) contains actual
commands, source/runtime identities, reduced observations and raw external hashes.
Base `c48a9dfe2a448599db97ac6bdc1d880ac3e0f0c1`, tree
`bdba74a4b7a5c054b4d0c103f3850e1fc32bdc50`. Final unrelated-CWD full gate has
13 stages: 192 Python units, 40 Pester tests per actual supported host, and all
current integration stages pass. The 76-path raw source digest is `7ae333509eaa80abf92529e5ed720c65e7a3cf307d92bb0c18e4fc0c43e689e7`.
All tested materialized bytes match clean implementation C
`877dd4bd7189d10d3dd9bbfd7efee8d42e631bae`, tree
`3ea8e45927c5ee669bae63baa906825cd1cc89de`. C was committed/pushed normally and
clean/live **SYNCED** at `2026-10-09T13:15:01.396075+00:00`.
No clean-C full rerun is claimed. E's own/final merge receipt remains external/PR/thread.

## Changes and reproduced behavior

The shipped `engine/WinBookSplit.Process.ps1` is a narrow direct-executable Windows
compatibility helper. It preserves BAT, console/drag-and-drop, Python/pypdf, optional
Calibre and the existing shared plans. It starts the exact executable suspended,
assigns an owned kill-on-close job before resuming, supplies an explicit inherited
handle list, and drains both byte streams concurrently. Normal completion includes
owned descendant stop and EOF; timeout/injected cancellation terminates only the
owned job and verifies stop. Readers, process/job/native handles and allocations
close in guarded cleanup. No process-name kill, shell evaluation, global policy
change, elevation, service, install or dependency substitution is introduced.

Both diagnostic tails retain at most 64 KiB with exact total/truncation counters.
Strict UTF-8 validation applies to the complete stream; a byte-budget cut drops a
leading incomplete character. A separate bounded frame collector preserves the
existing engine result even when it is larger than a human tail. Duplicate,
absent, malformed and over-budget frames reject. `HumanStdout` removes machine
frame fragments while retaining subsequent blank lines and no-newline human text.
The application logs bounded raw tails and `[PROCESS]` metadata; its unchanged
structured-result validator decides the outcome. Isolated Python explicitly uses
`-I -B -X utf8`; the child environment sets UTF-8 and bounded error handling.
`ProcessTimeout` is configurable; its default is at least 3600 seconds and at least
the conversion timeout plus 1800 seconds. Current 0/1/55 compatibility codes persist.

The runtime compatibility probe now delegates to this same owned helper rather
than parent-only termination. Existing conversion ownership, immutable source
capture and source/output cleanup rules remain. Source documents are immutable.

Before changes, actual extracted execution-function controls under both hosts
produced 4,000,698-byte flood logs and dropped blank stdout lines. A finite inherited
pipe control waited about 3.1 seconds without an engine deadline. Separate original
runtime probes returned around 0.31 seconds with an alive inherited child,
`DescendantsStopped=null` and incomplete streams; those authored children ended
naturally. The reports and exact baseline source snapshots remain outside Git.

## Actual commands and acceptance

Actual Windows 11 x64 workstation, 26H2 build 26300.9457; registry ProductName
"Windows 10 Pro" is a compatibility label. Not a clean OS/VM or Windows 10 pass.
Fresh supported regular GIL CPython 3.14.8 developer venv, pypdf 6.19.0,
ReportLab 5.0.1, Pillow 12.3.0 and charset-normalizer 3.5.2 were installed with
hashed requirements. Isolated imports and `pip check` passed; all 66 saved shell
tool sizes/hashes matched. Pester 6.2.0/PSScriptAnalyzer 1.25.0; separate actual
Windows PowerShell 5.1.26100.9444 and PowerShell 7.6.5. Portable Calibre 9.15.0 is
the already installed exact real converter; no new converter installation claim.

Final command, run from an unrelated owned directory (resolved argv/times/status
and stream hashes are in the machine record):

```powershell
& <M2-T05-root>\dev-venv\Scripts\python.exe -I -B <repo>\tests\run_tests.py `
  --layer full --tool-root <shell-tool-root> --calibre-path <Calibre915>\ebook-convert.exe `
  --shell-path <WindowsPowerShell-exe> --shell-path <PowerShell7-exe> `
  --report <M2-T05-root>\full-final-02.json
```

All 13 actual native stage statuses are zero. Python/Pester report no skips;
both hosts pass syntax and selected scaffold rules. The 43 legacy application
analyzer observations per host are recorded separately and remain limitations. Stored policies
match before/after. LongPathsEnabled=1 was observed, never changed.

| Acceptance | Actual evidence |
| --- | --- |
| AC-046 | Two actual native two-MiB dual floods with exact byte totals, 64 KiB tails and independent failure frames; PS5.1/PS7/unchanged BAT whole-application flood failures keep final stderr, nonzero status and bounded logs. |
| AC-047 | Exact fast-exit blank/no-newline streams and final stderr; independent 100 KiB frame, over-budget/duplicate/absent/malformed rejection. Timeout, injected cancellation, parent-exited inherited pipes and detached inherited-pipe children finish with job/tree/EOF evidence and independently stopped authored PIDs. |
| AC-048 | UTF-8 multibyte boundary tails and literal argument echoes under both hosts; three real Unicode path/bookmark/title/output/log PDFs preserve exact physical content, manifests and all original title data. Destination-aware shortened names match bound previews and independent path/title budgets. |
| AC-049 | Empty/quoted/trailing-backslash and shell/Python-looking argv remain literal. Percent/exclamation/bracket/space paths and marker-bearing hostile titles run through real PS/BAT; intended source/page identity is preserved and markers remain absent in wrapper, engine and book directories. |

The process stage requires 26 actual native controls plus nine actual application
cases. Six application failures substitute an authored engine only in an owned
copied application; three Unicode/title cases use the actual unchanged engine and
reopen all nine chapter PDFs. Failure controls are distinguished from split passes.
The unrelated fixture directly launches the actual ordinary base interpreter,
proves authored self-PID/executable readiness equals its retained `Popen` PID,
stays independently alive during all supervisor controls, then is independently
observed stopped after retained-handle termination. Only exact authenticated emitted
owned output/log paths are deleted; no private Documents entries are listed.

Twelve focused process Pester cases per host include already-cancelled no-launch
side effects and human tails after large JSON. Six synthetic independent-validator
units reject missing/contradictory records, counts, tails and exact case sets; these
units alone are not Windows acceptance. The strict validator also checks actual
full/direct/nested reports, original source/neighbor preservation, exact native
arguments, independent PIDs, complete physical content and ownership cleanup.

The final cumulative gate repeats eight real EPUB/genuine AZW3 conversion launches,
two captured-reader controls, eight invalid native-zero converters and BAT missing
converter refusal, plus all 33 runtime dependency/isolation/discovery cases. Earlier
manual/Level 1/Level 2/shared-plan/diagnostic/output/path guards and generated fixture
content remain unchanged. Converter/version/hash and per-stage raw hashes are in JSON.

## Defects, earlier attempts and limits

Initial actual empty-argument binding failed; `AllowEmptyString` fixed the API and
literal echo now passes. A pre-cancel review defect was corrected before final
tests. Defensive handle-transfer/dispose/diagnostic-copy changes were rechecked.
The initial scoped script policy refusal is recorded honestly; later successful
tests use only child process policy options, without changing stored policies.

The first Python transition ran 186 units and failed one source-unchanged assertion
while shared source edits were active. One development process attempt lost its
final output and produced no report: its outcome is UNKNOWN. A harness PID oracle
initially equated the venv shim with `os.getpid`; final native controls independently
query launched shim, authored parent and grandchild. Another development attempt
failed source-unchanged after concurrent edits. Precursor passing reports retain
their own older source identities and are not final acceptance.

First full run had 192 units, 40 Pester/host and its first 12 stages pass; the process
stage failed a hardcoded full-title filename oracle under the runner's longer TEMP
path. A separate nested process-only run reproduced it before correction. Production
safe destination-aware shortening was correct. The final test derives an actual-base
bound preview and independently measures complete path/title budgets and exact names.
Both ordinary and nested process runs then passed before the final full rerun.

Review also identified missing independent actual-PID/stop proof for the old
unrelated venv fixture. A scoped eight-second control found the actual child already
stopped after shim termination: no orphan/survivor bug was reproduced. The final
direct-base readiness/alive-to-stopped proof resolves the receipt ambiguity.

Raw/profile/installation paths, encoded commands and large stream/message bodies
remain external; public evidence uses aliases and hashes/sizes. No private document
was used or uploaded. Local independent review is separate from submitted GitHub
approval and remote CI. At the pre-C live audit PR14 was already merged at
2026-10-09T12:32:34Z; main/M2 matched the observed base, tags/releases were empty and
GitHub Actions had no runs. [PR15](https://github.com/PikkuJanne/WinBookSplit/pull/15) was observed OPEN and non-draft at
`2026-10-09T13:15:24.842434+00:00`, with C as its head. Independent local final
source/raw/public review reported no blocker before C, including all 41 then-listed
external raw hashes. C's normal commit/push/clean live receipt and PR creation are
separate additional external records. E/merge/final main receipts stay external or
in a later record; no future/self checkpoint SHA is embedded here.

Injected token cancellation does not certify human Ctrl+C/Explorer. The job/helper
is not a same-account sandbox or crash/power-loss durability guarantee. Earlier
source sharing/check/create/lock-close intervals remain. Windows10/ARM/UNC/arbitrary
long-path support, rendered PDF fidelity/features, expanded/remote/DRM cases,
CI/exact release-package and public v1.0.0 gates remain open. M3 owns standalone
version/CLI/menu/preview work. No tag or release is created by this evidence.
