# M1-T02 — Strict manual starts and complete page coverage

Observed: 9 October 2026, Europe/Berlin. Final full run began at
`2026-10-09T05:37:37.200646+02:00` (03:37:37 UTC).
Author: Codex primary; independent source/harness/privacy review agent;
separate manual-unit and runner-transition agents.
Base: `8c7c22fa8eda3c61c23b63d429d0b6abbd40432a`, tree
`61ddb79dff14e8c41d47a9fa346430ffda85d542`.
Implementation C: `ced0d96f0540d96adc2fdacc4e1ff31fa25be7b8`, tree
`05131b1df312a553e3a2794b7124445d8a37d7d2`.
Tested actual-path digest:
`27cf0208484d831d5e818e64b1f2478a5066b96f436e2fb9da53f8029a3fe1ad`.
The [machine record](M1-T02-manual.json) holds the source map, complete manual
observations/page identities/output hashes, raw receipt hashes and earlier failures.

## Behavior and scope

The direct pre-fix extraction run reproduced all 14 original observations:
`1` wrote zero files/exit 0, `1,4,7` lost pages 1–3, and mixed invalid tokens
were filtered. The shipped engine now validates every comma-separated ASCII
decimal token and 1..N bound before normalizing or writing. It rejects empty
tokens, signs, ranges, decimals, letters, non-ASCII digits and out-of-range starts
entirely. Zero-page manual input fails as `invalid_document`; other invalid
manual requests report `invalid_start_pages` and exit 1 before writer calls.

Valid requests sort/dedupe with notices and include physical page 1 with a notice
when inserted. `1` produces one valid whole-document PDF and the explicit
one-section notice. `1,4,7` and `4,7` both produce `[0,3),[3,6),[6,10)`;
`10` also preserves the final one-page slice. Lexical decimal bounds before
integer conversion handle huge invalid numbers and valid thousands-of-leading-
zero tokens without Python's integer digit-limit failure. The manual branch
executes exactly its constructed ranges; existing bookmark/writer/launcher
behavior is otherwise preserved.

The current full route deliberately uses corrected manual targets, with strict
child receipt completeness/preservation checks. Historical AC-011 AST evidence
reads immutable M1-T01 C `88c2149`; safe import still checks the current engine.
Original/extraction harnesses, baseline raw-byte guards, oracles, fixture
generator and historical evidence remain unchanged. Their explicit legacy
diagnostics refuse changed source as expected. Three current bookmark cases
still match the original complete known-bad observations, never repaired
bookmark acceptance. The two launchers, original README/license and all artwork
match their M1-T01 raw blobs; the six unedited originals retain baseline bytes.

## Actual environment, commands and outcomes

Windows 11 Pro 26H2 x64, build 26300.9457, existing workstation; no clean-OS
claim. Fresh external regular GIL CPython 3.14.8 x64 dev venv; pypdf 6.19.0,
ReportLab 5.0.1, Pillow 12.3.0, charset-normalizer 3.5.2. Hash-required binary-only
install, pip check, exact versions/GIL/x64 and all venv import-origin assertions
passed. All 66 external shell-tool file hashes/sizes matched the prior audited
manifest. Actual hosts: PS5.1.26100.9444 and PS7.6.5; Pester 6.2.0 and
PSScriptAnalyzer 1.25.0 imported from absolute pinned manifests. Children used
host-owned module paths and process-only RemoteSigned; stored policies remained
unchanged. No elevation/global installation or policy change.

Variables are symbolic redactions of actual absolute external paths. `$Repo`
is the trusted checkout, `$DevPython` the fresh venv, `$ToolRoot` the reverified
modules, `$PS51`/`$PS7` the discovered actual hosts, and `$M1T02ExternalRoot`
the new evidence directory. Raw reports remain external.

| Actual command/check | Outcome |
| --- | --- |
| `$SetupPython -I -m venv <new-external-dev-venv>`; `$DevPython -I -m pip install --require-hashes --only-binary=:all: --index-url https://pypi.org/simple -r $Repo\requirements-dev.txt`; `pip check`; version/import-origin assertions | All exit 0; fresh supported dependency setup |
| `$DevPython -I -B $Repo\tests\extraction\characterize_extraction.py --report $M1T02ExternalRoot\before-direct.json` before engine editing | Exit 0, all 14 original/current known-bad observations equal; relevant source/input/neighbor preservation and safe import |
| `$DevPython -I -B $Repo\tests\manual\characterize_manual.py --report $M1T02ExternalRoot\manual-first.json` from unrelated external CWD | Exit 0: 22 manual oracles, 8 extra CLI cases, 250 seeded plans, 25 real writer samples, 3 unchanged bookmark cases; targeted earlier harness bytes |
| `$DevPython -I -B $Repo\tests\run_tests.py --layer full --tool-root $ToolRoot --shell-path $PS51 --shell-path $PS7 --report $M1T02ExternalRoot\full-after-hostcase.json` from unrelated external CWD | Exit 0: 36 Python tests, 9 Pester tests per host, zero skips, all current manual/historical bookmark/import checks and 6 actual entrypoint probes |
| Both-host syntax/new-scaffold analysis | Zero findings/errors; 60 inherited application analyzer observations per host are separately recorded |
| Explicit legacy extraction diagnostic after behavior fix | Expected exit 1 / no report: processing AST differs; historical guards remain intact |
| `$DevPython -B $Repo\tools\codex-handoff\validate_plan.py --plan-root $Repo\docs\codex-v1.0.0`; `git -c core.whitespace=cr-at-eol diff --check` and staged equivalent | VALID/exit 0 and whitespace exit 0; structural consistency only |

Final full raw SHA-256:
`a6e14deb453a6ed61d1740f3c8155cebab08e5bf74ef9ffbb5c3727de142d264`.
Current manual child SHA-256:
`afd5a1dd02a4e328fe9058a490fe5f667c873770b83a46897b945c6483e26955`.

| Acceptance | Executed checks | Result |
| --- | --- | --- |
| AC-013 | MAN-01/06, one-section notice and real valid output | PASSED |
| AC-014 | MAN-02/03, full per-slice and flattened page identities | PASSED |
| AC-015 | MAN-04, sorted/deduplicated leading-zero starts and notices | PASSED |
| AC-016 | MAN-07/08/14/15/16 plus comma-only; no output/writer call | PASSED |
| AC-017 | MAN-09/10/11/17–20 plus fullwidth/embedded-space/late-invalid tokens | PASSED |
| AC-018 | MAN-05/12/13/21 plus standalone zero/5000-digit integer | PASSED |

MAN-22 additionally rejects a real zero-page PDF. Seed 20261009 checks 250 valid
partitions of 1..120 pages; 25 samples write/reopen ten-page marked fixture
slices. Each slice's full identity list and the flattened 1..N order are checked;
page totals alone never establish coverage. Invalid direct calls verify
`write_slice` is never reached, nonzero exits/diagnostics and source/neighbor
preservation. No requested acceptance ID is skipped.

Six bounded existing actual PS5.1/PS7/batch probes, including three simultaneous
launches, matched exact reference filenames/PDF hashes/page identities. Manual
probes use `4,7`; PS7 probes use unchanged known-bad BM-03. Each unrelated CWD
contains an unexecuted counterfeit engine. Inputs, synthetic Documents neighbors
and shared TEMP engine-name sentinel remained unchanged. Only exclusively
GUID/marker-owned Documents outputs were inspected/removed; private Documents
entries were not enumerated/read. Uncertain process/ownership states preserve
artifacts. The two PS7 piped probes still emitted the inherited RawUI cursor
invalid-handle stderr; selected engine resolution acceptance does not certify
console/error UX.

## Earlier failures and limits

The first central before-fix run rejected concurrent test/documentation edits;
its extraction child passed, followed by the stable direct 14-case reproduction.
An initial full command used an absent example PS7 path; preflight rejected it
without a report, then the actual host path was discovered. The first complete
full run passed 35 Python tests/9 Pester per host and manual child exit 0, but
the new receipt checker rejected `C:\Windows` versus `C:\WINDOWS`. A targeted
regression reproduced that error. Native WindowsPath equality now accepts
casing/slash variants, still rejects unrequested/duplicate hosts, and the final
36-test full rerun passed. Earlier receipts retain their own distinct hashes.
A checkpoint redaction assertion initially stopped before writing because a
host module-directory path retained the username; it was redacted and rechecked
before evidence publication.

Bookmark/front-matter/parent-boundary fixes remain M1-T03/M1-T04. Overwrite/output
transaction safety, all launcher/error/discovery paths, human Explorer, Calibre
conversion, rendering fidelity, CI, extracted release package, tag/release and
public downloads remain NOT RUN/open. Other platforms/Python versions stay
unclaimed. No private document was processed or committed; generated PDFs,
runtime/tool binaries, raw paths and reports remain outside Git.

## Git and next task

Start feature `3d2463c` freshly matched GitHub at 03:27:44 UTC. Live PR #5 was
already merged at 03:26:06 UTC to `8c7c22f`; normal fetch/fast-forward incorporated
that actual main history before changes. Branch remains `codex/winbooksplit-v1-m1`;
new [draft PR #6](https://github.com/PikkuJanne/WinBookSplit/pull/6) continues M1.
Keep it draft until M1-T06; the earlier PR merge does not waive remaining tasks.

Every tested actual-path byte matched clean C. C was normally pushed; clean
local HEAD equaled fresh live feature C at `2026-10-09T03:39:03.607417+00:00`
(05:39:03 Berlin). Tests ran on the pre-commit worktree; no clean-C full rerun
is claimed. Independent source/harness/evidence/privacy review found no remaining
M1-T02 blocker. No submitted GitHub review or CI pass is claimed. Live main and
feature were unprotected; rulesets/workflows/runs/tags/releases were empty;
settings were untouched. No merge/tag/release/history rewrite/stash/deletion.

This following evidence/status checkpoint E references C. Its own final normal
push/live receipt belongs in the thread/PR, with no self-referential SHA or
future synchronization claim here. Next: **M1-T03 — Normalize bookmarks and
fix Level 1 coverage**, after fresh checkpoint/live verification.
