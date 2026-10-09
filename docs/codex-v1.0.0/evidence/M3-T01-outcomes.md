# M3-T01 — truthful outcomes and owned cancellation

AC-050, AC-051 and AC-052 PASS on the frozen tested worktree. The fourteen-stage
full gate passed 217 Python tests, 49 Pester tests per actual supported
PowerShell host, all inherited regression layers and 51 actual outcome applications.
See [machine record](M3-T01-outcomes.json) for raw-source hashes, exact redacted
stage commands, report hashes, every case and the historical C synchronization receipt.

## Source, environment and commands

Base/main at thread start: 85916466ef45f25243806fd6d380bf9c3dcee5d4; actual clean
main already included merged PR15. No baseline reset. Tested pre-commit raw digest:
ac6c5b910891ef8a86f078db43b2677f66cd515022dd17251ed10bbc9b3d87a0 (85 paths). All tested raw paths match materialized clean
C b8c6ed0f9e1153e20ba9b265615c6c3cab8d1a7e; its tree is 0347790d7a9a4286450d62e940e6353ae068799d.
No clean-C full rerun is claimed. Evidence/status files are excluded from that digest.

Actual Windows build 26300.9457, registry DisplayVersion 26H2/Professional Client,
LongPathsEnabled 1; the registry retains ProductName Windows 10 Pro while Python reports
Windows 11. Fresh regular x64 CPython 3.14.8 developer venv, isolated hash-required
requirements-dev install and pip check passed: pypdf 6.19.0, reportlab 5.0.1,
Pillow 12.3.0, charset-normalizer 3.5.2. All 66 existing isolated shell-tool file hashes
match historical pinned evidence. Actual PS5.1.26100.9444 and PS7.6.5, Pester 6.2.0,
PSScriptAnalyzer 1.25.0 and real portable Calibre 9.15.0 were used.

The final command ran from a new owned unrelated CWD, with native exit 0:

```powershell
& <M3-T01-root>/dev-venv/Scripts/python.exe -I -B <repo>/tests/run_tests.py --layer full --tool-root <shell-tool-root> --shell-path C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe --shell-path <bundled-runtime>/dependencies/native/powershell/pwsh.exe --calibre-path <Calibre915-root>/ebook-convert.exe --report <M3-T01-root>/full-final04.json
```

Placeholders redact profile/external paths; the hashed external raw report retains
the actual argv/CWD, bootstrap/environment, streams and full child reports. The
public record is a summary rather than raw console output. Source stayed unchanged,
no tests were skipped, new scaffold analysis/syntax checks passed, stored execution
policies stayed unchanged, and final owned suite cleanup/source sentinel checks passed.
Legacy application analyzer findings remain separately reported, not an acceptance pass.

## Behavior and acceptance

Both runtimes load engine/WinBookSplit.Outcomes.json: success 0, invalid arguments 2,
runtime/dependency 3, conversion 4, unavailable outline 5, processing/output 6,
reserved unsupported 7, interruption 130. Final fallback outcome and native BAT exit
agree. Explicit fallback cancellation returns 130; final success waits for console
log flush/close. Completed engine output followed by a close/log error is reported
incomplete 6 with its validated location retained. A secondary log error preserves
the primary failure/cancel. The log's OPERATION-OUTCOME precedes finalization;
stdout OUTCOME is authoritative. Redirected PS7 skips cosmetic clearing and BAT
uses process-local UTF-8 output, preserving Unicode paths in the final record.

The 51 application controls are seventeen kinds under PS5.1/PS7/BAT: original failure,
successful/failed/cancelled fallback, startup dependency, actual native converter exit 17,
real-engine write/manifest/publish/foreign-member cleanup faults, engine and conversion
token cancellation/timeouts, post-publication console finalization, and both actual OS
Ctrl+C paths. Faults are limited to owned copied applications; original/controlled/saved
engine bytes and native fixture compiler source are hash-bound. BAT copy bytes match
shipped BAT; only copied PS defaults provide its controlled input-only dependency paths.

Owned job parent/descendant/pipe-EOF proofs and authored operation PIDs agree with
shutdown. The independently identified unrelated Python process survives all operations
and only its retained fixture handle stops it afterward. Synthetic source/neighbor/prior
output identity, hashes and attributes remain equal before/after operation and cleanup.
Success and incomplete publications reopen all three physical pages/content, names,
ranges, sizes/hashes and complete manifests. Foreign-member cleanup refuses adoption;
exact held-identity harness recovery then removes only its authored members. Outer
interruption retains marked staging because the console has no engine file ledger.
Unsafe/unproved stop never authorizes outer recursive cleanup.

Actual hidden private-console CTRL_C_EVENT checks use a rooted native handler throughout
supervision/shutdown/drains; it consumes only Ctrl+C/Break. Any observed CMD batch prompt
is answered automatically only after final exit 130 and owned PIDs stopped; those decisions are
in the case records. These are automated OS checks. Human Ctrl+C/Explorer checks NOT RUN.
Real Calibre EPUB/genuine AZW3 conversion passed in inherited full layers; the native
fault converter alone does not certify Calibre. Process regressions also retain 26 actual
native controls plus 9 actual applications. Unit/Pester authored close/transport failures
are distinguished from naturally failing native handles.

## Reproduction and earlier attempts

Pre-change Python controls observed no-plan 55, invalid manual 1, writer 1 and actual
KeyboardInterrupt 1. An authored real console-finalization tail reproduced Done/native exit 0
despite a dispose exception. Existing fallback already used the last returned result
and conditional Done; no invented baseline fallback bug is claimed. Whole-script OS
Ctrl+C subsequently exposed PowerShell pipeline cancellation despite the managed event
callback, after narrow helper controls had passed. Native handler controls and the final
whole applications verify the correction.

Earlier receipts remain distinct outside Git and are hashed in the machine record.
Early unit/shell expectations and PS7 JSON integer handling failed before correction.
Fourth outcome attempt exposed the actual whole-script signal defect. Fifth's child
partial report was deleted by the old outer runner: its exact case/count is UNKNOWN,
not a proved BAT defect or pass. Sixth retains 16 completed cases and an unproved later
signal after a controller ignore-flag inheritance fixture error. Seventh retains 17
completed cases before redirected PS7 header failure. Eighth retains 35 completed cases
before BAT Unicode outcome corruption. The controller restores its ignore flag with
checked success; missing/unproved child reports are retained as failed_evidence rather
than validated evidence. A separate raw-retention regression verifies this distinction.

First two full attempts were superseded by source fixes: source-change
guards refused evolving integration runs, and 299 legacy rejection subtests caught raw
failed reports being mislabeled validated evidence. Neither is an acceptance pass.
Frozen full-final03 passed 13 stages but failed the inherited process log parser: it counted the new provisional outcome footer as stderr, crossing the unchanged byte bound. The parser now bounds the raw stream before that separately validated footer. Final frozen full-final04 alone supplies the full-gate claim. Independent receipt review
also reproduced accepted contradictions, then added exact stdout/log/outcome, adapter
hash, staging and PID bindings with rejection tests. Original baseline/extraction guards
and physical-page/ownership negatives remain intact.

## Checkpoint, review and limits

Root and independent agents reviewed application sources, test expectations, strict
receipts and public evidence; no blocker remained. C was pushed normally and clean HEAD
equaled the freshly queried live M3 branch at 2026-10-09T14:49:29.758828+00:00. E references this
already observed receipt. E's own final receipt belongs outside its commit and in the
thread/PR, after its push. No GitHub submitted approval or CI pass is claimed; fresh
main was unprotected, Actions runs 0, tags/releases empty. No tag or interim release.

Next exact task: M3-T02. Version/CLI binding and M3-T03 menus/preview/no-input UX remain
open, as do human/Explorer, broader PDF fidelity, CI, extracted-package and public
v1.0.0 gates. Future packaging must ship Outcomes.json and Process.ps1 with the engine.
Unknown/replaced/reparse/unproved-stop objects remain retained; existing same-account
race/crash limitations are not a sandbox. No private textbook or document was processed
or published, no global execution policy/settings changed, and only scoped files were staged.
The external footer-diagnosis capture also had a filename-copy collision: its original
trial temporary directory was removed, but an unknown authored BAT-fast-tail console
path remains retained. No private Documents scan or deletion was attempted. The corrected
driver uses unique log IDs and finally-based existing exact-ledger cleanup; this was a
test-driver defect, not a production output/cleanup failure.
