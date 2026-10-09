# M2-T01 — Isolated validated output transactions

Task: M2-T01; AC-032 through AC-035 **PASSED**.
Author: Codex; independent local source/harness/native/raw-evidence review passed.
Full gate observation: `2026-10-09T08:27:43.889309+00:00`; user timezone Europe/Berlin.
Base: `406622f2c27a869f202701c550f91ec4e9adc7e9`.
Implementation C: `a58ce43866331da42e32f80521a24c4893ab58f5`; tree `5a0ced7e317a9bd7f751fe336136ee42fc15209b`.
Tested actual 51-path digest: `eb9c4dee60c60e022bd5fafbacb1acad6894a2312f7240bfe3fd7932d1bb1e10`.
Full raw SHA-256: `b31bdbd01222a886d8c00a27c32e6a6c743f79106f848ac4e63e27948b615a53`.
[Machine evidence](M2-T01-output.json) records exact observed commands, hashes,
cases, errors, prior receipts and limits with local paths redacted.

## Behavior and ownership

The existing manual/Level 1/Level 2 planners, immutable captured-source plan and
pypdf page writer remain. The output argument is now an existing base. Each run
atomically reserves a random hidden sibling stage, registers each exclusive
file immediately after opening, reopens every positive planned slice, and
checks complete manifest membership, counts, sizes and hashes. Native
same-volume handle rename publishes a unique child without replacement;
collision retries preserve the competing folder. PS passes actual Documents,
shows the explicit completed folder, and keeps console logs in their own
exclusive marked folder. BAT bytes remain unchanged.

Cleanup authenticates the created directory/flat-file identities and marker,
rejects unknown/reparse/changed members, then deletes only held owned objects.
It never recursively traverses a target or uses paths from an edited manifest.
Partial write/reopen/count/manifest/promotion failures report nonzero, remove
only known owned partials and retain bounded separately marked failure records.
Uncertain stages are preserved and explicitly named. A post-publication close
error retains the completed result and adds an explicit warning; all handles
still receive close attempts. Earlier stages are never automatically deleted.

Before changing production, six real synthetic-PDF probes reproduced partial
second-write files, failed retries/repeats and zero/wrong-count/unreadable
outputs incorrectly accepted as success. Later review reproduced and fixed the
post-publication false failure and observed initial owner-marker reparse escape.
Earlier failed/pass receipts remain distinct from final acceptance.

## Actual execution

Fresh hash-required regular GIL x64 CPython 3.14.8/pypdf 6.19.0 dev venv;
ReportLab 5.0.1/Pillow 12.3.0/charset-normalizer 3.5.2. Exact imports and pip check
passed. Actual Windows build 26300.9457 x64, DisplayVersion26H2 (Python reports
Windows11; registry compatibility ProductName is Windows10Pro).
PS5.1.26100.9444 and PS7.6.5; Pester6.2.0/PSScriptAnalyzer1.25.0;
all66 recorded external tool sizes/hashes passed. Stored policies unchanged.

| Executed command/scope | Actual result |
| --- | --- |
| `$DevPython -I -B $Repo\tests\run_tests.py --layer full --tool-root $ToolRoot --shell-path $PS51 --shell-path $PS7 --report $ExternalRoot\full-final-04.json`, unrelated external CWD | Exit0; nine stages, 146 Python tests/no skips; nine Pester tests per actual host; zero syntax/scaffold findings. |
| Ninth output stage | Four repeat/same-basename actual CLI runs; four actual barrier/PID-distinct identical-clock runs; five typed failures; eight ownership cases including real junctions. Every successful manifest/slice reopens exact physical IDs/content/order; prior/source/neighbor bytes preserved. |
| Focused final output units | 26 passed, including native collision/no-clobber, real junctions, partial writes, zero/wrong-count/corrupt outputs, bounded diagnostics, metadata/byte tamper, cancellation and pre/post-commit close boundaries. |
| Current full success launchers | Six actual PS5.1/PS7/BAT probes, three concurrent; explicit manifest/log paths, complete PDF bytes/IDs and held identity cleanup. |
| `$DevPython -I -B $ExternalRoot\probe_entrypoint_failures.py $ExternalRoot\entrypoint-failures-final3.json` | 12 actual controlled PS5.1/PS7/BAT cancellation/invalid-retry/corrupt/quoted-manual flows; expected55/1 codes, no Done/chapters, exact owned log cleanup and preserved source/neighbors. |
| Extra current-cleanup native probe | Four passed; ordinary owned removal, actual file replacement rejection with foreign bytes preserved, unexpected-member and wrong-marker preservation. |
| Native sharing/reparse probes | Earlier24 and13 checks cover retained-handle no-replace promotion, identity-preserving forward/reverse base sharing, actual rename denial, Unicode and handle cleanup. Later metadata probes distinguish compatible empty-directory mutation from observed pre-write rejection. Each receipt identifies its own source bytes. |
| Inherited current routes | Manual22 targets/eight extras/250 plans/25 writers; Level1 four targets/five aggregates/150 plans/10 writers; Level2 four targets/six aggregates/150 plans/10 writers; shared plan300 seeds/all-mode/source-snapshot parity; diagnostics23 records, both-host choices/malformed records/200K streams. |

Final raw bytes matched clean C independently; no clean-C full rerun is claimed.
Immutable original/extraction guards, fixtures/oracles and historical evidence
are unchanged. Legacy analyzer observations and redirected PS7 RawUI/PS5.1
progress stderr remain in raw receipts; selected static success does not turn
those into full application acceptance. No submitted GitHub approval/CI pass,
human Explorer, Calibre conversion, rendering, crash/power-loss durability or
extracted-package pass is claimed.

## Native limits and review

The API uses [Microsoft's no-replace rename structure](https://learn.microsoft.com/en-us/windows/win32/api/winbase/ns-winbase-file_rename_info)
and [handle sharing controls](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-createfilew).
Actual compatibility probes, rather than sharing assumptions, determine the
tested behavior. Compatible FILE_WRITE_ATTRIBUTES can mutate an empty
directory despite data-sharing locks. Checks reject an already mutated stage
before owner creation; check/create is not atomic. Windows requires releasing
child-file locks before final byte/identity checks and directory publication.
Neither interval promises a sandbox against an arbitrary same-account process.
No source/neighbor overwrite or target-following cleanup was observed.

The Windows helper's author is distinguished from its independent root/unit
reviewers. Independent source/harness/raw review found no remaining blocker in
AC-032..035. Native and public evidence are scoped workstation observations.

## Git and next task

Branch: `codex/winbooksplit-v1-m2`; [draft PR](https://github.com/PikkuJanne/WinBookSplit/pull/11).
C normal push/live receipt: **SYNCED** at `2026-10-09T08:31:04.317614+00:00`.
M2 remains in progress; no milestone merge/tag/release was created.
Documentation-only E receives its own external/thread/PR receipt after push.

Next: **M2-T02 — Harden literal paths and output filenames**. Read fresh status,
Git/live state and task/specs; keep transactional ownership and plan invariants.
No private inputs/raw logs/binaries were staged or uploaded.
