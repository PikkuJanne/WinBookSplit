# Next session

Next task: **M1-T02 — Fix manual start-page parsing**.
M0-T01 through M0-T04 and M1-T01 are done. M1-T01 implementation C
`88c2149b3b3034fbd0d7ef23c2382f4b01648e4c` was normally pushed and clean local
HEAD matched the fresh live feature branch at
`2026-10-08T19:08:08.466690+00:00` (21:08:08 Europe/Berlin).
[Draft PR #5](https://github.com/PikkuJanne/WinBookSplit/pull/5) continues M1;
keep it draft until cumulative review/merge at M1-T06. This following
documentation-only checkpoint receives its own final normal push/live receipt
in the thread/PR. Recheck actual current equality next session; historical
receipts and recorded done status do not prove current synchronization.

Workspace: `D:\projects\WinBookSplit-main`. Active branch/upstream:
`codex/winbooksplit-v1-m1` / `origin/codex/winbooksplit-v1-m1`.
Origin fetch/push: `https://github.com/PikkuJanne/WinBookSplit.git`.
Main was `0de84f367f9bd5ddfa3f408a9c29505d7a39633f`, freshly synchronized
at M1 start and still live at the read-only 19:00:29 UTC audit. Retain M0 branch
and closed PRs; do not reset to historical baselines or start another milestone
branch. Preserve unrelated edits/history and reconcile actual new remote work.

Read AGENTS.md, STATUS.md, TASKS.json, SCOPE_AND_DECISIONS.md,
GITHUB_WORKFLOW.md, TESTING.md, SUPPORT_AND_SETUP.md,
evidence/M1-T01-extraction.md, evidence/M1-T01-extraction.json,
tasks/M1-T02.md, specs/SPLIT_CONTRACT.md and PLAN_ORACLES.json.
Start with actual inspection:

```powershell
git status --short --branch
git remote -v
git worktree list
git rev-parse --abbrev-ref --symbolic-full-name '@{upstream}'
git rev-parse '@{upstream}'
python -B .\tools\codex-handoff\check_sync.py --repo .
python -B .\tools\codex-handoff\validate_plan.py --plan-root .\docs\codex-v1.0.0
```

The stdlib handoff helpers use their authored sibling import: do not add `-I`
to those two commands. Use the supported explicit Python `-I -B` for application
tests. Inspect authentication, live feature/main refs, PR state/protection/
required checks and tags/releases separately. Repair unknown/unsynced evidence
first. No auto-stash, force push, reset or deletion.

## Verified extraction to preserve

`engine/winbooksplit_engine.py` ships the original processing functions with
identical ASTs and a guarded CLI entry; import has no processing/configuration
side effects. PowerShell resolves it from `$PSScriptRoot`; no shared generated
engine file exists. Batch, original README/license and artwork remain raw-byte
identical. The explicit extraction route compares immutable Git originals,
never weakening expected_original.json or the original harness's source guards.
The historical `--layer baseline` correctly fails on changed PowerShell hashes.

Final full run: 20 Python tests, nine Pester tests per actual PS5.1.26100.9444/
PS7.6.5 host, zero syntax/new-scaffold findings, all 14 original/extracted cases
matching full observations, and six actual entrypoint probes including three
parallel launches. Exact outputs/page identities matched; source/input/neighbor/
decoy/shared-TEMP sentinel hashes were preserved. Actual Documents outputs were
new GUID/marker-owned directories; only these were inspected and observed
cleaned. Test-only tracked-tree termination and fail-closed preservation were
verified after a real descendant-after-timeout reproduction. Two piped PS7
probes retain RawUI cursor errors; this does not certify console UX/error paths.
Sixty retained launcher analyzer findings per host are observations, not a
runtime static pass. Tested-path digest:
`2a456dfc50d0ae62e93eaa52c10717ed9a90a3d819ab80560b5a5006f6387054`.
Independent source/receipt/privacy review found no remaining blocker; no
submitted GitHub review or CI pass is claimed.

Regular GIL CPython **3.14.8 x64 only**, plain **pypdf 6.19.0**;
dev ReportLab **5.0.1**, Pillow **12.3.0**, charset-normalizer **3.5.2**.
Use SUPPORT_AND_SETUP.md for a fresh explicit hash-required dev venv; do not
assume an old external TEMP environment remains present. Exact-version shell
tools are Pester **6.2.0** and PSScriptAnalyzer **1.25.0**. Use the central
runner, explicit absolute host paths/external reports and host-owned module
paths as described in tests/README.md. Children request process-only RemoteSigned;
stored policies remain unchanged and managed policy wins.

## M1-T02 boundary

Reproduced manual defects remain: `1` creates zero files/exit 0; `1,4,7` loses
pages 1-3; mixed invalid tokens are filtered. `4,7` is the complete-coverage
control. Implement the strict SPLIT_CONTRACT grammar, validate every token,
always include physical page 1 and normalize only after validation. Cover
AC-013 through AC-018, deterministic page identities and seeded valid starts.
Keep source inputs immutable; no bookmark/output/discovery/UI scope expansion.

Behavior fixes intentionally invalidate current mechanical-equivalence AST/
known-bad assertions. Preserve original and M1 historical receipts/guards, but
explicitly transition the affected current suite/central route to corrected
manual oracles; do not claim changed code is still the original or silently
rewrite its historical expectations. Remaining bookmark/plan/metadata fixes
belong to later M1 tasks. One task per thread; do not start M1-T03 now.

## Installed Calibre and remaining gates

User-requested official portable **Calibre 9.15.0** is installed at
`$UserProfile\Apps\Calibre915\Calibre Portable`. Converter:
`$UserProfile\Apps\Calibre915\Calibre Portable\Calibre\ebook-convert.exe`.
Published SHA-256/SHA-512 and Valid Windows signatures were checked; installer
and version probe exited 0. No elevation, global PATH or policy changes occurred.
This trusted nonstandard absolute path is outside the launcher's unchanged
search paths: integration remains M2-T04. GUI interaction and actual EPUB/AZW3
conversion remain NOT RUN; real conversion gates remain M4-T04 and package/
release acceptance. The binaries are external and must not be bundled/committed.

Explorer, all launcher/error/discovery paths, repaired engine, renderer 26.10.0,
clean OS/extracted release package, CI and public downloads remain NOT RUN.
Canonical CI task is **M4-T05**, correcting older M0 prose's M5-T01 reference.
Live main had no protections/rulesets/required checks or workflows/runs; inspect
current state again. Windows10/ARM/UNC/other Python remain unclaimed. No tag or
release exists. Only RELEASE_RUNBOOK.md's public non-draft/non-prerelease
v1.0.0, matching anonymous downloads, fixed tag and synchronized final main
finish the project. Unsigned release remains allowed with disclosure/checksums.
