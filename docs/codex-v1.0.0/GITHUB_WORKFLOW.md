# Local/GitHub synchronization and thread handoff

## First inspection

Determine Git root, applicable instructions, actual origin fetch/push URLs, branch, upstream, HEAD, dirty/untracked files and in-progress merge/rebase. Verify origin is PikkuJanne/WinBookSplit (SSH or HTTPS, correct host); inspect `git remote -v` locally without publishing credentials. Compare live branch SHA independently, not just a cached `origin/main`. Inspect tags, GitHub releases and any ongoing PRs. New work after the baseline must be reconciled, never reset away.

If dirty: preserve unrelated work. Do not auto-stash, `reset --hard`, `git clean`, overwrite, force push or abort another person's operation. A task-specific worktree may be used after safely identifying the actual source state; do not fork off a stale clean tree to dodge existing changes. Ask only for genuinely non-resolvable ownership decisions; otherwise record the safe continuation.

## Branches and normal writes

Use a normal feature branch such as `codex/winbooksplit-v1-m0`, then one per milestone. Keep one active workstream unless explicitly coordinated. Push each task checkpoint to the matching remote branch; create/reuse a PR with the scoped changes and evidence. At a milestone review the cumulative diff, check required CI and merge through the available protected process. The start prompt delegates these ordinary writes; permissions/security gates still apply. Do not change repository rules, visibility, default branch, or secrets.

After merge, fetch the actual default branch, update local main without discarding changes (`pull --ff-only` only when appropriate), and verify local/main against live GitHub. Prefer retaining milestone branches until final audit; deletion is not needed. If conflict/divergence occurs, inspect/reconcile normally; never force a branch to appear synced.

## Checkpoint protocol without self-referential SHA loops

1. Implement a narrow task; run targeted tests; record local evidence and the runtime source-tree digest. Stage intended paths only. Commit the implementation as C.
2. Run verification on clean C as needed and push C normally. Compare live branch SHA with C. Preserve receipt outside the worktree or in the PR. A failed push/check leaves the task in progress/blocked.
3. Commit a small evidence/status/handoff update E referencing **C** and its already observed test/sync results. Store the actual commands, test IDs and next task. This is where TASKS.json can record `done`, `evidence_paths`, and `implementation_commit: C`. Do not claim E was synchronized inside E before it is pushed. A later checkpoint can record E's historical receipt.
4. Push E, run the live sync checker, and state the actual clean local HEAD/live SHA and timestamp in the thread/PR receipt. Do not edit tracked files afterward solely to insert E's own SHA. Evidence directories excluded from runtime packaging can advance without retesting unchanged runtime bytes; if application/build/dependency inputs changed, retest them.
5. At the next thread start, recheck live state. A `done` task with an incomplete evidence push is not a valid checkpoint; repair synchronization before continuing. TASKS.json is a durable record, not authority over live reality.

Tests may be run before C; cite the exact source digest and rerun after C if changes differ. Never fabricate clean-commit test results. A single implementation+evidence commit may be used when evidence honestly describes the tested worktree digest; the next receipt still lives outside its own commit. This is a bookkeeping convention, not a demand for gratuitous commits.

## Fresh live verification

`tools/codex-handoff/check_sync.py` requires a clean Git root, correct single fetch/push repository identity, a non-detached branch, and a fresh `git ls-remote` comparison using the approved GitHub URL. It returns UNSYNCED for dirty/missing/different refs and UNKNOWN for network/auth/config errors. It does not trust cached remote-tracking refs. Run it after every push and at thread start. It does not inspect CI, review or releases: query those separately with authenticated GitHub tooling.

A representative manual sequence (check every native exit code; placeholders are not literal branches):

```powershell
git status --short --branch
git diff --check
git add -- <explicit-intended-paths>
git commit -m '<task-id>: <specific-change>'
git push -u origin HEAD
python .\tools\codex-handoff\check_sync.py --repo .
```

No 'synchronized' claim from `git status` alone, a previous tool output, a successful local commit, or a queued push. `git ls-remote` is a point-in-time observation, not a lock; recheck before release [S7].

## End-of-thread report

Task/substep, changes, tests actually run, unresolved/conditional IDs, implementation/checkpoint commits, branch and PR, live remote result, and next exact task. Stop at a coherent task boundary; do not start a broad rewrite because another milestone doesn't fit in context.

## Release closure

Release source commit R is fixed and tested. The final documentation-only closure commit D may follow R. Local main must equal live main at D; R must be its ancestor and the release payload paths must be unchanged. Do not retag R to D just to make all displayed SHAs equal. See RELEASE_RUNBOOK.md for source vs asset/evidence integrity.
