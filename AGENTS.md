# WinBookSplit working agreement

## Scope

Incrementally improve this existing local Windows PDF/ebook splitter through its first public GitHub `v1.0.0`. Preserve `WinBookSplit.bat`, `WinBookSplit.ps1`, the console/drag-and-drop workflow, Python/pypdf, and optional Calibre. Extracting existing Python into a shipped file is allowed; replacing the application/engine with a new framework is not.

No website, service, uploads, telemetry, native ebook output, OCR/AI chapter inference, DRM removal, installer, updater, or interim public release. No paid/signing prerequisite. Never process private textbooks/documents in CI or commit them.

## Start each thread

Read `docs/codex-v1.0.0/STATUS.md`, `NEXT_SESSION.md`, `TASKS.json`, `SCOPE_AND_DECISIONS.md`, then the selected task and its specs. Inspect actual Git/origin/branch/worktree and live remote before changing anything. The recorded baseline is evidence, not a reset target. Preserve unrelated edits and existing instructions.

## Implementation and correctness

Use one thread-sized task at a time. Reproduce bugs before changing them. Reuse existing code; keep refactors separate from behavior fixes. All split modes share one validated plan: ordered nonempty ranges cover every physical input page exactly once. No silent first-page loss, parent-boundary crossing, ignored invalid manual tokens, overwrites, or success on zero files.

Run both process streams safely. No shell evaluation of paths/bookmark text. Temporary/output cleanup owns only this run's files. No global execution-policy changes or elevation. Source files are immutable inputs. Preserve meaningful errors and nonzero exit codes through the batch launcher.

## Evidence and checkpoint

Use `TESTING.md` and each task's acceptance IDs. Record commands, runtime versions, source commit/tree digest, actual outcomes, skips, and defects. No fabricated/manual/Windows/Calibre passes. Ordinary syntax/unit success is not an end-to-end Windows test.

Update task/evidence/status/next-session files; stage intended paths only; commit; push normally; verify clean local HEAD equals the live remote branch. Missing network/auth or a failed push means UNKNOWN/UNSYNCED. `GITHUB_WORKFLOW.md` explains historical receipts and avoids self-referential commit hashes. Never force push, reset hard, auto-stash, or delete work.

## Authority and completion

The user delegated ordinary scoped commits, pushes, PRs, merges of passing reviewed work, and final v1.0.0 publication. Respect environment approvals and existing protection; never bypass them. Destructive Git operations, settings/visibility changes, spending, new release versions, or scope expansion require separate authority.

A tag or draft is not completion. Follow `RELEASE_RUNBOOK.md`: tested commit and package, real Windows/Calibre checks, public non-draft/non-prerelease v1.0.0, anonymously downloaded matching assets, fixed tag, and synchronized final main. Signing is optional and unsigned status must be disclosed.

## Code review rules

Block silent page omissions, unsafe output/cleanup, ignored stderr/exit codes, dependency shadowing, filename/command injection, unverifiable test claims, public private-data leaks, and tagged/package byte mismatch. Do not exchange the user's narrow tool for a generalized platform.
