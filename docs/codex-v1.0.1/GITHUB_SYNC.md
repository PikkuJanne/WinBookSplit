# Local and GitHub workflow

Use the existing authenticated origin for PikkuJanne/WinBookSplit. Verify fetch and push destinations, current branch, worktree state and live remote before changes. Review URL rewrites/custom SSH aliases locally; helper refusal is a prompt for inspection, not permission to overwrite configuration. Preserve unrelated dirty work rather than resetting/stashing it automatically.

Use a scoped branch such as `codex/v1.0.1-compatibility`. A live branch with that name may already contain work; inspect and resume rather than overwriting it. Import only new handoff paths and then deliberately update current root AGENTS guidance. Current main's artwork is retained.

After each meaningful task: run its tests, update brief evidence/status/next task, stage only intended files, commit with a meaningful message, push normally and query the live branch. The supplied `tools/codex-v1.0.1/check_sync.py --repo <root>` returns SYNCED only for a clean worktree whose HEAD equals a fresh live remote branch; it is not a CI or release check. Failed network/auth is UNKNOWN; dirty/different/absent branch is UNSYNCED. Do not substitute cached origin refs or 'push attempted'.

Do not create a circular proof by embedding a commit's own hash into that commit. A committed receipt records the earlier checkpoint it observed. After a receipt commit, verify that new HEAD against the remote in the session's external output or next receipt.

Use the repository's normal PR/merge rules. Review all relevant diffs and passing checks before merge; do not bypass protections or invent an approval requirement absent from actual policy. Publication of v1.0.1 is delegated, but actual tool/environment approval gates remain binding. No force push, moving tags, deleting old releases, settings changes, new versions or spending.

The release source R is the reviewed commit chosen after merge/integration. A later documentation-only closure commit C is allowed with an explicit changed-file/package-member audit; R must remain the release tag. Local main equals remote main at C, while the public artifact still identifies R. Do not retag to C just to make those two different purposes share a hash.
