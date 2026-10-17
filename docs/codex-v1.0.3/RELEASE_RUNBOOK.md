# v1.0.3 publication and synchronized closure

## Gates

All T01-T07 acceptance requirements must pass with actual recorded scope, except
optional observations explicitly labelled as such. In particular require:

1. Old Store-route regression reproduced and corrected; query, downstream identity,
   alias activation, enumeration and concrete interpreter probes tested.
2. Existing application regressions and required CI pass on the release source.
   Native Win11 regression and existing supported Python branches are retained.
3. The actual Windows10 Store installation passes automatic discovery, setup and
   BAT on the EXACT frozen distribution ZIP, with no forbidden substitution.
4. Native receipts are complete and reviewed, not simply declared passed. Artifact
   hashes, source R and package hash agree; offline native gate checker passes.
5. Documentation/remedies and release version are correct, root routing is updated,
   secrets/private logs excluded, package allowlist correct, current dependency
   security scope reviewed honestly. Scripts may remain unsigned.

## Git discipline at every task

Read `git status`, current branch/upstream/remotes and live origin state. Preserve
unrelated changes. Work on a scoped branch (for example
`codex/v1.0.3-store-discovery`) without auto-reset/stash/force operations. Stage
specific paths, commit meaningful changes, push normally, and compare `git
rev-parse HEAD` with a freshly queried live branch SHA (`git ls-remote`), not only
a cached tracking branch. Failed push/network/auth => UNSYNCED/UNKNOWN, never
synchronized. Follow branch protections and required reviews; no settings changes.
Write each task's receipts/status before committing. Record that commit hash in
the user handover or next evidence record, not recursively in its own contents.

## Freeze

Inspect live existing releases/tags. Preserve v1.0.0-v1.0.2 peeled/tag-object SHAs,
asset names/digests and release text. If v1.0.3 already exists, do not overwrite or
select another number silently; inspect/resume the known release or stop with a
precise conflict. Do not mutate published existing releases.

Set VERSION to 1.0.3, fix docs and CHANGELOG, include setup BAT/runtime changes in
the release allowlist. Use the existing build pipeline. Do not ship handoff tools,
raw diagnostics, fixtures, .venv, caches, Python/Calibre installers or test libraries.
Freeze a committed source R on the branch/merged main that will be tagged. Build
`WinBookSplit-v1.0.3.zip`, `release-manifest.json`, and `SHA256SUMS.txt` from R.
Record exact hashes and member inventory. Keep original build bytes unchanged
through native acceptance and publication. Packaging changes require a fresh
freeze and new exact-ZIP acceptance, not a reused result from an older artifact.

## Native gate command

Using any compatible existing developer Python OUTSIDE the target scenario, run:

```
python tools/codex-v1.0.3/check_native_store_report.py <returned-report.json> --expected-commit <R> --expected-zip-sha256 <SHA256>
```

The helper verifies report structure, pre-existing runtime correlation, required
no-substitution observations and supplied receipt hashes. It does not attest that
Windows ran; inspect the raw-to-redacted native receipts and maintained validators
as well. The target kit itself must not need this Python to start discovery.

## Publish only accepted bytes

After passing reviewed work reaches the intended release commit, tag exactly R.
Verify the peeled tag commit, not just tag name or GitHub target_commitish. Use a
single explicit release path (existing tooling or GitHub CLI), not two competing
publishers. Upload the frozen ZIP, manifest, checksums and precise release notes.
Publication must be public, non-draft, non-prerelease. Do not regenerate the ZIP
from GitHub's auto-source archive or a later worktree. No retag/rebuild/reupload
on a failure: stop and preserve evidence.

Release notes explain the Store discovery repair, actual tested installation
route, one-time setup and unchanged security/format scope. Do not market supplied
runtime, mocked, or Server tests as owner-installed Store discovery. Disclose
unsigned status and any untested physical Explorer action. Avoid exposing local
paths or identifying workstation details in public evidence.

## Verify and close

Download the public assets WITHOUT authentication and hash the bytes. Compare to
the frozen accepted values and manifest; require successful public API/release
status. Verify old releases/tags/assets remain unchanged and v1.0.3 points to R.
Record public URL, R, ZIP digest, native receipt and CI results. A documentation-
only closure commit C may follow; state explicitly that the tag remains R. Merge
and synchronize local main to the live GitHub main with a clean intended worktree.
Do not delete the owner's verified extraction or point an existing shortcut at an
untested directory. Final handover identifies the tested launcher location and
normal setup/use commands, remaining limitations, and evidence links.

A tag, draft, source test suite, or schema-valid report alone is not completion.
