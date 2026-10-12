# Publish v1.0.2 — completion runbook

This specific publication is already delegated. Respect actual platform approval and branch protection. No routine 'may I publish?' stage. Never replace existing published v1.0.0/v1.0.1 assets/tags or silently reuse an existing conflicting v1.0.2.

## 1. Preflight and freeze

Read active tasks and RELEASE_GATES.json. Confirm required Python 3.15, setup BAT, executable remedies, existing-branch, security and source regression work is real and reviewed. Check dependency advisory sources and exact current artifact availability; do not turn 'latest' into an arbitrary new pin. Security-induced version changes require tests, release-note scope and policy rationale.

Read live refs/releases first. If v1.0.2 already exists, inspect peeled tag, body and asset hashes; resume verification only for the same intended release. A conflicting existing release/tag is a blocker, not permission to delete/move/overwrite it. Snapshot both prior releases/refs/assets; reuse baseline only as an initial comparison, not fresh proof.

Ensure VERSION and all active docs/help/release metadata say 1.0.2. New Setup-WinBookSplit.bat and shared policy/helper changes must be in the distribution allowlist and source/member checks. No personal documents, `.venv`, wheels, installers, log files, caches, credentials or Codex handoff development files in the end-user ZIP. Existing artwork and licenses remain untouched except correcting packaging references if necessary.

Merge reviewed changes as permitted and identify a committed source **R**. Run required source CI/tests on R. Worktree must be clean before building with the existing reviewed packager. Record exact builder/toolchain/dependency locks, source R and resulting asset hashes. A reproducible developer toolchain does not constrain the accepted end-user patch range.

## 2. Build once and test actual bytes

Create `WinBookSplit-v1.0.2.zip`, `release-manifest.json`, `SHA256SUMS.txt` using the existing release tooling. Manifest binds source R, version, dependencies and allowed members/per-file hashes. Checksum file covers ZIP and manifest; it need not hash itself. Independently inspect no traversal/duplicate/unsafe member names and exact member inventory. Verify actual source-to-package bytes. Preserve these accepted candidate bytes; signing is optional and not a purchase/blocking requirement.

Call ZIP hash **Z**. On extracted Z run native Windows 10 and Windows 11 Python 3.15 gates plus scoped old-branch/PS7 regressions as TESTING.md states. A new ZIP is not automatically covered by a previous run. Verify MOTW controls on actual shipped scripts, new BAT, emitted command round-trip, final selected interpreter, PDF output counts/content/order and unchanged source. Capture approved scope accurately. No obligation to repeat an unchanged manual GUI programme.

Any changed runtime/dependency/package bytes after acceptance invalidate impacted Z evidence. Fix, rebuild and rerun focused/full gates appropriate to the change. Do not merely edit hashes in reports. If only external final evidence/docs change and package bytes remain identical, do not manufacture a new candidate loop.

## 3. Publish the accepted release

After V102-11 and other gates pass, create a fixed annotated `v1.0.2` tag at R (unless the same verified tag already exists); push normally and check the live peeled commit. Create a draft release and upload the exact accepted three assets. Inspect names, sizes, hashes, version and source before switching to public non-draft, non-prerelease. Use GitHub CLI or API through available approved tooling; consult the actual action/schema rather than inventing tool flags. No assets rebuilt during upload and no `--clobber` against prior public releases.

Release notes describe Python 3.15 plus preserved serviced 3.13/3.14, numeric same-branch patch acceptance, unknown-minor policy, new setup BAT and correct process-scoped invocation. Include exact tested OS/shell/runtime combinations, automation scope and honest limitations, immutable source R, unsigned status, hashes, Windows 10 servicing notice and existing document/conversion boundaries. Do not say every later patch or all Windows 10 builds were executed. Do not call a source ZIP the tested distribution asset.

## 4. Verify from the public user's side

Using a fresh unauthenticated download context without authorization headers, download each asset from the public v1.0.2 release to a new owned directory. Allow only ordinary GitHub asset redirects; compare hashes and sizes against accepted local bytes, inspect archive and VERSION/new BAT/member hashes, and compare manifest's source R with the peeled tag and reviewed source. An authenticated draft read is not anonymous publication evidence. If public access fails or bytes differ, mark release verification failed and diagnose; never claim done.

Re-fetch v1.0.0/v1.0.1 refs/release assets and compare the identity snapshots. Download counts may change; compare content identities, tag targets, asset names/sizes/digests and release status, not volatile counts. Do not modify old release text/assets. Record public release URL, R, Z, manifest/checksum identities, publication state/time, final native reports and verification results.

## 5. Close and synchronize

Update the v1.0.2 status/case records truthfully and commit documentation-only closure D if necessary. Keep R/tag fixed. Ensure local main contains R, is clean, and equals fresh live origin/main at D. Compare release-member bytes between R and D when D differs. No need to rerun native tests for a closure-only file outside the runtime package. Final response includes public release URL, source R, asset SHA256s, actual test matrix and any supplemental checks not run, unchanged old releases and fresh sync result.

A release with missing main features/native proof, a draft, tag only, green tests alone, failed push, or unverified public asset download is **not complete**.
