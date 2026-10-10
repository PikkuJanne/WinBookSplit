# Actual v1.0.0 publication and completion

Publication is the requested endpoint, not a separate future project. No interim public GitHub releases or release tags. A draft **of v1.0.0** for assembly is allowed, but never completion. Do not change repository settings or delete/retarget any existing tag/release. If v1.0.0 already exists, inspect it first; verify and close only if it is exactly the intended completed release. A conflicting existing version is a blocker requiring the owner's decision, not permission to overwrite or publish v1.0.1.

## Gate A: accepted scope and source

All mandatory acceptance cases through M5 pass with actual evidence; no open data-loss/overwrite/process/security blockers. Any unsupported platform/feature is explicitly scoped in user docs, not falsely tested. Windows 11 x64, Windows PowerShell 5.1, tested PS7, actual EPUB and AZW3 conversions, and real launcher smoke are required. Encrypted/unsupported document rejection must match D05/D12. Full automated tests/static checks/CI for the intended release input must pass. Package/user docs/license/notices and unsigned status are accurate.

Merge/reconcile completed scope through the repository process. Verify clean main, live GitHub identity and main SHA. Use one `VERSION`/equivalent canonical source = `1.0.0` in the final candidate; do not publish intermediate development labels. Record historical `2.1` banner as pre-release development naming, not an old public SemVer release.

Freeze a clean tested **release commit R** on main. Record R, its tree/payload input manifest, CI evidence and environment. No runtime/build/dependency edits after acceptance without a new candidate and retesting. Documentation evidence may follow, but released bytes must still match R.

## Gate B: deliberate package from R

The M5-T02 builder/validator is `tools/release/build_package.py`, with explicit
`tools/release/payload.json` membership and the source recipe in
`tools/release/README.md`. Use its committed tooling bytes and pass R as a full
commit ID. Its successful candidate build establishes package integrity only;
it does not authorize skipping Gate A, exact-asset Gate C or publication checks.

Implement/reuse a small package builder that accepts an explicit source commit and allowlists runtime files. Include entry points, extracted Python engine, any required PowerShell helpers, canonical version, requirements/setup instructions, README, LICENSE, actual third-party notices and useful icon(s). Exclude `.git`, `.venv`, tests, handoff docs/tools, caches, logs, textbooks, credentials, converter work files, developer-only scripts and the large poster unless deliberately justified. Do not ZIP the working directory indiscriminately or assume source-archive downloads are a ready-to-run distribution.

Build `WinBookSplit-v1.0.0.zip` from R in a clean controlled build directory. Build `release-manifest.json` identifying R, version, dependency set, allowlisted ZIP paths and per-file hashes, archive hash/size, and builder/tool versions. Exclude self-hashes: a manifest never includes its own hash inside itself. Create `SHA256SUMS.txt` listing the ZIP and release-manifest (and other explicit assets, if any), not SHA256SUMS itself. Hash the actual final bytes after all metadata/signing operations. Pin file ordering and archive timestamps where practical; otherwise promise a repeatable recipe and per-file provenance rather than untested byte-reproducibility.

Assets live outside source control or in an intentionally ignored dist directory. Check ZIP traversal/absolute/symlink entries, missing engine files, accidental secrets/private paths, non-allowlisted files, and exact VERSION. Produce human release notes from tests/changes, not unsupported marketing claims.

## Gate C: exact-asset acceptance

Extract **that ZIP** to a fresh Windows directory; verify installed/runtime dependency paths and no imports from checkout; test both advertised PowerShell hosts, real Explorer/launcher behavior, manual/auto/Level2, PDF/EPUB/AZW3 and safety cases. Use D11's explicitly documented isolated-workstation test or an available Sandbox/VM. Record the ZIP SHA-256, R, Windows build, tool versions, fixture IDs and actual results. Inspect representative rendered page fidelity. No rebuilding or modifying the ZIP after this gate without a new hash and retest.

Mandatory test not executable in the current environment = blocked with the precise missing local action. Do not rename a mock/syntax/CI-only test 'Windows smoke'. A user-run result can qualify when its artifact hash/environment/procedure/outcome are recorded. Signing remains optional; unsigned disclosure and checksum instructions are required.

## Gate D: fixed tag and assembled draft

Recheck live main/R ancestry, clean state, release/tag absence (or exact safe resumption). Create one annotated `v1.0.0` tag on R and push **only that tag**, never `--tags` blindly. Verify the remote annotated tag's peeled commit equals R; lightweight tag comparisons must resolve to the commit too. Never infer tag identity from release `target_commitish` alone.

Create/resume a draft for that existing remote tag. Explicitly verify the tag, attach the tested ZIP, manifest and SHA256SUMS, and inspect names/sizes/hashes before publication. No wildcard upload of arbitrary dist files, overwrite/clobber, or automatic latest-main tag creation. Current GitHub CLI supports `--draft` and `--verify-tag` [S8]. Respect any existing immutable-release behavior; finish assets before publishing.

Representative commands below are intentionally manual runbook steps, **not an unattended script**. Resolve and check every variable/path and every native exit status first:

```powershell
$repo = 'PikkuJanne/WinBookSplit'
# $R is the full accepted release commit; all paths name already-tested assets.
git tag -a v1.0.0 $R -m 'WinBookSplit v1.0.0'
git push origin refs/tags/v1.0.0:refs/tags/v1.0.0
git ls-remote origin 'refs/tags/v1.0.0' 'refs/tags/v1.0.0^{}'
gh release create v1.0.0 $zipPath $manifestPath $sumsPath --repo $repo --draft --verify-tag --title 'WinBookSplit v1.0.0' --notes-file $notesPath
```

Do not rerun tag/release creation blindly after a partial failure. Inspect and resume exactly the intended draft; conflicting artifacts remain a blocker. The bundle contains no automatic publisher.

## Gate E: publish, then independently verify public availability

After A-D pass and the intended draft is inspected, publication is authorized by the start prompt. Use the repository tooling and retain its actual result; an authentication/permission prompt must not be bypassed. GitHub CLI offers draft-to-public editing [S9].

```powershell
gh release edit v1.0.0 --repo 'PikkuJanne/WinBookSplit' --draft=false --prerelease=false --latest
```

Inspect the returned release through GitHub/API: tag `v1.0.0`, draft=false, prerelease=false, public published URL/timestamp, expected assets, correct repository. Fetch the remote tag again and prove R. Do not use a successful command's exit alone as proof of public visibility.

Download ZIP, manifest and checksums **without authentication** into a new directory from the public release links. Recompute SHA-256 and compare to the pre-publication accepted values and downloaded SHA256SUMS; do not merely compare an uploaded file to itself. Inspect/extract the downloaded ZIP and run at least its version and minimal PDF smoke in the supported Windows setup; compare exact archive bytes to the full tested package. A checksum mismatch, missing/private/broken asset or wrong tag is a failed gate. Do not silently replace a published immutable artifact: report the failure for owner resolution. No new public patch version is authorized by this project.

## Gate F: durable closure and synchronized main

Fill a real copy of `templates/RELEASE_EVIDENCE.json`: R, tag/ref proof, actual release URL, publication fields, public asset URLs, expected/downloaded hashes, Windows/Calibre evidence, CI references and outcomes. Set completed task/status/next-session records accurately. Commit only these closure records as D, merge/push through the scoped process, and verify clean local main equals live main. R must be an ancestor of D and runtime/package input paths unchanged. Do not move v1.0.0 to D, rebuild assets from D, or claim D was the tested release source.

The final sync receipt for D can live in the thread/PR outside its own commit; no self-referential hash loop. In the first closure record, `closure_commit` may remain null with `final_current_sync_receipt_location` pointing to that external receipt; do not manufacture D's own SHA inside D. A later historical record may fill it without moving the release tag. A historical receipt recorded earlier does not replace a fresh check.

## Definition of done: all true

- All in-scope mandatory acceptance tests and actual Windows/Calibre package evidence passed; remaining limits are accurately disclosed.
- Public non-draft non-prerelease release v1.0.0 exists in PikkuJanne/WinBookSplit; no interim versions were published.
- Remote tag resolves to accepted R; shipped version/provenance and asset bytes match R's tested package.
- Anonymous download verification and downloaded-package smoke succeeded with recorded hashes.
- Closure records are durable; local main is clean and equals live GitHub main; any post-R commits are allowable non-payload closure only.

Anything less is IN PROGRESS or BLOCKED. Final user report gives the actual release URL, R, final checkpoint/main SHA, tested environments, asset hashes, and truthful known limitations. Never say 'published' based only on a plan, tag, draft, CI artifact or uploaded ZIP.
