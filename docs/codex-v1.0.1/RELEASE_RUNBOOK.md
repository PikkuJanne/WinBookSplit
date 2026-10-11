# Publish and verify WinBookSplit v1.0.1

The user authorized this publication. No extra business signoff or repeated owner test programme is needed when the required automated checks pass. Tool approvals and actual GitHub protection still apply.

## 1. Inspect state and protect the earlier release
Query current branches, tags and releases live. Snapshot v1.0.0's peeled tag/source, published state and asset names/sizes/hashes. Do not change them. If a public v1.0.1 already exists, do **not** clobber it or recreate/move its tag: verify its source and bytes and either finish valid closure or report the conflict. A draft can be resumed only after proving identity with this work. Do not treat a failed request as absence.

Confirm V101-01 through V101-11, current security/source review, compatibility policy, scope exclusions, tests/CI and changed-code review. Correct README's active candidate labels and version references before freeze; historical evidence can remain historical. Do not reopen completed v1.0.0 release tasks.

## 2. Integrate and freeze source R
Merge scoped passing reviewed work using normal protection. Reconcile any concurrent main changes without losing artwork. Select exact commit **R**, containing application VERSION `1.0.1`, setup, shared compatibility policy, documentation and the updated package allowlist. R must be reachable from main. Record exact build/test tools separately from user runtime compatibility.

Run integration checks for R, not an earlier branch head. Build with the existing reviewed `tools/release/build_package.py` after updating only the necessary contracts/payload; inspect its real CLI rather than inventing arguments. Freeze these assets outside the checkout:

- `WinBookSplit-v1.0.1.zip`
- `release-manifest.json`
- `SHA256SUMS.txt` (ZIP and manifest; no self-hash loop)

The manifest includes source R, application/policy version, actual build tools, accepted runtime range/security floor distinction, package members and hashes. Include Setup-WinBookSplit.ps1 and any policy/helper it needs. Exclude .venv, developer dependencies/fixtures, handoff plans, raw diagnostics, credentials and user documents. Keep the current intended icon, not an old path silently substituted from v1.0.0. Existing artwork need not all be shipped in the runtime ZIP.

Audit that package bytes came from R and only allowed files are shipped. Do not edit the ZIP manually or build from uncommitted working files. Test the actual extracted package, not a developer checkout with hidden dependencies. Repeatability may be checked with the same build toolchain; do not require ZIPs from different Python minors/zlib builds to be byte-identical. Runtime output correctness is a different check.

## 3. Final-package gates
Run V101-13's native Windows 10 and Windows 11 harness on the **same frozen ZIP hash**, plus required runtime-matrix/reference regression, setup, PDF-policy and real EPUB/AZW3 checks. Use HUMAN_TESTING.md for the only fallback owner session. Record native build and test scope. No old GUI checklist or independent manual re-audit is required. Test-kit harness code/report hashes are recorded separately and must not mutate production package members.

Any shipped-file change creates a new R/ZIP and invalidates affected acceptance. Complete fixes before publishing; rerun the necessary automated checks rather than arguing that a previous ZIP was 'close enough'. No pass label for skipped/native-unavailable tests. A documentation-only non-package evidence commit may follow without rebuilding R's accepted ZIP.

Review security sources again before publication. Resolve material applicable issues; keep exact runtime floors justified, not latest-patch equality. Preserve existing process/cleanup/encryption/origin protections. Record any security-only older-patch rejection as distinct from syntax incompatibility.

## 4. Tag, draft and publish the accepted bytes
Check actual `gh release create --help` / `edit --help` and repository tag workflows. A tag push must not accidentally trigger an unreviewed second build/publication. Only after all prepublish gates pass create an annotated v1.0.1 tag at **R**, verify locally, push that exact tag once and read its peeled remote target. Never force it.

The following is an illustrative command sequence to resolve to actual paths and validate step by step, **not a blind script**:

```powershell
git tag -a v1.0.1 $R -m 'WinBookSplit v1.0.1'
# Check success, local tag target and live tag state before the next write.
git push origin refs/tags/v1.0.1
# Verify the remote peeled tag equals R.
gh release create v1.0.1 $ZipPath $ManifestPath $ChecksumsPath --repo PikkuJanne/WinBookSplit --verify-tag --draft --title 'WinBookSplit v1.0.1' --notes-file $ReleaseNotesPath
# Inspect exact draft metadata and uploaded assets; check every exit status.
gh release edit v1.0.1 --repo PikkuJanne/WinBookSplit --draft=false --prerelease=false --latest
```

Explicitly stop on each native nonzero exit; PowerShell does not universally make external-command failure terminating. A draft is for assembly, never evidence that the project is done. Release notes describe Windows 10/11, the actually supported minors/floors, one-time setup, reviewed library/converter, known limits, unsigned status, test scope and security-servicing caveat. Do not promise every patch was tested or that a hash is a signature.

## 5. Public verification (required after publication)
Read public release metadata **without credentials**. Confirm tag v1.0.1, draft=false, prerelease=false, a publication timestamp and expected assets. Download each uploaded asset anonymously through its actual public URL to a new directory and compute full hashes. Compare ZIP and manifest against the frozen local identities and check checksum contents and per-member manifest. Authenticated draft download or a successful HTTP status alone is insufficient.

Use the verified public ZIP in an automated extraction/smoke check. Equality with the exact native-tested ZIP establishes the connection to earlier native acceptance; this does not require a second human session. Read the live peeled tag and confirm R. Re-read v1.0.0 and prove its tag/assets remain unchanged. Network or auth anomalies produce UNKNOWN, not success.

## 6. Close and synchronize
Commit a concise redacted closure record/status update to main using normal rules. Audit that any **C after R** changes only non-shipped evidence/handoff files; if it changes shipped members, those are future unreleased changes and this patch cannot claim them. Avoid them during closure. Tag v1.0.1 stays at R.

Verify clean local main at C equals a fresh live remote main. Record public release URL, R, ZIP SHA-256, tested runtime/OS matrix, native automated evidence, any single human observation and limitations. A final session message can supply the current sync receipt without another self-referential commit.

Finish only after all six sections. Do not publish a false Windows 10 support claim to satisfy the version deadline; finish independent work and report a precise blocker when an essential native capability is genuinely unavailable.
