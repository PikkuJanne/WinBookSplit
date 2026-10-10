# Build and inspect an unpublished package

This developer tool builds three local assets from an explicit full commit:
`WinBookSplit-v1.0.0.zip`, `release-manifest.json` and `SHA256SUMS.txt`.
It does not create a tag, upload, sign or publish. M5 candidates remain separate
from M6's accepted release commit and exact-asset Windows acceptance.

Run with a trusted Python, Git and the builder files from the intended source
commit. The builder, `payload.json` and the reused CI input checker must match
their committed bytes exactly. A checkout that converts these files' line
endings must use an ordinary Git source archive/extraction with original bytes
for the tooling; the source repository argument still names the Git root.
Changed/untracked tooling is refused. Payload bytes are always read from Git
blobs with replacement refs disabled; no working-tree file is adopted.

`payload.json` explicitly lists 28 files. The shared checker validates all
runtime import/helper dependencies and fixed hashed runtime requirements.
Current setup/policy/fidelity/ebook/support docs, license/notices and icon are
included. Tests, development tools, handoff history, poster, venvs, caches,
documents and runtime/converter binaries are excluded. Relative document links
must resolve in the payload; source-only references use GitHub links.

Use a fresh **absent**, absolute output directory whose parent already exists
outside the Git checkout. Existing destinations are refused, including junctions,
symlinks, UNC and device paths. No recursive cleanup occurs. A failed write or
validation keeps `.package-incomplete` and diagnostic artifacts; verification
rejects the marker. Preserve failed work for diagnosis and choose a new path.
The tool is not a sandbox against another process running as the same account.

```powershell
$BuildPython = 'C:\Tools\Python3148\python.exe'
$SourceRoot = 'D:\projects\WinBookSplit-main'
$Builder = Join-Path $SourceRoot 'tools\release\build_package.py'
$SourceCommit = (& git -C $SourceRoot rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0) { throw 'Cannot resolve the intended committed source' }
$CandidateAssets = 'C:\LocalBuilds\WinBookSplit-M5-candidate-01'
& $BuildPython -I -B $Builder build --repo $SourceRoot --commit $SourceCommit --output $CandidateAssets
if ($LASTEXITCODE -ne 0) { throw 'Package build failed' }
& $BuildPython -I -B $Builder verify --repo $SourceRoot --commit $SourceCommit --output $CandidateAssets
if ($LASTEXITCODE -ne 0) { throw 'Package verification failed' }
Get-FileHash -LiteralPath (Join-Path $CandidateAssets 'WinBookSplit-v1.0.0.zip') -Algorithm SHA256
Get-FileHash -LiteralPath (Join-Path $CandidateAssets 'release-manifest.json') -Algorithm SHA256
```

Check the source SHA before building: `HEAD` here is resolved by the operator,
and the builder accepts only the full lowercase commit ID. Final M6 builds
must use the frozen, accepted main commit R instead of guessing latest HEAD.

ZIP entries use sorted paths, no directory entries, stored bytes, a fixed
1980-01-01 timestamp and regular-file 0644 attributes. Independent verification
checks every entry and raw byte against the commit, regenerates the expected
archive recipe and manifest, and verifies both checksum lines. Two real builds
must establish byte repeatability for a recorded environment; the recipe alone
is not evidence for every future Python/Git version.

The external manifest records source commit/tree, file Git IDs/modes/SHA-256/
sizes, payload digest, runtime dependencies, builder inputs and actual Python/
Git versions. It includes the final ZIP hash, never its own hash. SHA256SUMS
lists only the ZIP and manifest, never itself. Verification retains the recorded
build-runtime strings while checking bound tool inputs and archive recipe.
Hashes detect changes against a trusted value; they do not authenticate the
unsigned publisher. Actual Explorer, Windows/Calibre, rendering and anonymous
download checks remain mandatory at their separate task boundaries.

Targeted tests use real, authored temporary Git repositories and mutated assets:

```powershell
& $BuildPython -I -B -m unittest discover -s (Join-Path $SourceRoot 'tests\python') -p 'test_release_package.py' -v
if ($LASTEXITCODE -ne 0) { throw 'Package tests failed' }
```

The shared Python runner discovers these tests automatically and binds the
release tools into its source-change guard. Its Python child has a bounded
600-second deadline, retaining nonzero timeout/failure behavior. The prior
M5-T01 300-second local timeout remains historical failed evidence.
