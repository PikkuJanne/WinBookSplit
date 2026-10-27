# Build and inspect a WinBookSplit v1.0.3 package

The owner's latest [27 October decision](../../docs/codex-v1.0.3/OWNER_DECISION_20261027.md)
closes further owner work. Earlier mandatory-owner/policy wording below describes
the original gate contract; unavailable final-ZIP ownerStore and genuinecompany-
policy scenarios now remain NOT_RUN/owner-waived with disclosed release limits.
All agent/source/CI/package/anonymousbyte/main verification remains mandatory.
Do not request another owner run or change the native checker to invent a pass.


This developer tool builds three local assets from an explicit full commit:
`WinBookSplit-v1.0.3.zip`, `release-manifest.json` and `SHA256SUMS.txt`.
It does not create a tag, upload, sign or publish. An unpublished candidate is
separate from source freeze, exact-package native acceptance and public release
verification in the [v1.0.3 release runbook](../../docs/codex-v1.0.3/RELEASE_RUNBOOK.md).

Run with a trusted Python, Git and the builder files from the intended source
commit. The builder, `payload.json`, reused CI input checker and
`frozen_assets.py` must match their committed bytes exactly. A checkout that
converts these files' line endings must use an ordinary Git source
archive/extraction with original bytes for the tooling; the source repository
argument still names the Git root. Changed or untracked tooling is refused.
Payload bytes are always read from Git blobs with replacement refs disabled;
no working-tree file is adopted. Do not weaken those guards or change global
Git configuration to make a build pass.

`payload.json` explicitly lists 34 files. The shared checker validates runtime
import/helper dependencies and fixed hashed runtime requirements. The package
includes the processing BAT/PowerShell entry points, setup BAT/PowerShell entry
points, shared
compatibility policy/helpers, splitter/converter code, current setup/fidelity/
ebook/support docs, license/notices and intended icon. Tests, development tools,
native acceptance kit/fixtures, handoff history, poster, venvs, caches, documents
and runtime/converter binaries are excluded. Relative document links must
resolve in the payload; source-only references use GitHub links.

Use a fresh **absent**, absolute output directory whose parent already exists
outside the Git checkout. Existing destinations are refused, including junctions,
symlinks, UNC and device paths. No recursive cleanup occurs. A failed write or
validation keeps `.package-incomplete` and diagnostic artifacts; verification
rejects the marker. Preserve failed work for diagnosis and choose a new path.
The tool is not a sandbox against another process running as the same account.

```powershell
$BuildPython = 'C:\Tools\Python3150\python.exe'
$SourceRoot = 'D:\projects\WinBookSplit-main'
$Builder = Join-Path $SourceRoot 'tools\release\build_package.py'
$SourceCommit = (& git -C $SourceRoot rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0) { throw 'Cannot resolve the intended committed source' }
$CandidateAssets = 'C:\LocalBuilds\WinBookSplit-v1.0.3-candidate-01'
& $BuildPython -I -B $Builder build --repo $SourceRoot --commit $SourceCommit --output $CandidateAssets
if ($LASTEXITCODE -ne 0) { throw 'Package build failed' }
& $BuildPython -I -B $Builder verify --repo $SourceRoot --commit $SourceCommit --output $CandidateAssets
if ($LASTEXITCODE -ne 0) { throw 'Package verification failed' }
Get-FileHash -LiteralPath (Join-Path $CandidateAssets 'WinBookSplit-v1.0.3.zip') -Algorithm SHA256
Get-FileHash -LiteralPath (Join-Path $CandidateAssets 'release-manifest.json') -Algorithm SHA256
```

Check the source SHA before building: `HEAD` here is resolved by the operator,
and the builder accepts only the full lowercase commit ID. Final builds must
use the frozen source commit **R** containing VERSION `1.0.3`, as required by
STORE-031. T06 freezes the candidate on the scoped branch; T08 must preserve those
accepted bytes when bringing R to final main. T07 tests that exact ZIP on the
owner's existing Windows 10 Store installation through automatic discovery,
normal setup BAT and ordinary application BAT without a supplied or private
runtime. Independent Windows 11 and authored EPUB/AZW3 regression, Internet-marked
script behavior and enforced-policy checks retain their separate evidence scopes.
A source test or earlier candidate ZIP does not satisfy those native gates.

ZIP entries use sorted paths, no directory entries, stored bytes, a fixed
1980-01-01 timestamp and regular-file 0644 attributes. Independent verification
checks every entry and raw byte against the commit, regenerates the expected
archive recipe and manifest, and verifies both checksum lines. Two real builds
establish byte repeatability for a recorded environment. Different Python/Git
toolchains are not promised to emit byte-identical archives.

The external manifest records source commit/tree, file Git IDs/modes/SHA-256/
sizes, payload digest, policy-bound runtime dependencies, builder inputs and
actual Python/Git versions. It includes the final ZIP hash, never its own hash.
SHA256SUMS lists only the ZIP and manifest, never itself. Verification retains
the recorded build-runtime strings while checking bound tool inputs and archive
recipe.

Runtime acceptance and build-tool observations are distinct. Python metadata
comes from the selected commit's strictly validated compatibility policy and
records accepted minor branches, numeric compatibility/security floors,
recommended patches, blocked patches, exclusions and the exact policy hash.
The source-test targets 3.13.16, 3.14.8 and 3.15.0 are not patch-equality gates.
Later stable patches are accepted numerically only within these reviewed branches;
unlisted minors are not yet validated and are not labeled insecure. pypdf
6.21.3 and its reviewed wheel SHA remain separately controlled, and the selected
requirements must agree with the policy. Calibre remains separately installed.
The manifest's builder Python/Git strings identify what actually built those
bytes; they do not narrow application acceptance to that Python patch.

`frozen_assets.py` also provides a Git-free, pure-stdlib verifier for the separately
pinned native acceptance kit. It checks all three external asset hashes before
parsing or returning validated payload bytes, and never executes application
payload code. The native kit records its own helper/fixture identities separately
and uses the existing guarded extraction/process APIs. It does not become an
application payload or an alternative release builder.

After acceptance and publication, use the named `WinBookSplit-v1.0.3.zip` from
the accepted public release;
GitHub's automatic source ZIP or tarball is a source checkout, not the tested
runtime distribution. Setup uses the BAT or complete OS-known host invocation
documented in [SETUP](../../docs/SETUP.md), with process-only policy where permitted;
enforced organization controls and unsigned-script limits remain binding.

Hashes detect changes against a trusted value; they do not authenticate an
unsigned publisher. Exact-package Windows/Calibre checks and anonymous public
download/byte verification remain required by the runbook. Historical v1.0.0
through v1.0.2 build evidence, tags and assets remain unchanged. The v1.0.3
candidate is unpublished while the existing-Store and remaining policy gates
are blocked; building it does not authorize publication.

Targeted tests reuse real authored temporary Git repositories and deliberately
mutated assets:

```powershell
& $BuildPython -I -B -m unittest discover -s (Join-Path $SourceRoot 'tests\python') -p 'test_release_package.py' -v
if ($LASTEXITCODE -ne 0) { throw 'Package tests failed' }
& $BuildPython -I -B -m unittest discover -s (Join-Path $SourceRoot 'tests\python') -p 'test_frozen_assets.py' -v
if ($LASTEXITCODE -ne 0) { throw 'Frozen-asset tests failed' }
```

The shared Python runner discovers these tests and binds release/compatibility
tools into its source-change guard. The full reference Python child has a finite
1200-second deadline with nonzero timeout/failure reporting. Earlier 300- and
600-second observations remain historical evidence. Test commands alone do not
claim a run; release evidence records actual results and exact source identities.
