# v1.0.3 release closure

**Complete, with disclosed owner-waived manual scenarios.** The owner closed further
manual work on 27 October 2026. No further owner test is requested.

[Public release](https://github.com/PikkuJanne/WinBookSplit/releases/tag/v1.0.3)

The annotated `v1.0.3` tag points to source `7c5f7afeef9b09733ee8853d6df5d86c42d9c726`,
tree `9d321cf13be478638d2e864c1fe1724a63bd8902`. The release ZIP has 938,143 bytes
and SHA256 `ba80fc35e0262f7f422dc1596d887e4a8207ec4ede94b754800b994084da4d20`. The companion JSON records all
three asset identities, raw/public receipt hashes and source bindings.

## Repair and dependency

The project and distribution now select official pypdf 6.21.3, policy version 5,
with the reviewed exact wheel hash. The Store registration query, package identity
and installed-interpreter discovery repairs are included. Python and Calibre
security floors, isolated dependency-origin checks and engine/setup/process/file
guards remain intact.

Small reduced-limit cross-reference stream regressions fail as expected against
6.21.1 and pass against 6.21.3 on the three supported Python targets. The selected
version is outside all seven reviewed affected ranges in the current official
advisory review. This does not establish a sandbox or a universal guarantee for
hostile documents.

## Actual checks and source scope

| Evidence | Actual result and scope |
| --- | --- |
| Final source H CI, run [40532432470](https://github.com/PikkuJanne/WinBookSplit/actions/runs/40532432470) | All four jobs passed; 715 Python tests; 252 Pester tests in each hosted shell; three application targets, each with 150 runtime units and eight PDF cases. All four downloaded artifact validators passed; all six report maps match 225 raw Git H hashes. Hosts were Windows Server 2025; no CI Calibre or owner Store acceptance is claimed. |
| Local source G, `74efe6b3f3da43053f2a67d2941db009c0be54ed` | On Windows 11: 713 Python tests; 252 Pester tests per shell; 35 syntax files, 34 scoped analysis files and 73 outcomes per shell. Shells were inbox 5.1 and PowerShell 7.6.6. |
| Local supported-runtime application matrix | Three regular x64 Python targets, each with 150 runtime units, ten PDF/Calibre cases and 21 native probes. Actual Calibre 9.15.0 converted authored EPUB/AZW3 inputs. |
| Exact-ZIP-derived setup/BAT/remedy/MOTW checks | Four setup/remedy cases with 16 steps, three marked setup BAT routes with 21 steps, and two explicitly mocked policy-refusal controls passed. Independent audit: 1,432 checks passed. These are developer checks, with explicit-runtime/child-PATH/mock scopes recorded; they do not prove Store or genuine company-policy acceptance. |
| Ordinary processing BAT | The unmodified BAT/PS1 processed an authored ten-page PDF from an unrelated working directory using a healthy application environment, without runtime or environment overrides. Exit zero, empty stderr, three outputs and all ten physical pages verified. Sources, neighbors, environment, marks and existing outputs were preserved; only the two known owned test-output entries were added. |
| Source/package equivalence | Final H ZIP is byte-identical to G. All 34 payload members and actual native helper bytes match. The local source maps differ only in three explained CI-tool/test/doc files. Local results retain G scope; they are not described as H reruns. Independent H audit: 85 checks passed. |
| Frozen builds | Both complete asset sets match byte for byte. All four build/verify commands passed, including manifest, checksum and raw Git bindings. |
| Public verification | Anonymous HTTPS downloads of all three new assets match frozen bytes. The trusted frozen-source package verifier passed on those downloads. Release is public, non-draft and non-prerelease; published notes match reviewed notes. All nine old asset bytes, three old release bodies/asset records and six old tag references match the baseline. |

G's earlier CI workflow remains a recorded **FAILURE**: Gallery returned HTTP 404
for the locked PSScriptAnalyzer download, so hosted shell steps did not run. The
minimal fix uses the official PowerShell GitHub release for exactly the same
version and SHA256, preserving acquisition guards. Eight focused tests, a fresh
bootstrap with all 74 module-file hashes and actual final H CI passed.

## Limits and preservation

Fresh final-ZIP owner Windows 10 Store acceptance, genuine company-enforced policy
denial and unavailable physical Explorer actions remain **NOT_RUN / owner-waived**
under [the owner decision](../OWNER_DECISION_20261027.md). These limits are in the
public release notes. Original T07 remains accepted only for original R and its
exact old ZIP. Scripts are unsigned and organization controls still apply.

No private books were processed or uploaded. Desktop Python/Manager, global
packages, PATH, aliases and persistent policy were not changed. Frozen R/F,
owner contracts, accepted receipts, older release assets and tags are preserved.
There was no force push, settings change, spending, retag, rebuild or reupload of
published bytes.

## Normal use and closure

1. Download `WinBookSplit-v1.0.3.zip` from the public release and extract it into a
   new writable folder.
2. Run `Setup-WinBookSplit.bat` once to create the application's own environment
   from an existing compatible Python.
3. Run `WinBookSplit.bat`, or drag a PDF onto it. Optional EPUB/AZW3 conversion
   requires Calibre.

These are normal use instructions, not another owner acceptance request. A durable
copy of all three final assets was saved in a new Downloads folder; existing
installations and shortcuts were preserved.

Before writing this closure, actual clean local main H matched freshly queried
live main and the scoped branch H; the peeled release tag was also H. STORE-048's
recorded pre-closure PASS has that observed scope. The following C commit changes
only documentation and receipts, leaving tag/assets at H. C's own hash and final
clean local main/live branch equality are verified externally after ordinary
commit/push; this document makes no advance claim about that later check.

There is no pending owner work or T09 for v1.0.3.
