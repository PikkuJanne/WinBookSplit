# Checkpoint 5 — Verified public v1.0.5

[WinBookSplit v1.0.5](https://github.com/PikkuJanne/WinBookSplit/releases/tag/v1.0.5) was published on 2026-11-12 13:24:06 UTC.
GitHub release 442980603 is public, current, and marked as a stable release.
The annotated tag object `23b3afb09a41610d595f77d2d685a731bf85a2f5`
peels to the frozen source `56dc118313cd610bde4cbb578f46e9acbb273610`.
[Exact package acceptance](CHECKPOINT_4.md) records the Windows 11 results.

All four jobs in [Windows CI run 43223083591](https://github.com/PikkuJanne/WinBookSplit/actions/runs/43223083591) passed
on the frozen source. Windows Server 2025 ran the shared Python regression,
PowerShell 5.1 and PowerShell 7 syntax, scoped analysis and Pester, plus the
application matrix on Python 3.13.16, 3.14.8 and 3.15.0. All four downloaded
report sets passed the source-bound receipt validator.

The three assets were verified in the draft before publication. Anonymous
downloads from the public release matched every expected byte length and
SHA-256 and passed the package verifier against the tagged source.

| Asset | GitHub asset ID | Bytes | SHA-256 |
| --- | ---: | ---: | --- |
| [WinBookSplit-v1.0.5.zip](https://github.com/PikkuJanne/WinBookSplit/releases/download/v1.0.5/WinBookSplit-v1.0.5.zip) | 739581086 | 1,133,973 | `756a467136f7c41a70fbdbcfcfcf073a772bc338a10ce2141fe2669f396431c3` |
| [release-manifest.json](https://github.com/PikkuJanne/WinBookSplit/releases/download/v1.0.5/release-manifest.json) | 739581087 | 17,778 | `a092723d2dfab81f7ffc1a40a4a2fef48427f2fa25ef9a6134da4360a4c46bce` |
| [SHA256SUMS.txt](https://github.com/PikkuJanne/WinBookSplit/releases/download/v1.0.5/SHA256SUMS.txt) | 739581089 | 178 | `22b80f6647bea8234bcbef4ca4f35382d209065c10661a98038ed54eab883625` |

Live tag objects, peeled commits and release metadata for v1.0.0 through
v1.0.4 matched the prepublication records. All 15 earlier assets were
downloaded again anonymously and matched their prepublication SHA-256
values and lengths.

The local publication receipt has SHA-256 `9787901e710dc11fd909629a0a1cd9432caf101d48d3576bbb9abea2d5ee705e`.
The application scripts are unsigned. Windows 10, ordinary Store first-run
setup, and genuine XFA rendering remain unavailable for acceptance.

The frozen source was integrated into main by fast-forward. This closure
records evidence after publication. The annotated tag points to the tested
source commit.
