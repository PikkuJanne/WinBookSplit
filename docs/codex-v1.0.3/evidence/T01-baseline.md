# T01 read-only baseline and published-byte comparison

Status: PASS for this baseline sub-scope; this is not application or native Store acceptance.

Source inspected: `4cc6b00b99685607f3214e731edaaec7d65926df`, tree
`8a5a246d2d18b3e2b3fd746a9347f77fb95f6e08`, scoped branch
`codex/v1.0.3-store-discovery`. Live `origin/main` matched that source at collection.
The only existing worktree was retained. The JSON records concurrent authorized
routing, handoff and regression-test edits rather than claiming an unchanged
worktree after those edits began. No production-file edits were made here.

`T01-baseline.json` records the collection time, exact live refs, old tag object
and peeled commit identities, release bodies and UTF-8 text hashes, asset IDs,
names, sizes, API digests and anonymous downloaded digests. The live release set
was exactly `v1.0.0`, `v1.0.1`, `v1.0.2`; no `v1.0.3` release existed. All nine
existing assets were downloaded using `urllib` without an authorization header,
token or authenticated cookie jar. Every downloaded size and SHA-256 matched the
API metadata. Release downloads naturally change download counts; the immutable
comparison deliberately excludes that counter.

The anonymously downloaded `WinBookSplit-v1.0.2.zip` SHA-256 is
`37d017e2a7bf704676b637f89239ec2cbaaa67e76f190f3dbbd17fec87e5092e`.
All four incident fingerprints match exact bytes from the current worktree,
raw Git HEAD blobs, raw `v1.0.2` blobs and the corresponding published ZIP members:

| File | Captured / all four source SHA-256 | Measured line endings |
| --- | --- | --- |
| `WinBookSplit.ps1` | `7d62bfd45ff312248bf215f2b6c29de93dd6b3e7a74eebbeb1bd7a3d0a3b398a` | 1,179 CRLF; no bare LF |
| `WinBookSplit.bat` | `c48e10b06e1177c723488944bcc0280ff4b51d1cc6d2f138206e6e3cab482893` | 29 CRLF; no bare LF |
| `engine/WinBookSplit.Runtime.ps1` | `88f452bf14b75a6bab2883e945ff9bd2ff3b780fb3813083c2a8f239db7d5796` | 726 LF; no CRLF |
| `engine/WinBookSplit.Compatibility.json` | `57811a77c837bd891750ea126d19c876e0ad7099793b490445d126720baf4582` | 163 LF; no CRLF |

No line-ending conversion was needed to match these fingerprints. The collector
also records LF-normalized and CRLF-normalized hashes, empty Git attribute output,
system `core.autocrlf=true`, unset `core.eol`, and absence of root `.gitattributes`.
Those configuration observations do not substitute for the measured exact bytes.
Raw Git bytes were captured through a binary subprocess pipe, avoiding PowerShell
text-redirection re-encoding. The entire collector source and actual command
vectors are retained in the JSON for reproducibility.

Collection used bundled developer Python 3.12.14, PowerShell 7.6.5, inbox
PowerShell 5.1.26100.9711, Git 2.56.0.windows.1 and GitHub CLI 2.97.0 on Windows
x64 build 26300. It did not invoke the desktop's Python aliases, install or change
Python/manager/packages, alter PATH or alias settings, create an application venv,
process working documents, mutate Git refs or edit any release. Downloaded public
assets and collector code remain in a uniquely owned local temporary directory;
private owner files and account names are not included in the public evidence.

The original PS5.1 nested Python collector remains `COLLECTOR_FAILED`, and pypdf
in the captured compatible base remains `UNKNOWN`. Only supplied sanitized
incident material was examined. The owner's Windows 10 Store gate is `NOT_RUN`;
this baseline does not satisfy `STORE-037` through `STORE-042`. Subsequent T01
behavioral regression results and routing/import review are recorded separately.

No commit, push, reset, stash, deletion or publication was performed by this
read-only baseline subtask. The main task owns its checkpoint and synchronization.
