# V102-08 required CI target and source controls

The existing application CI matrix now requires regular x64 3.15.0 beside
3.13.16 and 3.14.8. The separate 3.14.8 fixture driver and runtime-only
application environments remain. Existing action pins, contents:read,
bounded jobs, source/report gates and authored artifact allowlists are retained.
The report validator now requires an exact target from the typed, unique,
canonical source runtime-lock; it still rejects substituted versions,
skipped/partial/failed suites, wrong source and workstation-as-Server evidence.

The existing synthetic report controls cover all three exact targets. The
frozen metadata guard deliberately lowers the retained 3.14 recommendation
below its floor instead of making a no-op mutation of preferred 3.15's zero
floor. Exact payload count 34 matches the separate reviewed addition of the
setup BAT to package/checker inputs; no other member was added here.

The [source receipt](V102-08-ci-source.json) binds source base
`3cbc35e5a1636a4eefc481d3a0bf3cfb1665710d`, scoped actual bytes, commands,
raw receipt hashes and source review. Real Python 3.15.0 and developer 3.14.8
each passed all 7 report-control methods and all 15 frozen-asset methods.
Both final runs kept all 11 CI/package source paths unchanged, including the
actual payload manifest, checker and release-package guard. The parent's
additional package-input guard hash is recorded without claiming this agent
ran that separate suite. No skips were reported. Initial passing runs had a
runner-map omission; they remain distinct from the complete-binding reruns.

After those focused runs, the package checker gained validation of the setup
BAT's sibling-script dependency graph. The final focused receipts bind checker
`6693f07b19124598ac45038828f4d4e7e0ccfc164fd7556efb4685ea7b899173`;
the later source snapshot records checker
`e9cd7ac6238fcd826b9deb01cd96915bcc7a8ff5442e8d1b2d3252d8944d4381`.
The source snapshot is distinct from the tested binding. Focused package checks
and the eventual committed full CI must verify that graph change; the earlier
passes do not cover it.

Before correction, actual 3.15 reproduced the existing no-op metadata failure.
New synthetic application controls also reproduced rejection of 3.15 by the
old minor guard and rejection of an unpinned old patch only at a later mismatch
check. Those failures and exact source identities remain in the receipt.
Independent source review found no blockers. Structure and the per-command
CRLF-aware whitespace check passed.

Fresh official action tag checks matched every retained full pin. The official
[Actions Python catalog](https://github.com/actions/python-versions/blob/main/versions-manifest.json)
lists all three regular x64 targets as stable, and the new
[3.15.0 Actions release](https://github.com/actions/python-versions/releases/tag/3.15.0-38076878863)
provides the matching ordinary Windows x64 artifact. Raw catalog/API hashes
and the API-declared asset digest are recorded; this availability check did
not download or execute that artifact.

This is source evidence prepared for C1; no C1 commit or clean freeze is claimed.
Task8 remains in progress:
actual required hosted CI on committed source, the full suite and independent
native-kit preflight still need their own results. No Windows Server result
substitutes for mandatory final-ZIP Windows 10/11/Python 3.15 acceptance.
