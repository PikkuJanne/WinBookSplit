# V101-12 - reviewed integration and frozen release assets

Task 12 is DONE at integrated source **6f2c2e5d4d5fc484bcb1530474f09957e09775ef**, tree **5bc1a6e045e6abfc04689087bb0e25692ab9dd08**.
[PR 31](https://github.com/PikkuJanne/WinBookSplit/pull/31) merged normally. The
[completed main CI](https://github.com/PikkuJanne/WinBookSplit/actions/runs/38170965005) and [strict receipt](V101-12-main-ci.json)
record actual hosted Windows Server checks for this source. They are separate
from native final ZIP Windows 10/11 acceptance.

The [completion record](V101-12-review-freeze.json),
[freeze receipt](V101-12-release-freeze.json) and
[production expectations](V101-12-production-expectations.json) bind the exact
three existing assets and 33 allowlisted Git R members. The unsigned ZIP SHA256 is
`747d578209b3b3e84734620c3e071af828dbe038b55af779ace8508dbb7a6b3a`.
The existing producer/frozen verifier supplied those bytes; no manual ZIP edit
or working-copy payload replacement occurred. R and these asset identities
remain fixed through native acceptance and publication. Both actual bare
CPython 3.14.8 builder `build` and `verify` invocations returned exit 0 and
`PACKAGE_PASSED`, for each of the two exclusive asset directories. A held
read inspection confirmed identical three-asset sets and exact Git R members.
This is same-tool repeatability, not a cross-toolchain equality requirement.

The fresh [security receipt](V101-12-security-refresh.json) and
[review notes](V101-12-security-refresh.md) retain the integrated-R recheck from
20:10:47-20:10:59 UTC. All 41 current primary requests returned HTTP 200, with
no new material finding. The fixed Calibre route evidence remains finite static
inference, not a full binary call graph or security guarantee. Process jobs do
not provide memory/disk/network sandboxes. Recheck primary sources again at
publication. Earlier Task 10 reviews remain unchanged historical evidence.

Copied JSON evidence is written as explicit LF. Its linked byte hashes identify
the saved files; distinct collected-raw hashes and exact CRLF-to-LF transforms
remain recorded. Earlier 087 failures/corrections and historical v1.0.0 evidence
remain preserved. Review used a separate context whose author contributed
some earlier code, and is not an independent security audit.

Task 13 is IN_PROGRESS. Final ZIP native Windows 10/11, real ebook validation on
the final package, public tag/release and anonymous downloads remain NOT_RUN.
No owner session or new tag/draft is required by this checkpoint. The same ZIP
must now pass both available native Python branches and the reviewed PS7 path;
the Windows10 owner fallback is one consolidated command/report only after all
other independent work is complete.
