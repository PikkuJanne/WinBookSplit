# V101-06 diagnostics checkpoint

Ordinary console dependency failures now show one fixed reason, one explicit
action and the owned local log path. Full original exception/probe records stay
in the existing bounded console log and finalized run manifest. Preview and
NonInteractive preserve their structured console records and early no-disk
behavior. Record creation/finalization failure preserves the primary exit,
prints the full diagnosis and retains partial files. Processing never installs
dependencies.

Guidance distinguishes absent interpreters, unsupported branches/flavors,
security floors, package absence/version/origin and converter failures. Python,
pypdf and converter text comes from the shared policy. Broken application
environments receive explicit Repair guidance; their interpreter and preserved
backup paths are not recommended as setup bases. Unproved process shutdown
takes the preserve-and-inspect action before retry advice. Existing outcome
codes/status/schema, null accepted versions, source metadata and support-export
allowlists remain enforced.

The pre-change actual inbox5.1 reproduction exited3 with the giant dependency
frame, no early log and obsolete direct pip advice. Its receipt is retained.
The frozen work follows base commit `cd4e9795e0786069348a27e8fa176912f979b636`,
tree `0f95e05273ed493c29faad253167919aa785c638`; raw changed-file hashes are in
[the redacted receipt](V101-06-diagnostics.json). It binds the actual commands,
source snapshots and local raw-report hashes, rather than naming its own future
commit. Main SHA256 is
`dfb946eff564805cfc9037637a01205eb865bd30d1ed5d2dc7bccdd5d662158d`;
Runtime is `abd6a60f1cc6d5dd33888396cf7807fc39272c199560ff6ec52cce68d4093b3d`.

Actual available-host results:

| Check | Result |
| --- | --- |
| Focused Outcome/Logging/Support Pester | 58/58 on each actual host |
| Shared full shell regression | 182/182 on each actual host |
| Skipped/not-run/failed containers | Zero in both focused and cumulative runs |
| Syntax/scoped PSSA | Zero errors/findings; legacy Main findings remain observations |
| Full native diagnostic cases | 32/32; exact primary native exits3/130 preserved |
| Available OS | CIM Microsoft Windows11Pro,10.0.26300/build26300,ProductType1,64-bit OS/process |
| Actual hosts | Inbox5.1.26100.9444/CLR4.0.30319.42000 and isolated7.6.6/.NET10.0.12 |
| Actual Python bases | Regular x64 3.13.16 and3.14.8 |
| Source/neighbor/prior output | Hashes, size, identity and mtime unchanged |
| Cumulative owned test work | Removed only after its successful guarded run |

The native cases exercise both bare bases, explicitly authored wrong-version
and outside-origin package stubs, missing explicit paths, full Main messages,
owned logs/manifests and actual redacted support exports. Stub-package controls
are not upstream pypdf tests. Actual deny-Write ACLs verify fallback; original
owner/group/protection/effective ACE permissions are restored. Main+VERSION-only
copies answer Version/help without dependency helpers under both hosts.

Six explicitly recorded copied-source fault controls use actual owned-process
cancellation/deadlines and a closed-writer finalizer fault. They prove original
codes3/130 survive the secondary failure, with native stop records retained;
they are not unmodified-package end-to-end acceptance. Synthetic3.14.2,
future-patch/flavor/architecture and unsafe-stop flags are labeled validator or
formatter controls, not installed-runtime/process-stop observations.

The initial focused pair failed53/57 on each host: one real omission of the
policy-derived package version and three Pester6 assertion-authoring defects.
The omission and matcher spelling were corrected; safety assertions remain.
The initial native wrapper omitted Diagnostics before Process and failed before
any Main child. Both attempts and the separate static-wrapper preparation
failures remain local historical evidence, not passes.

Final cumulative report SHA256:
`a365563434de8794ba66c8ca3856a2486becda64fc532fce721a1214017f03cc`.
Native report: `41369b402333f6331513d8916ab2ee786b1de888c2626f6d1f82693cbdc86584`.
Final index: `c9bf3f6f4aaf32911ccff36e7054f724db2b7e11f98f9da1eba3476a876b6c23`.
The [independent API review](V101-06-api-review.md) found no additional
Windows11-only API or blocking diagnostic defect; static compatibility does not
replace native execution. The separate [setup CI receipt](V101-05-ci.md) certifies
only the earlier committed setup checkpoint.

AC034 is closed here only for actual available Windows11 inbox5.1 parsing and
entry-point behavior plus unchanged BAT bytes. AC036 verifies correct native
OS recording on this workstation. Their native Windows10 portion remains
NOT_RUN and is retained in mandatory AC068. No final ZIP, Windows10, Explorer
or real Calibre conversion pass is claimed. No owner action was requested.
Next: V101-07 full reference Python and actual two-minor application matrix.
