# M4-T03 regression and fault injection

Run the shipped test entry point with the explicit isolated developer Python:

```powershell
& $TestPython -I -B $Runner --layer faults --shell-path $PS51 --shell-path $PS7 --report $NewExternalReport
```

Both actual supported PowerShell executables are mandatory. This targeted route
executes the existing path, process and outcome suites once, then the new fault
stage. Full includes the new stage as number 21 and already runs those three
required error suites once. A direct call to `characterize_faults.py` establishes
only its narrower combined-isolation, callable-invariant and source-binding
observations. AC-075 also requires the same-source path/process/outcome reports;
four captured-source controls alone cannot establish the full error suite.

AC-073 adds five actual engine worker roles: two prior repeated publications,
two overlapping runs and one repeated publication after failure. Both overlapping
workers must hold their real marked stages after the first completed slice.
The failing worker writes a real partial second slice and raises an injected
write error while the other worker remains held. The parent then compares the
held stage's ledger, file identities and hashes, source, neighbor and prior runs,
verifies cleanup affected only the failing run, and releases the successful
worker to publish. Reopened outputs must contain every expected physical page
exactly once. Exact run IDs, PIDs, barriers, stream/native receipts and bounded
failure records are required; a matching total alone is insufficient.
The Windows venv launcher and actual interpreter are recorded separately. A
boot acknowledgement, retained worker handle and membership in the exact owned
job bind the interpreter; matching native statuses, an empty job and pipe EOF
must be observed before cleanup.

AC-074 uses seed `20261010074`: 240 manual cases and 240 independently specified
hierarchies checked in both bookmark levels give 720 accepted plan observations.
Per-page ownership labels are constructed before arranging the outline, then
run-length encoded independently of planner helpers. Front matter, parent opening
and fallback ranges, aliases, shuffled nodes, unusable/outside children and depth
limits remain explicit. Another 160 manual tokens and 96 malformed outline
observations must reject with declared diagnostics. Twelve real callable writer
samples reopen physical markers and bind source, neighbor and prior identities.
These are callable/PDF writer checks, distinct from native PS/BAT observations.

AC-075 combines the existing real-engine and authored fake-child error suites
under PS5.1 and PS7 with four new captured-source controls: replace or delete an
authored processing copy after planning, then execute the captured original
snapshot through each actual host. The copied engine adapter and its exact byte
changes are disclosed; it denies later reads of that processing source while
allowing required staged-output reopens. Original fixtures, neighbors and prior
publications remain immutable. Native outputs must retain the original captured
hash and page markers. Existing paths cover literal hostile names, providers and
held/unreadable inputs; process and outcomes cover simultaneous stream floods,
fast exits, terminal records, nonzero native statuses, cancellation, timeouts,
owned descendant shutdown and unrelated-process preservation. Cleanup/junction
checks remain in the output/path suites and the full route.

Strict independent validators require exact cases, argv/raw stream hashes,
native status, timing/readiness, source maps and identity-bound preservation.
Synthetic receipt mutation tests verify rejection of forged or incomplete claims;
they are not native passes. Process assertion failures now preserve the last
attempt, raw host/payload/PID-file observations and authored workspace. Retained
historical failures are not relabeled or diagnosed from a later passing retry.

Only original synthetic fixtures are used. Injected faults do not fill a physical
disk, test power loss, certify arbitrary hostile PDFs or establish a same-account
security sandbox. Human M3-T01 remains closed. Explorer, clean OS, CI, package and
release checks remain separately recorded gates. Raw commands, paths, PDFs and
receipts stay outside Git; public task evidence contains reviewed summaries and
byte/hash seals only.
