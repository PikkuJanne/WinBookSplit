# M2-T05 process acceptance

Run with the explicit isolated pinned developer interpreter and both actual hosts:

```powershell
& $python -I -B tests/run_tests.py --layer process --shell-path $ps51 --shell-path $ps7 --report C:\external\new-process-report.json
```

`characterize_process.py` requires actual Windows PowerShell 5.1 and PowerShell 7.
It exercises the shipped direct-executable supervisor with 26 authored native
controls: simultaneous two-MiB stdout/stderr, fast exit/blank lines/no-newline tails,
UTF-8 byte boundaries, literal empty/quoted/backslash/hostile argv, a separate
100 KiB result frame, over-budget/duplicate/absent/malformed records, and finite
parent/grandchild timeout, injected-token cancellation, inherited pipes after
parent exit, and a detached inherited-pipe child. Native status, exact byte totals,
bounded tails, job assignment, EOF and independently queried authored PIDs are
required. An unrelated independently launched process must remain alive and is
then terminated through its retained `Popen` handle. This fixture launches the
actual ordinary base interpreter directly, checks authored self-PID/executable
readiness against the retained PID and independently observes alive before/after
supervision, then stopped after retained-handle termination.

Nine actual whole-application processes use PS5.1, PS7 and unchanged BAT: dual
flood and fast-tail failure controls substitute only an engine in an owned copied
application; three real generated Unicode/hostile bookmark PDFs exercise the
unmodified engine. The copied launchers/helpers retain actual reviewed bytes.
Literal environment-variable wrappers avoid another expansion of supplied BAT
paths. Native failure must remain nonzero, final stderr must survive, console logs
remain below 400 KiB, and actual child argv/UTF-8 receipts must match. Logs
separate the terminal provisional finalizer record from exact raw stream
tails. That record must equal the sole final stdout outcome and preserve the
actual native exit and validated engine result; stream byte limits remain exact.
Real PDF
outputs are reopened for all three physical page IDs/content, exact Unicode/title
filenames, manifests and source/neighbor/read-only preservation. Exact names come
from the destination-bound preview with independently measured complete path and
title budgets, preserving all original title data. This also runs beneath the
full runner's longer nested TEMP paths. Only emitted
authenticated owned output/log paths are removed; private Documents is not listed.

The independent `validate_process_report.py` requires exact cases and contradicting
or missing promised evidence fails the runner. Synthetic validator unit tests are
structural checks, not Windows passes. Pester includes twelve focused actual native
checks in each host. Source-byte digests, actual host versions, policies and native
commands are recorded. The full runner retains earlier regression layers.

Injected cancellation does not certify human Ctrl+C or Explorer. This is a controlled
workstation test, not a clean OS, release-package test or sandbox claim. Children
are finite authored programs; no private document or external command text is
executed. Failed/development attempts must remain separate external receipts.
