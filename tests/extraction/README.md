# Mechanical extraction evidence

`characterize_extraction.py` reads the original engine from immutable local Git
commit `0de84f367f9bd5ddfa3f408a9c29505d7a39633f`. Original launcher raw hashes must
match the historical `expected_original.json` guards; that file, the original
harness, batch launcher, fixture generator and corrected plan oracles must still
match their original Git bytes. The old harness is unchanged and must refuse the
modified PowerShell launcher. This separate route explicitly compares original
and extracted sources; it never labels current code as the unmodified baseline.

The three processing function ASTs must be identical. Each of the 14 understood
known-bad cases runs both engines against generated original PDFs. Exit codes,
stdout/stderr, filenames, every page identity/order and output SHA-256 bytes must
match. Inputs/neighbors/source hashes must remain unchanged. An isolated import
probe verifies callable functions with no processing, stdout/stderr, CWD files,
logger-level/handler or stdout-configuration changes. Passing this route proves
mechanical equivalence, including existing defects; it does not prove repaired
page coverage or release acceptance.

Use the explicit selected developer interpreter with `-I -B`, an absolute script
path and a new external `--report`. Add both actual host paths with repeated
`--shell-path` to exercise AC-012. The central `--layer extraction` runs this
route; `--layer full` explicitly includes it. `--layer baseline` retains the
historical guard and intentionally fails after extraction.

AC-012 invokes the actual PowerShell file in Windows PowerShell 5.1 and
PowerShell 7 and the actual unchanged batch launcher (which launches Windows
PowerShell). It supplies responses through stdin, selects the fresh developer
Python through controlled PATH and supplies each host's own built-in module
directory. There are no UI/path/host shims. Each invocation has an unrelated CWD
with a counterfeit engine that must remain unexecuted, and a distinct generated
PDF with a GUID filename. Three further actual invocations run concurrently,
sharing one run-owned TEMP directory containing the old engine filename sentinel;
that sentinel must remain byte-identical.

These are real current-workstation probes. The application still defaults to
the actual observed Documents folder. Each test atomically creates its absent
GUID output directory with an ownership marker and synthetic neighbor before
launching. Cleanup verifies resolved containment, exact GUID/marker, no reparse
points/directories and every allowed member before deleting only those direct
files and the owned directory. It refuses unknown neighbors or ownership/path
changes. It neither enumerates nor processes private Documents content. Public
reports replace the observed Documents path with a symbolic label.

Actual-entrypoint timeouts terminate only the tracked live parent PID tree with
the system `taskkill /PID ... /T /F`, then require a successful termination
command, bounded parent wait and both streams drained before cleanup. A timeout
still fails the test. If the parent already exited or termination/draining
cannot be verified, the run fails closed and preserves its marked GUID output
directory for inspection. No image-wide process kill or automatic deletion of
an uncertain directory is used.

No Explorer interaction, Calibre conversion, fixed-engine acceptance, renderer
fidelity or release-package claim is made here. Dependency discovery and process
supervision remain future runtime tasks; the narrow extraction preserves their
existing behavior.
