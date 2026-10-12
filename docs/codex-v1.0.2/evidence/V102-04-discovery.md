# V102-04 discovery and preserved environment upgrades — 12 October 2026

Six actual native cases passed on Windows 11 Pro x64 26300.9457 under inbox
PowerShell 5.1.26100.9444 and PowerShell 7.6.6. Source base is
`2c0a1f24717d6e7e448c55b00a5138df73b247d7`; policy hash remains
`57811a77c837bd891750ea126d19c876e0ad7099793b490445d126720baf4582`.
This task changed one existing test expectation and required no production
discovery/setup change.

An external finite recipe verified an independently supplied 35-file source
map, held source leases, copied the current 33-member payload into fresh owned
application directories, and reused existing setup, owned process supervision,
snapshots and native discovery-control helpers. The work is retained outside
the checkout; no source-ZIP installation or historical evidence was modified.
Exact script/receipt/source hashes, command aliases and redacted returned
reports are in V102-04-discovery.json; raw paths/probes stay private.

Each native host passed:

- Automatic selection of the actual regular 3.15.0 runtime with only its
  directory supplied on controlled PATH; selected Path/Version/Source bound to
  its actual probe, with no processing installation or new application venv.
- Existing CONTROLLED_NATIVE discovery helper with an actual authored py.exe
  reporting `Unknown option: -0` and exit 2; the independently validated direct
  3.15 PATH candidate still succeeded and the original listing failure remained
  recorded. This does not claim an installed classic launcher was tested.
- Fresh app-local 3.13.16 and 3.14.8 environments installed from the exact
  reviewed offline pypdf wheel. Both remained authoritative after PATH changed
  to 3.15; setup with a valid 3.15 base and Check were no-ops. Every environment
  byte and timestamp remained unchanged.
- An actual offline negative-fixture venv created using the existing trusted
  helper 3.14.7. Its authoritative failure named the local venv, observed 3.14.7
  and `python_security_floor`; it never fell through to PATH 3.15. Normal setup
  returned `setup_repair_required` and preserved the entire environment.
- Explicit regular 3.15.0 base plus Repair installed locked pypdf 6.20.0 into
  the original local venv path. The unique old backup preserved every previous
  file byte/timestamp; the neighboring sentinel survived. Repeated Repair was
  a no-op, retaining exactly one backup.

Application, setup and controlled-launcher child processes were checked for
owned jobs, stopped parent/descendants,
complete streams and finite deadlines. Application source hashes stayed fixed;
PATH and execution-policy observations were restored/unchanged. The PS7 report
retains the existing native-kit classifier's BLOCKED label because that
classifier requires inbox 5.1; separate host facts and passed cases describe
these as supplemental developer checks, never final-ZIP OS acceptance.

Before editing the current test, the focused inbox Pester invocation executed
its optional-launcher case under actual 3.15 and reproduced the obsolete
`^3\.(?:13|14)\.` assertion failure. One case ran and failed; 69 cases were
filtered/not-run and are not counted as passes. The test now records the
supplied interpreter-only version and compares the actual selected version
with that observation. It retains the original listing/exit-code assertions.

After the focused test correction, existing Discovery/Runtime/Setup containers
passed **70/70 on each native shell** with actual isolated runtime15, zero
skips/not-run/failed containers. Existing invalid overrides, shadow packages,
unsafe discovery sources/aliases, rollback, cancellation, owned-file and
dependency-origin controls remained exercised. Pester 6.2.0 and
PSScriptAnalyzer 1.25.0 came from existing reviewed external tooling; all three
shell test files had zero parser/analyzer findings and unchanged test hashes
during their runs.

Both first external recipe invocations failed before any work directory or
application execution because module autoloading was disabled and the wrapper
omitted the trusted Security module needed for Get-ExecutionPolicy. Adding its
absolute host-module import corrected only that external test wrapper. Those
initial failures remain failures and are recorded in the JSON.

The handoff structure validator returned STRUCTURE_VALID/exit 0. No final ZIP,
Windows 10 execution or publication is claimed here. Existing raw runtime
identity/errors satisfy this task's selected-runtime observations; concise
console rendering and executable emitted remedies remain V102-06 work. Parent
owns acceptance/status changes and the reviewed commit/push/live checkpoint.
