# Local verification

Use the explicit fresh developer-venv interpreter from
`docs/codex-v1.0.0/SUPPORT_AND_SETUP.md`. Run the same entry point from any
working directory; it resolves the trusted repository from its own location.
Do not activate a venv or install dependencies globally.

```powershell
$TestPython = 'C:\Tools\WinBookSplit\.venv-dev\Scripts\python.exe'
$Runner = 'C:\Tools\WinBookSplit\tests\run_tests.py'
$ToolRoot = 'C:\Tools\WinBookSplit-test-tools'
$PS51 = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'
$PS7 = 'C:\Program Files\PowerShell\7\pwsh.exe'
$Report = Join-Path $env:TEMP ('WinBookSplit-tests-' + [guid]::NewGuid().ToString('N') + '.json')
& $TestPython -I -B $Runner --layer full --tool-root $ToolRoot --shell-path $PS51 --shell-path $PS7 --report $Report
if ($LASTEXITCODE -ne 0) { throw 'Local verification failed; inspect the report' }
```

Substitute actual absolute interpreter/tool/host paths. Every report must be a
new absolute path outside the checkout. Existing reports are refused, and
reports containing commands or local paths should stay outside Git. Commit only
the reviewed summary with hashes in `docs/codex-v1.0.0/evidence/`.

Use `--layer python`, `--layer shell`, `--layer manual`, `--layer bookmarks`, `--layer level2`, `--layer plan` or `--layer diagnostics`
for targeted checks.
The shell layer requires explicit hosts and the isolated exact-version modules;
the full layer must include both supported hosts for the milestone gate.
The manual layer executes all 22 corrected manual oracles, eight extra CLI cases,
250 seeded valid-start coverage checks, 25 real writer samples, safe import and
the corrected six-section Level 2 BM-03 launcher reference. Promised reports must include
all case records and preserve the immutable baseline, source, inputs/neighbors
and owned cleanup. When hosts are requested, the report must also contain all
six successful entrypoint records, including three concurrent launches.
It accepts the two `--shell-path` options for the controlled actual entrypoint
probes; omit them for a focused Python check and record entrypoint checks as not
run. The eight-stage full layer runs this manual regression route, corrected
Level 1/2 bookmark routes, shared-plan acceptance and diagnostics after the units and both shell stages.

The bookmark layer executes corrected BM-01/02/05/08 targets, five bounded
normalization cases, 150 seeded Level 1 plans and ten real writer samples. Its
new report must include every case, exact page identity checks and preservation/
cleanup results. This Python route receives no shell arguments; the full
layer's six existing owned launcher probes still exercise manual MAN-03 and
corrected Level 2 BM-03. See `bookmarks/README.md` for both levels' coverage and limits.
Historical M1-T02 evidence retains the three then-unchanged bookmark observations;
M1-T03 evidence retains its then-unchanged BM-03. Current historical comparisons
are empty after Level 2 corrections; their guarded earlier evidence is unchanged.

The level2 layer executes BM-03/04/06/07 actual CLI targets, six hierarchy
aggregates, 150 seeded parent/child plans and ten real writer samples. Direct
children remain inside their own retained parent intervals; front matter,
opening/fallback sections, aliases, invalid descendants and no-selected-level
rejection are checked. This Python route receives no shell arguments.

The plan layer checks complete shared partition validation, deeply immutable
preview metadata, and three actual manual/Level 1/Level 2 writer results against
preview filenames, order, ranges and exact physical page identities. It includes
300 seeded plans and deliberate changes/repointing/removal of owned source copies;
execution must preserve its captured original reader snapshot. Original fixtures,
inputs/neighbors, historical guards and owned cleanup are required. This route
receives no shell arguments. See `plans/README.md` for coverage and limits.

The diagnostics layer requires both actual shell hosts. It exercises 23 structured
engine results, zero-new-output failures, three successful mode controls and the
real category/choice handler under PS5.1/PS7. Explicit no-outline, no usable
destinations and no-Level2 categories offer only applicable choices; invalid/read/
write failures offer none. Eighteen malformed protocol records reject in each
host. Actual native argument quoting and the production execution function's
200K dual-stream/final-result handling are separate controlled probes. No top-level
UI or Documents processing is performed here. See `diagnostics/README.md`.

`--layer extraction` remains an explicit historical diagnostic comparing the
current engine with immutable original behavior. It is expected to fail after
the intentional M1-T02 manual fixes; it is no longer part of the full gate.
The historical AC-011 unit assertion instead reads the shipped engine blob at
M1-T01 commit `88c2149b3b3034fbd0d7ef23c2382f4b01648e4c`, preserving its extraction
evidence without requiring current code to retain known manual bugs.
`--layer baseline` still invokes the untouched historical harness and deliberately
refuses the modified PowerShell source after M1.
`--failure-probe native`, `--failure-probe python` and
`--failure-probe pester` deliberately fail and must return nonzero even after
the runner's later successful command/report step. These are runner acceptance
probes, not application defects. The Pester probe requires the shell arguments.

The shell bootstrap executes only the authored repository test script in a
fresh `-NoProfile` host, using an encoded command and trusted script block.
The child requests process-only `RemoteSigned` so Windows PowerShell 5.1 can
import the isolated tools when its default is `Restricted`. Machine/user policy
is unchanged; enterprise policies take precedence and a blocked import fails.
This developer invocation does not certify the application's ordinary launcher
under the workstation's policy.
Each child uses its own host's built-in module directory and the absolute
external pins, so an inherited PowerShell 7 module path cannot supply tools to
Windows PowerShell 5.1.

Fixture PDFs, direct-engine slices, Pester TestDrive files and child working
directories live inside unique temporary directories. Actual M1 entrypoint
probes use the application's observed Documents folder with separately reserved
GUID directories, ownership markers and synthetic neighbors. Their nonrecursive
cleanup verifies containment, marker, allowed files and absence of reparse
points before removing only the run's owned outputs. No private Documents
content is enumerated or used as input. See `extraction/README.md` for the actual
six-probe unrelated-directory/parallel method and its limits. `manual/README.md`
describes the corrected manual regression route and `bookmarks/README.md`
describes corrected Level 1 checks. Environments,
generated outputs and raw evidence are ignored; ignore rules do not replace
reviewing the exact staged paths.

The explicit legacy extraction layer compares immutable original Git bytes with
the current shipped engine using fourteen understood known-bad cases. Processing ASTs, exits,
stdout/stderr, filenames, exact page identities and PDF bytes must match, and
engine import must have no processing/configuration side effects. A pass means
mechanical equivalence, not repaired splitting; current corrected manual code is
expected to fail this diagnostic. The historical baseline layer
continues to require the original launcher hashes without exceptions.
The shell layer gates syntax and the new scaffold's selected static checks;
unchanged application analyzer findings are reported as legacy observations.
The manual route verifies corrected manual splitting directly against synthetic
PDF page identities. Controlled actual M1 entrypoint splitting under both shells
and batch verifies engine resolution and selected paths. Level 1 regression
checks exercise corrected front-matter/outline planning directly. Level 2 checks
exercise parent boundaries and selected-level rejection.
Shared-plan checks exercise the callable preview and captured-source writer.
Full UX/Explorer, Calibre and the
extracted release package remain separate mandatory later checks. CI installation and workflow
execution belong to M4-T05; this local scaffold does not claim a CI run.
