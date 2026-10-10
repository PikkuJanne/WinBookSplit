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
$Calibre = 'C:\Tools\Calibre\ebook-convert.exe'
$Renderer = 'C:\Tools\Poppler\pdftoppm.exe'
$RendererPython = 'C:\Tools\PDFium-dev\python.exe'
$Report = Join-Path $env:TEMP ('WinBookSplit-tests-' + [guid]::NewGuid().ToString('N') + '.json')
& $TestPython -I -B $Runner --layer full --tool-root $ToolRoot --shell-path $PS51 --shell-path $PS7 --calibre-path $Calibre --renderer-path $Renderer --secondary-python-path $RendererPython --report $Report
if ($LASTEXITCODE -ne 0) { throw 'Local verification failed; inspect the report' }
```

Substitute actual absolute interpreter/tool/host paths. Every report must be a
new absolute path outside the checkout. Existing reports are refused, and
reports containing commands or local paths should stay outside Git. Commit only
the reviewed summary with hashes in `docs/codex-v1.0.0/evidence/`.

M4-T01 adds `--layer fidelity` as the nineteenth full stage. Fidelity/full
requires both supported shell paths, the absolute development-only Poppler
`pdftoppm` executable and an isolated development Python with `pypdfium2`.
The application does not require either renderer. See `fidelity/README.md` for
generated fixtures, structural checks, exact same-renderer pixel comparisons
and preserved external PNGs for agent visual inspection. This is automated
native Windows evidence; it does not reopen or certify human testing.

Use `--layer python`, `--layer shell`, `--layer manual`, `--layer bookmarks`, `--layer level2`, `--layer plan`, `--layer diagnostics`, `--layer output`, `--layer paths`, `--layer conversion`, `--layer runtime`, `--layer process` or `--layer outcomes`
for targeted checks. `--layer cli` requires both actual supported hosts and the
real pinned Calibre executable; it adds M3-T02 CLI acceptance as stage 15 of full.
See `cli/README.md` for the 72 native controls, dependency-independent Version,
complete scripted modes, no-prompt/no-fallback behavior, plan-only PDF/ebook
previews and real EPUB retained-output example. Preview tests validate complete
physical coverage and safely cleaned temporary conversion independently of
execution results. Source/type/receipt unit checks remain distinct from actual
Windows acceptance.

The M3-T01 outcomes layer checks final fallback/exit mapping, owned-tree cancellation
and timeout, manifest/publish/cleanup failures and console-log finalization under
both actual supported hosts and BAT. Its separate OS console-signal controls use
hidden private consoles; human keypress/Explorer checks remain separate. The full
route includes this layer and inherited real Calibre checks. See
`outcomes/README.md` for controlled copied-application provenance and recovery limits.
Current regression helpers authenticate immutable historical Git inputs while
allowing the reviewed launcher to evolve; untouched baseline/extraction guards
still refuse current launchers that differ from their historical originals.
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
run. The thirteen-stage full layer runs this manual regression route, corrected
Level 1/2 bookmark routes, shared-plan acceptance, diagnostics, output transactions,
path acceptance, real conversion, dependency preflight and process supervision after the units and both shell stages.
Conversion/runtime/full requires an explicit absolute converter with the pinned Calibre
9.15.0 bytes. PDF-only targeted layers do not require or probe Calibre.

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
write failures offer none. Existing loose output files are preserved beside a
separately published successful run. Eighteen malformed protocol records reject in each
host. Actual native argument quoting and the production execution function's
200K dual-stream/final-result handling are separate controlled probes. No top-level
UI or Documents processing is performed here. See `diagnostics/README.md`.

The output layer checks four actual repeat/same-basename runs, four simultaneous
fixed-timestamp Python processes, five injected failure paths and eight ownership
cases, including two actual controlled Windows junctions. Complete manifests,
owner markers, PDF digest/size/count, exact physical IDs and content parity are
required; earlier runs, sources and neighbors must remain unchanged. Failure
results require no successful final and explicitly owned staging/diagnostic
handling. This route requires actual Windows and receives no shell arguments.
See `output/README.md` for the precise scope and recovery limits.

The paths layer requires both actual supported hosts. Three PS5.1/PS7/BAT
literal input successes and eight rejection probes check file-only/provider
validation, mixed extension case, special characters/Unicode, wildcard decoys,
correct metadata, parser rejection, actual held-unreadable inputs and read-only
source successes with restored attributes. Direct API writers reject
replanning/source reopens through installed execution traps. Nine title
cases and 120 real sections per mode check safe Unicode/fallback names and
chronological dynamic numbering. Destination-bound previews preserve exact
names within conservative UTF-16 budgets; impossible/changed bases, actual
held-unwritable bases and explicit ENOSPC failure retain zero published output.
PS uses external owned bases; BAT uses only its explicit authenticated Documents
run/log children. Stored policies, source/input/neighbor hashes and held cleanup
are checked. See `paths/README.md` for exact methods and limits.

The conversion layer generates an original offline EPUB and genuine Calibre
AZW3, then executes eight real PS5.1/PS7 EPUB/AZW3 default/retention runs. It
checks original ebook identities, same-name neighboring PDF hashes, read-only
sources, actual captured/retained PDF contents and all physical pages (including
AZW3's fourth inline-contents page). Two direct API controls independently inspect
the captured reader and prohibit replanning/reopening. Eight actual native
zero-exit controls reject missing, empty, corrupt and zero-page output under both
hosts; the unchanged BAT launcher preserves a missing-converter failure.
The runtime layer covers successful BAT converter discovery. Exact run/manifest members and
held cleanup are required. See `conversion/README.md` for observation limits.

The runtime layer checks AC-043/044/045 under both actual hosts: dependency
rejections before console/conversion/output writes, explicit/application-venv/
launcher/PATH interpreter selection, isolated imports, excluded book/CWD decoys,
PDF-only operation without Calibre, and trusted nonstandard converter discovery.
Actual native statuses, selected paths/versions/import origins, source/neighbor
preservation, complete output contents, stored policies and known cleanup are
required. See `runtime/README.md` for the exact executed cases and scoped limits.

The process layer checks AC-046..049 under both actual hosts and the unchanged
BAT. It requires bounded flood/tail/UTF-8 capture, literal native arguments,
independent final results, owned-tree timeout and cancellation, and complete
real PDF outputs for hostile Unicode paths/titles. It does not require Calibre;
the full route separately executes the inherited real EPUB/AZW3 regression.
See `process/README.md` for the exact scope. Human Ctrl+C and Explorer are not
inferred from automated cancellation controls.

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
GUID sentinel parents and explicit emitted run/log children. Current nonrecursive
cleanup verifies each exact child's containment, marker, allowed files and absence
of reparse points before removing only known owned members. The original process
and GUID-parent helper guards remain unchanged. No private Documents content is
enumerated or used as input. See `manual/README.md` for the current six-probe method;
`extraction/README.md` retains the historical method and limits. `manual/README.md`
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
Full UX/Explorer, expanded Calibre compatibility and conversion from the
extracted release package remain separate mandatory later checks. The local
conversion route covers the two original fixtures on the pinned converter.
CI installation and workflow
execution belong to M4-T05; this local scaffold does not claim a CI run.

M3-T04 adds `--layer ux` as stage 17 of the full route. Actual PS5.1/PS7 processes
hold the displayed plan before consent, verify no chapter reservation, then
reopen outputs for exact title/filename/range/page-content parity. Cancel, blank,
EOF, invalid answers, fallback and pending-input timeout are exercised. Real EPUB
and AZW3 manual controls use generated physical PDF numbering; direct engine
controls request an owned working copy and verify its availability and cleanup.
Actual PDF viewer and Explorer observations are separate recorded Windows checks,
not inferred from launch calls. See `ux/README.md`. Completed M3-T01 human testing
remains historical and closed.

M3-T05 adds `--layer support` as stage 18 of the full route. Both explicit
supported PowerShell hosts inspect actual UTF-8 local run manifests, finalized
bounded log bytes, captured versions/settings/source identities and final
outcomes. Authored parser and bookmark warnings must remain categorized and
visible. The explicit local exporter is checked with real run records and mock
private paths, titles, content and secrets, plus literal no-overwrite/alias and
invalid-record controls. Disclosed copied finalizer failures verify retained
chapter output and preservation of primary failure/cancellation codes. See
`support/README.md` for the native matrix and strict receipt checks; this route
does not reopen human testing or claim GUI, upload, CI or release evidence.
