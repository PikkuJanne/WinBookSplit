# Current manual regression route

M1-T02 deliberately replaces the current manual known-bad comparisons with the
unchanged `PLAN_ORACLES.json` targets. `characterize_manual.py` executes all 22
manual oracles against the shipped engine with generated PDFs, reopens every
slice and checks each page marker and the flattened physical page sequence.
Invalid lists fail before any output; synthetic source/neighbor hashes remain
unchanged. Eight extra actual CLI cases cover comma-only/zero/missing input,
non-ASCII digits, embedded spaces, a late invalid token, a 5000-digit out-of-range
integer and a valid 5001-character leading-zero token.

Seed 20261009 exercises 250 valid manual partitions; 25 samples also exercise
the real PDF writer on page-marked fixtures. Normalization and one-section
notices are asserted. Current safe import remains checked. Three original
bookmark observations (BM-01/02/03) still match the immutable original engine,
including their known defects; repaired bookmark acceptance starts M1-T03.

The original harness, its raw-byte guards, fixture generator, plan oracles,
batch launcher and extraction harness remain unchanged. Historical AC-011 AST
verification reads the extracted engine at immutable M1-T01 commit
`88c2149b3b3034fbd0d7ef23c2382f4b01648e4c`. The historical baseline route still
refuses changed PowerShell bytes; the extraction route still refuses changed
processing ASTs. Neither route calls corrected code the unchanged original.

Use the supported fresh developer interpreter with `-I -B`, an absolute runner
path and a new external report: `tests/run_tests.py --layer manual --report ...`.
`--layer full` runs current manual acceptance, Python units and the shell layer.
Optional repeated actual PS5.1/PS7 `--shell-path` arguments reuse the existing
six GUID-owned entrypoint probes, including three simultaneous invocations.
Manual `4,7` uses the corrected exact output references; bookmark BM-03 remains
a known-bad comparison. The test output/cleanup/process safety boundaries are
documented in `../extraction/README.md` and are unchanged. Reports must include
all promised cases and requested entrypoints or the central runner fails.

These controlled stdin/PATH/current-workstation probes do not certify Explorer,
all launcher/error paths, Calibre conversion, output transaction safety, fidelity
or an extracted release package. No private document is processed or uploaded.
