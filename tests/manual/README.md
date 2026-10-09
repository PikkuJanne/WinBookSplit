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
notices are asserted. Current safe import remains checked. Current historical
bookmark comparisons are empty after M1-T04 repairs Level 2. Separate Level 1/2
routes verify corrected targets. Historical M1-T02 evidence retains the three
then-unchanged observations; M1-T03 evidence retains its then-unchanged BM-03.

The original harness, its raw-byte guards, fixture generator, plan oracles,
batch launcher and extraction harness remain unchanged. Historical AC-011 AST
verification reads the extracted engine at immutable M1-T01 commit
`88c2149b3b3034fbd0d7ef23c2382f4b01648e4c`. The historical baseline route still
refuses changed PowerShell bytes; the extraction route still refuses changed
processing ASTs. Neither route calls corrected code the unchanged original.

Use the supported fresh developer interpreter with `-I -B`, an absolute runner
path and a new external report: `tests/run_tests.py --layer manual --report ...`.
`--layer full` runs current manual acceptance, corrected Level 1/2 bookmark
acceptance, Python units and the shell layer.
Optional repeated actual PS5.1/PS7 `--shell-path` arguments run six current
GUID-owned entrypoint probes, including three simultaneous invocations.
Manual `4,7` uses corrected exact output references. The actual current CLI
also builds a corrected BM-03 launcher reference with all six ranges, complete
physical page identities and Front matter/A opening/A1/A2/B opening/B1 titles.
The six probes compare its corrected bytes. This is recorded as
`level2_launcher_reference`, separate from historical observations. The test
process safety and GUID-parent guards reuse the unchanged historical helpers.
`current_launchers.py` reads only exact emitted Output/Log paths, validates
the direct child manifest/marker/files, and removes those known regular members
nonrecursively through held native delete handles after capturing and matching
their identities under root-to-parent directory guards. The unchanged historical
parent cleanup then removes its separate GUID sentinel directory. Unknown/reparse paths or
uncertain owned process termination are preserved. Reports must include
all promised cases and requested entrypoints or the central runner fails.

These controlled stdin/PATH/current-workstation probes do not certify Explorer,
all launcher/error paths, Calibre conversion, complete output transaction safety, fidelity
or an extracted release package. No private document is processed or uploaded.

Since M2-T01, every current writer check uses the explicitly returned published
child and its complete manifest, owner marker, ordered output digests/sizes and
reopened page IDs. The output argument is a base; preserved root neighbors and
earlier runs cannot be mixed into the new slices. The separate output layer
checks failure handling, concurrency and actual Windows junction boundaries.
