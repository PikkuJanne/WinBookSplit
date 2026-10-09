# Current Level 1 regression route

`characterize_level1.py` executes corrected BM-01/02/05/08 against the shipped
engine's actual CLI using original generated page-marked PDFs. Each output is
reopened and its exact page identities/order and complete flattened coverage are
checked. Front matter, first-source duplicate titles, reordering warnings and
the no-bookmark rejection are asserted. Input and synthetic output-neighbor
bytes remain unchanged.

Five normalization aggregates cover invalid/external/raw/named destinations,
depth/lineage/source order, cyclic outlines, malformed trees and traversal
bounds. Callable mock readers test otherwise difficult invalid page types,
resolution exceptions, reused nodes/containers and raw outline/name-tree
preflight before recursive outline retrieval. An actual serialized cyclic PDF
must fail the CLI before writing. A generated PDF containing valid named/raw,
external, unresolved named and invalid-fit destinations must retain only usable
boundaries and preserve every physical page with diagnostics. Rejected callable
cases assert that the slice writer is never reached. These records distinguish
actual PDF/CLI probes from synthetic-reader tests.

Seed 20261009 exercises 150 Level 1 plans, including duplicate/out-of-order
starts and first-source title selection. Ten samples also use the real writer
with generated marked PDFs. Depth 64 and 10000 nodes are accepted; exceeding
those bounds fails with a diagnostic. Raw tree/name-tree cycles and limits must
be caught before recursive pypdf outline retrieval.

Use the supported fresh developer interpreter with `-I -B`, an absolute runner
path and a new external report: `tests/run_tests.py --layer bookmarks --report ...`.
The full route includes this stage after manual regressions. This route receives
no shell arguments. Existing full/manual owned launcher probes still exercise
corrected MAN-03 and unchanged Level 2 BM-03; actual Level 1 launcher paths remain
a separate check. The current manual route compares only BM-03 after Level 1
corrections. Historical baseline/extraction guards, expected observations,
fixture generator/oracles and prior evidence remain unchanged.

The report records actual tested source/input hashes, immutable Git guards,
safe import, every promised case and owned temporary cleanup. It does not
certify corrected Level 2, all launcher/error paths, Explorer, Calibre,
rendered fidelity, output transaction safety or an extracted release package.
No private document is read, committed or uploaded.
