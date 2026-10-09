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
corrected MAN-03 and corrected Level 2 BM-03; actual Level 1 launcher paths remain
a separate check. Current manual historical comparisons are empty after Level 2
corrections. Historical baseline/extraction guards, expected observations,
fixture generator/oracles and prior evidence remain unchanged.

The report records actual tested source/input hashes, immutable Git guards,
safe import, every promised case and owned temporary cleanup. It does not
certify corrected Level 2, all launcher/error paths, Explorer, Calibre,
rendered fidelity, output transaction safety or an extracted release package.
No private document is read, committed or uploaded.

# Current Level 2 regression route

`characterize_level2.py` executes corrected BM-03/04/06/07 against the actual
CLI with original marked PDFs. Exact slice identities/order and the complete
flattened physical sequence are checked. The parent-aware example yields six
sections, including front matter and both parents' opening pages. Whole-parent
fallback, a child equal to its parent start, outside-child diagnostics and
structured no-selected-level rejection are verified without silent fallback.

Six hierarchy aggregates exercise duplicate-parent and invalid-parent subtrees,
child ordering/aliases, invalid child destinations, deeper/malformed outlines
and no usable selected level. Callable mocks distinguish unusable ancestors
from valid direct children and assert that rejected calls never reach the
writer. Actual generated PDFs test valid named child destinations, invalid
raw integer/object-ID collisions, unsupported fits, unresolved names, external
children and a serialized cyclic outline. These records distinguish real CLI
checks from mock-reader checks.

Seed 20261009 runs 150 independently expected parent/child plans. Ten further
samples use actual generated PDFs and the real CLI writer. Every emitted entry
must stay within a retained parent's interval and every selected bookmark must
be that parent's own direct child. Exact titles, first-source alias selection,
range adjacency and physical page identity/order are checked.

Use `tests/run_tests.py --layer level2 --report ...` with the supported explicit
developer interpreter, `-I -B`, an absolute runner path and a new external report.
This route receives no shell arguments. The full route adds it as its sixth
stage. The manual route supplies the corrected actual six-section BM-03
reference to the existing six owned launchers, including three concurrent runs.
Its historical comparison list is empty; earlier evidence/guards stay intact.

The report requires all four CLI targets, all six hierarchy records, 150 seeded
plans, ten nonempty writer samples, immutable original guards, safe import,
source/input/neighbor preservation and owned temporary cleanup. Interactive
fallback offers, all launcher/error paths, Explorer, Calibre, rendered fidelity,
output transactions and extracted-package acceptance remain separate checks.
