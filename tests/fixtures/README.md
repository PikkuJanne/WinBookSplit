# Original PDF fixtures

`generate_pdf_fixtures.py` creates synthetic text pages from source code. It reads
no input documents and contains no textbooks, private material, workplace names,
or downloaded content. The generator and its original generated content use the
repository's MIT license in `LICENSE`. Generated PDFs and per-run artifacts must
stay outside Git.

This is a development helper requiring Python, ReportLab and pypdf. M0-T02 uses
the locally available isolated Codex runtime; its exact versions are recorded in
the run's provenance. That observation does not define the application's
supported dependency versions; M0-T03 selects those separately.

Run it from the repository root with an explicit new or empty output directory:

```powershell
& $FixturePython .\tests\fixtures\generate_pdf_fixtures.py --output-dir $FixtureRunDirectory
```

Set `$FixturePython` to the chosen isolated Python executable and
`$FixtureRunDirectory` to a new per-run directory outside the repository. The
helper refuses nonempty directories and uses exclusive file creation. It never
overwrites or removes files. The caller owns only its chosen per-run artifacts
and their cleanup. A second generation uses a separate empty directory.

| API key / PDF | Pages | Outline destinations (physical, 1-based) |
| --- | --- | --- |
| `simple10` / `simple-10.pdf` | 10 | A: 4, B: 7 |
| `nested12` / `nested-12.pdf` | 12 | A: 3 (A1: 4, A2: 7), B: 9 (B1: 11) |

Every physical page contains exactly one visible `WBS-PAGE-001` style marker.
`page_marker(page_id)` returns this string using `PAGE_MARKER_TEMPLATE`; IDs run
from 1 to the fixture's page count. `page_ids(path: Path) -> list[int]` extracts
the original page IDs in file order and rejects pages with missing, multiple or
zero markers. Split pages retain original IDs, so tests can detect missing,
reordered and duplicated pages instead of relying only on the output page count.

`generate_fixtures(output_dir: Path) -> dict[str, Path]` returns paths with the
keys `simple10` and `nested12`. Before returning, it reopens each PDF and checks
the exact page ID sequence and nested outline destinations. It also writes
`provenance.json` containing the generated file sizes and SHA-256 hashes,
round-tripped page IDs/outlines/metadata, Python/ReportLab/pypdf versions and
source SHA-256 hashes for the generator, this README and `LICENSE`. It records no
absolute workstation paths, usernames, hostname or current timestamps.

ReportLab uses `invariant=1`; final pypdf metadata has a fixed date and fixed
title/author/creator/producer. Identical source bytes and dependency versions
produce byte-identical PDFs and provenance on repeated runs. Byte equality
across different dependency versions is not promised. Render the generated PDFs
with Poppler and inspect the visible markers as part of fixture validation.

The required corrected split ranges live in
`docs/codex-v1.0.0/PLAN_ORACLES.json`. Those are target requirements, not observed
baseline results. The baseline characterization harness records the unchanged
engine's actual behavior separately. Generating a fixture is not a passing
Windows launcher, Explorer or Calibre acceptance test.
