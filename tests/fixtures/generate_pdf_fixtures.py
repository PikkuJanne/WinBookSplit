"""Generate original, reproducible page-marked PDFs for baseline characterization.

This development helper reads no source documents. The generated files and this
code use the repository's MIT license; see tests/fixtures/README.md.
"""

from __future__ import annotations

import argparse
from hashlib import sha256
from importlib.metadata import version
from io import BytesIO
import json
from pathlib import Path
import platform
import re

from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


PAGE_MARKER_TEMPLATE = "WBS-PAGE-{page_id:03d}"
PAGE_MARKER_RE = re.compile(r"(?<![A-Z0-9-])WBS-PAGE-([0-9]{3})(?![0-9])")
FIXED_PDF_DATE = "D:20000101000000Z"
FIXTURE_DEFINITIONS = {
    "simple10": {
        "filename": "simple-10.pdf",
        "pages": 10,
        "title": "WinBookSplit synthetic simple outline fixture",
        "outline": [
            {"title": "A", "start_page": 4, "children": []},
            {"title": "B", "start_page": 7, "children": []},
        ],
    },
    "nested12": {
        "filename": "nested-12.pdf",
        "pages": 12,
        "title": "WinBookSplit synthetic nested outline fixture",
        "outline": [
            {
                "title": "A",
                "start_page": 3,
                "children": [
                    {"title": "A1", "start_page": 4, "children": []},
                    {"title": "A2", "start_page": 7, "children": []},
                ],
            },
            {
                "title": "B",
                "start_page": 9,
                "children": [
                    {"title": "B1", "start_page": 11, "children": []},
                ],
            },
        ],
    },
}


def page_marker(page_id: int) -> str:
    """Return the visible original-page marker for a physical 1-based ID."""
    if isinstance(page_id, bool) or not isinstance(page_id, int) or not 1 <= page_id <= 999:
        raise ValueError("page_id must be an integer between 1 and 999")
    return PAGE_MARKER_TEMPLATE.format(page_id=page_id)


def page_ids(path: Path) -> list[int]:
    """Read original IDs, requiring exactly one visible marker on each page.

    IDs retain their original value in split PDFs. The returned order and any
    repeated IDs let callers detect omissions, reordered pages and duplication.
    """
    result = []
    for number, page in enumerate(PdfReader(path).pages, start=1):
        matches = PAGE_MARKER_RE.findall(page.extract_text() or "")
        if len(matches) != 1:
            raise ValueError(f"Page {number} has {len(matches)} page markers; expected one")
        page_id = int(matches[0])
        if page_id == 0:
            raise ValueError(f"Page {number} has an invalid zero page ID")
        result.append(page_id)
    return result


def _outline_nodes(reader: PdfReader, entries: list) -> list[dict]:
    """Read back the actual nested destinations using physical 1-based pages."""
    result = []
    for entry in entries:
        if isinstance(entry, list):
            if not result:
                raise ValueError("Outline children have no preceding parent")
            result[-1]["children"] = _outline_nodes(reader, entry)
        else:
            destination = reader.get_destination_page_number(entry)
            if destination is None or destination < 0:
                raise ValueError("Fixture outline has an unresolved destination")
            result.append(
                {"title": entry.title, "start_page": destination + 1, "children": []}
            )
    return result


def _pdf_bytes(definition: dict) -> bytes:
    raw_pages = BytesIO()
    document = canvas.Canvas(
        raw_pages, pagesize=A4, invariant=1, pageCompression=1
    )
    document.setTitle(definition["title"])
    document.setAuthor("WinBookSplit synthetic fixtures")
    document.setSubject("Original synthetic pages; no source document was used")
    width, height = A4
    for page_id in range(1, definition["pages"] + 1):
        document.setFont("Helvetica-Bold", 20)
        document.drawString(54, height - 70, "WinBookSplit test fixture")
        document.setFont("Helvetica", 12)
        document.drawString(54, height - 98, "Original generated content - MIT license")
        document.setFont("Courier-Bold", 32)
        document.drawCentredString(width / 2, height / 2, page_marker(page_id))
        document.setFont("Helvetica", 12)
        document.drawCentredString(
            width / 2, height / 2 - 32,
            f"Physical page {page_id} of {definition['pages']}",
        )
        document.setFont("Helvetica", 10)
        document.drawString(54, 54, "Synthetic fixture for page coverage and outline checks")
        document.showPage()
    document.save()

    writer = PdfWriter()
    for page in PdfReader(BytesIO(raw_pages.getvalue())).pages:
        writer.add_page(page)
    writer.metadata = None
    writer.add_metadata(
        {
            "/Title": definition["title"],
            "/Author": "WinBookSplit synthetic fixtures",
            "/Subject": "Original synthetic content under the repository MIT license",
            "/Creator": "WinBookSplit fixture generator",
            "/Producer": "WinBookSplit fixture generator (ReportLab and pypdf)",
            "/CreationDate": FIXED_PDF_DATE,
            "/ModDate": FIXED_PDF_DATE,
        }
    )

    def add_nodes(nodes: list[dict], parent=None) -> None:
        for node in nodes:
            item = writer.add_outline_item(
                node["title"], node["start_page"] - 1, parent=parent
            )
            add_nodes(node["children"], item)

    add_nodes(definition["outline"])
    writer.page_mode = "/UseOutlines"
    result = BytesIO()
    writer.write(result)
    return result.getvalue()


def _source_hashes() -> dict[str, str]:
    root = Path(__file__).resolve().parents[2]
    paths = [
        "tests/fixtures/generate_pdf_fixtures.py",
        "tests/fixtures/README.md",
        "LICENSE",
    ]
    return {name: sha256((root / name).read_bytes()).hexdigest() for name in paths}


def generate_fixtures(output_dir: Path) -> dict[str, Path]:
    """Write both fixtures and provenance.json to a new or empty directory.

    Existing files are never overwritten. Callers own the supplied per-run
    directory and its cleanup; the generator does not remove any files.
    """
    output_dir = Path(output_dir)
    if output_dir.exists():
        if not output_dir.is_dir() or any(output_dir.iterdir()):
            raise FileExistsError("Fixture output must be a new or empty directory")
    else:
        output_dir.mkdir(parents=True, exist_ok=False)

    paths = {}
    records = {}
    for key, definition in FIXTURE_DEFINITIONS.items():
        path = output_dir / definition["filename"]
        payload = _pdf_bytes(definition)
        with path.open("xb") as stream:
            stream.write(payload)
        ids = page_ids(path)
        if ids != list(range(1, definition["pages"] + 1)):
            raise ValueError(f"Fixture {key} page IDs did not round-trip")
        reader = PdfReader(path)
        outline = _outline_nodes(reader, reader.outline)
        if outline != definition["outline"]:
            raise ValueError(f"Fixture {key} outline did not round-trip")
        paths[key] = path
        records[key] = {
            "filename": path.name,
            "size_bytes": len(payload),
            "sha256": sha256(payload).hexdigest(),
            "page_ids": ids,
            "outline": outline,
            "metadata": dict(reader.metadata),
        }

    provenance = {
        "schema_version": 1,
        "content_origin": "Original programmatically generated text; no source documents",
        "license": "MIT - same license as repository LICENSE",
        "page_marker_template": PAGE_MARKER_TEMPLATE,
        "fixed_pdf_date": FIXED_PDF_DATE,
        "reproducibility_scope": "Same source bytes and Python/ReportLab/pypdf versions",
        "runtime_versions": {
            "python": platform.python_version(),
            "pypdf": version("pypdf"),
            "reportlab": version("reportlab"),
        },
        "source_sha256": _source_hashes(),
        "fixtures": records,
    }
    with (output_dir / "provenance.json").open("x", encoding="utf-8", newline="\n") as stream:
        json.dump(provenance, stream, ensure_ascii=True, indent=2, sort_keys=True)
        stream.write("\n")
    return paths


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output-dir", required=True, type=Path,
        help="Explicit new or empty per-run directory outside the repository",
    )
    args = parser.parse_args()
    try:
        paths = generate_fixtures(args.output_dir)
    except (FileExistsError, OSError, ValueError) as error:
        parser.exit(1, f"Fixture generation failed: {error}\n")
    for key, path in paths.items():
        print(f"{key}: {path}")
    print(f"provenance: {args.output_dir / 'provenance.json'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
