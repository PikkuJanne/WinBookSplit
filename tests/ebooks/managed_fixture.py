"""Inspect chapters from a real conversion of the authored illustrated EPUB."""

from hashlib import sha256
from pathlib import Path
import runpy
import unicodedata

from pypdf import PdfReader


ROOT = Path(__file__).resolve().parents[2]
RICH_CHAPTERS = runpy.run_path(str(ROOT / "tests/conversion/generate_ebook_fixtures.py"))["RICH_CHAPTERS"]


def require(condition, message):
    if not condition:
        raise ValueError(message)


def normalized(value):
    return " ".join(unicodedata.normalize("NFC", value).split())


def inspect_page(page):
    content = page.get_contents()
    operations = content.operations if content is not None else ()
    drawn = {str(operands[0]) for operands, operator in operations if operator == b"Do" and operands}
    images = []
    for image_key in page.images.keys():
        first_name = image_key[0] if isinstance(image_key, tuple) else image_key
        if str(first_name) not in drawn:
            continue
        image = page.images[image_key].image
        require(image is not None, "A drawn PDF image could not be decoded")
        rgb = image.convert("RGB")
        width, height = rgb.size
        images.append((width, height, sha256(rgb.tobytes()).hexdigest(), rgb.getpixel((width // 2, height // 2))))
    return {
        "content_sha256": sha256(content.get_data() if content is not None else b"").hexdigest(),
        "text": normalized(page.extract_text() or ""),
        "geometry": (list(page.mediabox), list(page.cropbox), page.rotation),
        "images": sorted(images),
    }


def verify_chapters(generated_pdf, chapters, *, auto1=False):
    """Compare actual output files against the full generated PDF and authored content."""
    generated_pdf = Path(generated_pdf)
    source = PdfReader(generated_pdf)
    require(3 <= len(source.pages) <= 32, "Unexpected physical page count for authored illustrated ebook")
    source_pages = [inspect_page(page) for page in source.pages]
    marker_pages = {}
    for title, marker, text, color in RICH_CHAPTERS:
        matches = [index for index, page in enumerate(source_pages) if marker in page["text"]]
        require(len(matches) == 1, f"Generated PDF lost or duplicated authored chapter {marker}")
        page = source_pages[matches[0]]
        require(text in page["text"], f"Generated PDF lost the Unicode text for {marker}")
        require(any(width >= 8 and height >= 8 and max(abs(value - expected) for value, expected in zip(pixel, color)) <= 20
                    for width, height, _, pixel in page["images"]), f"Generated PDF lost the drawn RGB illustration for {marker}")
        marker_pages[title] = matches[0]
    require(list(marker_pages.values()) == sorted(set(marker_pages.values())), "Converted author chapters lost their order")

    chapter_receipts, cursor = [], 0
    for start, end, path in chapters:
        require(type(start) is int and type(end) is int and start == cursor and start < end <= len(source_pages),
                "Chapter ranges omit, duplicate, or reorder generated physical pages")
        chapter = PdfReader(path)
        require(not chapter.is_encrypted and len(chapter.pages) == end - start, "Chapter could not be reopened with the planned page count")
        require([inspect_page(page) for page in chapter.pages] == source_pages[start:end],
                "Chapter changed generated page text, content, geometry, or drawn images")
        chapter_receipts.append({"range": [start, end], "filename": Path(path).name,
                                 "sha256": sha256(Path(path).read_bytes()).hexdigest()})
        cursor = end
    require(cursor == len(source_pages), "Chapter output does not cover the final generated page")

    outline = []
    for node in source.outline:
        if not isinstance(node, list):
            page = source.get_destination_page_number(node)
            if page is not None and 0 <= page < len(source_pages):
                outline.append((normalized(node.title), page))
    if auto1:
        authored = [(title, page) for title, page in outline if title in marker_pages]
        require(authored == list(marker_pages.items()), "Generated outline lost or misdirected the authored EPUB chapter navigation")
        starts = sorted({0, *(page for _, page in outline)})
        expected_ranges = [[start, end] for start, end in zip(starts, starts[1:] + [len(source_pages)])]
        require([entry["range"] for entry in chapter_receipts] == expected_ranges, "Auto1 split did not follow generated outline pages")
    return {
        "generated_sha256": sha256(generated_pdf.read_bytes()).hexdigest(),
        "page_count": len(source_pages),
        "authored_marker_pages_zero_based": marker_pages,
        "outline": outline,
        "chapters": chapter_receipts,
    }
