"""Author tiny offline EPUB and generate genuine AZW3 with external Calibre.

All book text is original test material under the repository MIT license. No
existing book or remote resource supplies fixture content. Converter invocation
does not edit installed Calibre settings or its library; Calibre may read its
normal configuration.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
from hashlib import sha256
import json
from pathlib import Path
import subprocess
import zipfile


EXPECTED_CALIBRE_VERSION = "9.15.0"
EXPECTED_CALIBRE_SHA256 = "f46a01c9b8cd392e1190ce2e50a2a5a48b326c8c3eab670bbcd8b298bede501d"
CHAPTER_MARKERS = ["WBS-PAGE-001", "WBS-PAGE-002", "WBS-PAGE-003"]
CHAPTER_TITLES = ["Chapter One", "Chapter Two", "Chapter Three"]


def digest(path):
    return sha256(path.read_bytes()).hexdigest()


def authored_epub(path):
    """Write reproducible EPUB 3 bytes, including EPUB 2 NCX compatibility."""
    chapters = []
    for number, (title, marker) in enumerate(zip(CHAPTER_TITLES, CHAPTER_MARKERS), 1):
        chapters.append((f"OEBPS/chapter{number}.xhtml", f'''<?xml version="1.0" encoding="utf-8"?>
<html xmlns="http://www.w3.org/1999/xhtml"><head><title>{title}</title>
<link rel="stylesheet" href="style.css" type="text/css"/></head><body>
<h1>{title}</h1><p class="marker">{marker}</p>
<p>This chapter is original WinBookSplit test material. Its visible marker identifies the authored chapter.</p>
<p>The three small chapters contain no images, scripts, links to remote resources, or private source material.</p>
</body></html>'''))
    items = "".join(f'<item id="c{n}" href="chapter{n}.xhtml" media-type="application/xhtml+xml"/>' for n in range(1, 4))
    spine = "".join(f'<itemref idref="c{n}"/>' for n in range(1, 4))
    nav_links = "".join(f'<li><a href="chapter{n}.xhtml">{title}</a></li>' for n, title in enumerate(CHAPTER_TITLES, 1))
    ncx_links = "".join(f'<navPoint id="n{n}" playOrder="{n}"><navLabel><text>{title}</text></navLabel><content src="chapter{n}.xhtml"/></navPoint>'
                        for n, title in enumerate(CHAPTER_TITLES, 1))
    files = [
        ("mimetype", "application/epub+zip"),
        ("META-INF/container.xml", '''<?xml version="1.0"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles>
<rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles></container>'''),
        ("OEBPS/content.opf", f'''<?xml version="1.0" encoding="utf-8"?>
<package version="3.0" unique-identifier="bookid" xmlns="http://www.idpf.org/2007/opf">
<metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:identifier id="bookid">urn:uuid:34d2be18-390e-4c58-bd8f-59f7272c7abd</dc:identifier>
<dc:title>Original WinBookSplit Three Chapters</dc:title><dc:creator>WinBookSplit test authors</dc:creator>
<dc:language>en</dc:language><dc:rights>Original synthetic material; MIT license</dc:rights><meta property="dcterms:modified">2000-01-01T00:00:00Z</meta></metadata>
<manifest>{items}<item id="style" href="style.css" media-type="text/css"/>
<item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>
<item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/></manifest>
<spine toc="ncx">{spine}</spine></package>'''),
        ("OEBPS/nav.xhtml", f'''<?xml version="1.0" encoding="utf-8"?>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops"><head><title>Contents</title></head>
<body><nav epub:type="toc"><h1>Contents</h1><ol>{nav_links}</ol></nav></body></html>'''),
        ("OEBPS/toc.ncx", f'''<?xml version="1.0" encoding="utf-8"?>
<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1"><head><meta name="dtb:uid" content="urn:uuid:34d2be18-390e-4c58-bd8f-59f7272c7abd"/></head>
<docTitle><text>Original WinBookSplit Three Chapters</text></docTitle><navMap>{ncx_links}</navMap></ncx>'''),
        ("OEBPS/style.css", "body { font-family: serif; } h1 { page-break-before: always; } .marker { font-size: 24pt; font-weight: bold; }"),
        *chapters,
    ]
    with path.open("xb") as stream, zipfile.ZipFile(stream, "w") as archive:
        for name, text in files:
            info = zipfile.ZipInfo(name, date_time=(2000, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_STORED if name == "mimetype" else zipfile.ZIP_DEFLATED
            info.external_attr = 0o100644 << 16
            archive.writestr(info, text.encode("utf-8"))
    return path


def run(argv, cwd):
    started = datetime.now(timezone.utc).isoformat()
    completed = subprocess.run([str(value) for value in argv], cwd=cwd, capture_output=True,
                               text=True, encoding="utf-8", errors="replace", timeout=180)
    return {"argv": [str(value) for value in argv], "cwd": str(cwd), "started_at": started,
            "finished_at": datetime.now(timezone.utc).isoformat(), "exit_code": completed.returncode,
            "stdout": completed.stdout, "stderr": completed.stderr}


def generate(directory, calibre):
    directory = Path(directory).resolve()
    calibre = Path(calibre).resolve(strict=True)
    if not calibre.is_file() or digest(calibre) != EXPECTED_CALIBRE_SHA256:
        raise RuntimeError("Fixture generation requires the recorded trusted Calibre 9.15.0 executable bytes")
    directory.mkdir(parents=True, exist_ok=True)
    if any(directory.iterdir()):
        raise RuntimeError("Fixture destination must be a new or empty owned directory")
    version = run([calibre, "--version"], directory)
    if version["exit_code"] or f"calibre {EXPECTED_CALIBRE_VERSION}" not in version["stdout"]:
        raise RuntimeError("Actual converter version does not match the selected Calibre version")
    epub = authored_epub(directory / "original-three-chapters.epub")
    azw3 = directory / "original-three-chapters.azw3"
    conversion = run([calibre, epub, azw3, "--output-profile", "tablet"], directory)
    if conversion["exit_code"] or not azw3.is_file() or not azw3.stat().st_size:
        raise RuntimeError("Actual EPUB to AZW3 generation failed: " + json.dumps(conversion, ensure_ascii=True))
    report = {"schema_version": 1, "kind": "original-offline-ebook-fixtures", "license": "MIT",
              "authored_original": True, "remote_resources": False, "chapter_markers": CHAPTER_MARKERS,
              "chapter_titles": CHAPTER_TITLES,
              "calibre": {"path": str(calibre), "sha256": digest(calibre), "version": EXPECTED_CALIBRE_VERSION,
                           "version_observation": version},
              "azw3_generation": conversion,
              "files": [{"name": path.name, "sha256": digest(path), "size_bytes": path.stat().st_size,
                         "format": path.suffix[1:], "origin": "authored-epub" if path == epub else "actual-calibre-conversion"}
                        for path in (epub, azw3)]}
    (directory / "fixture-provenance.json").write_text(json.dumps(report, indent=2, ensure_ascii=True) + "\n", encoding="utf-8")
    return report


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--calibre-path", type=Path, required=True)
    arguments = parser.parse_args()
    print(json.dumps(generate(arguments.output_dir, arguments.calibre_path), ensure_ascii=True))
