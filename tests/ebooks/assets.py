"""Small authored-ebook format and physical PDF observations for M4-T04.

These checks describe the existing three-chapter fixture, not arbitrary ebooks.
Actual conversion/rendering is invoked only by the task's coordinated native runs.
"""
from __future__ import annotations

from hashlib import sha256
import importlib.util
import io
import json
import math
import os
from pathlib import Path, PurePosixPath
import posixpath
import re
import struct
import sys
from urllib.parse import urlsplit
import xml.etree.ElementTree as ET
import zipfile

from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[2]
MAX_EBOOK_BYTES = 8 * 1024 * 1024
MAX_MEMBER_BYTES = 64 * 1024
MEMBERS = {"mimetype", "META-INF/container.xml", "OEBPS/content.opf", "OEBPS/nav.xhtml",
           "OEBPS/toc.ncx", "OEBPS/style.css", *(f"OEBPS/chapter{n}.xhtml" for n in range(1, 4))}
AUTHORED_CSS = "body { font-family: serif; } h1 { page-break-before: always; } .marker { font-size: 24pt; font-weight: bold; }"


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


fixtures = load("wbs_ebook_original_fixtures", ROOT / "tests/conversion/generate_ebook_fixtures.py")
fidelity = load("wbs_ebook_render_helpers", ROOT / "tests/fidelity/characterize_fidelity.py")
behavior = load("wbs_ebook_asset_native", ROOT / "tests/ebooks/behavior.py")
NATIVE_OPERATIONS = behavior.NATIVE_OPERATIONS


def native(command, cwd, timeout=180):
    observation = behavior.native(command, cwd, dict(os.environ), timeout=timeout)
    observation["argv"] = observation["command"]
    return observation


def need(condition, message):
    if not condition:
        raise ValueError(message)


def digest(path):
    return sha256(Path(path).read_bytes()).hexdigest()


def bounded_bytes(path):
    with Path(path).open("rb") as stream:
        data = stream.read(MAX_EBOOK_BYTES + 1)
    need(0 < len(data) <= MAX_EBOOK_BYTES, "Authored ebook size outside bounded fixture scope")
    return data


def inspect_epub_data(data, *, allowed_urls=()):
    """Verify EPUB3/NCX and exact local resource closure of the authored family."""
    need(isinstance(data, bytes) and 0 < len(data) <= MAX_EBOOK_BYTES, "Invalid bounded EPUB bytes")
    need(len(set(allowed_urls)) == len(allowed_urls) and len(allowed_urls) in (0, 2), "Unexpected canary URL set")
    for url in allowed_urls:
        split = urlsplit(url)
        need(split.scheme == "http" and split.hostname == "127.0.0.1" and split.port is not None
             and 0 < split.port < 65536 and split.username is None and not split.query and not split.fragment,
             "Canary fixture can reference only exact authored loopback URLs")
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        infos = archive.infolist()
        need(len(infos) == len(MEMBERS) and {item.filename for item in infos} == MEMBERS,
             "EPUB fixture has missing, duplicate or foreign members")
        need(infos[0].filename == "mimetype" and infos[0].compress_type == zipfile.ZIP_STORED,
             "EPUB mimetype must be first and uncompressed")
        need(all(0 < item.file_size <= MAX_MEMBER_BYTES and not item.flag_bits & 1
                 and item.filename == PurePosixPath(item.filename).as_posix() for item in infos)
             and sum(item.file_size for item in infos) <= 256 * 1024, "Unsafe or oversized EPUB member")
        members = {item.filename: archive.read(item.filename) for item in infos}
    need(members["mimetype"] == b"application/epub+zip", "EPUB declared format differs")
    parsed, refs, external = {}, [], []
    for name, raw in members.items():
        if name.endswith((".xml", ".opf", ".xhtml", ".ncx")):
            need(b"<!DOCTYPE" not in raw.upper() and b"<!ENTITY" not in raw.upper(), "EPUB fixture contains declarations outside scope")
            node = ET.fromstring(raw)
            parsed[name] = node
            for element in node.iter():
                tag = element.tag.rsplit("}", 1)[-1]
                need(tag.lower() not in {"script", "iframe", "object", "embed", "style"}, "EPUB fixture contains active or uninspected inline style content")
                for attribute, value in element.attrib.items():
                    key = attribute.rsplit("}", 1)[-1]
                    need(not key.lower().startswith("on"), "EPUB fixture contains an event handler")
                    need(key.lower() != "style", "Authored EPUB has no inline style attributes")
                    if key not in {"href", "src", "full-path"}:
                        continue
                    need(isinstance(value, str) and value and len(value) <= 1024, "Unbounded EPUB resource reference")
                    split = urlsplit(value)
                    if split.scheme or split.netloc:
                        need(value in allowed_urls, "EPUB contains an undeclared remote resource")
                        external.append({"from": name, "reference": value})
                        continue
                    need(not split.query and not split.path.startswith("/") and "\\" not in split.path,
                         "EPUB reference escapes its authored archive")
                    target = posixpath.normpath(posixpath.join(posixpath.dirname(name), split.path)) if key != "full-path" else split.path
                    if not split.path:
                        target = name
                    need(target in members, "EPUB has an unresolved local reference")
                    refs.append({"from": name, "reference": value, "target": target})
    css = members["OEBPS/style.css"].decode("utf-8")
    need(css == AUTHORED_CSS, "Local CSS differs from the exact original no-fetch authored stylesheet")
    package = parsed["OEBPS/content.opf"]
    need(package.tag == "{http://www.idpf.org/2007/opf}package" and package.get("version") == "3.0", "Not authored EPUB3 OPF")
    opf = {"o": "http://www.idpf.org/2007/opf"}
    spine = [element.get("idref") for element in package.findall("o:spine/o:itemref", opf)]
    need(spine == ["c1", "c2", "c3"], "Authored EPUB chapter spine changed")
    nav = parsed["OEBPS/nav.xhtml"]
    nav_links = [("".join(element.itertext()), element.get("href")) for element in nav.iter()
                 if element.tag == "{http://www.w3.org/1999/xhtml}a"]
    need(nav_links == list(zip(fixtures.CHAPTER_TITLES, [f"chapter{n}.xhtml" for n in range(1, 4)])), "Authored EPUB navigation changed")
    ncx = parsed["OEBPS/toc.ncx"]
    ncx_paths = [element.get("src") for element in ncx.iter() if element.tag.endswith("}content")]
    need(ncx_paths == [f"chapter{n}.xhtml" for n in range(1, 4)], "Authored NCX navigation changed")
    need(sorted(item["reference"] for item in external) == sorted(allowed_urls), "Canary resource declarations differ from actual archive")
    return {"format": "epub", "sha256": sha256(data).hexdigest(), "size_bytes": len(data),
            "epub_version": "3.0", "mimetype_first_uncompressed": True, "spine": spine,
            "navigation": [{"title": title, "href": href} for title, href in nav_links],
            "members": {name: {"size_bytes": len(raw), "sha256": sha256(raw).hexdigest()} for name, raw in sorted(members.items())},
            "local_references": sorted(refs, key=lambda item: (item["from"], item["reference"])),
            "external_references": sorted(external, key=lambda item: (item["from"], item["reference"])),
            "scripts_or_active_content": False, "resource_scope": "closed-authored-family" if not allowed_urls else "exact-authored-loopback-canary"}


def inspect_azw3_data(data):
    """Bounded PalmDB/MOBI8 structure check; never decrypt or execute a book."""
    need(isinstance(data, bytes) and 128 <= len(data) <= MAX_EBOOK_BYTES, "Invalid bounded AZW3 bytes")
    need(not zipfile.is_zipfile(io.BytesIO(data)) and data[60:68] == b"BOOKMOBI", "AZW3 must be genuine PalmDB BOOKMOBI, not a renamed EPUB")
    count = struct.unpack_from(">H", data, 76)[0]
    need(2 <= count <= 4096 and 78 + count * 8 <= len(data), "Invalid bounded PalmDB record table")
    offsets = [struct.unpack_from(">I", data, 78 + number * 8)[0] for number in range(count)]
    need(offsets[0] >= 78 + count * 8 and offsets[-1] < len(data)
         and all(first < second for first, second in zip(offsets, offsets[1:])), "Invalid PalmDB record offsets")
    zero = data[offsets[0]:offsets[1]]
    need(128 <= len(zero) <= MAX_MEMBER_BYTES and zero[16:20] == b"MOBI", "AZW3 lacks bounded MOBI record zero")
    header_length, version = struct.unpack_from(">I", zero, 20)[0], struct.unpack_from(">I", zero, 36)[0]
    encryption = struct.unpack_from(">H", zero, 12)[0]
    need(112 <= header_length <= 500 and 16 + header_length <= len(zero) and version == 8 and encryption == 0,
         "Authored AZW3 must be unencrypted standalone MOBI8/KF8")
    return {"format": "azw3", "sha256": sha256(data).hexdigest(), "size_bytes": len(data),
            "palmdb_identity": "BOOKMOBI", "record_count": count, "record_offsets": offsets,
            "record_zero_size_bytes": len(zero), "record_zero_sha256": sha256(zero).hexdigest(),
            "mobi_header_length": header_length, "mobi_version": version, "encryption_type": encryption,
            "not_zip": True, "scope": "bounded-generated-standalone-KF8-fixture"}


def generate(directory, calibre):
    NATIVE_OPERATIONS.clear()
    previous = fixtures.run
    fixtures.run = native
    try:
        provenance = fixtures.generate(directory, calibre)
    finally:
        fixtures.run = previous
    sources = {item["format"]: Path(directory) / item["name"] for item in provenance["files"]}
    checks = {"epub": inspect_epub_data(bounded_bytes(sources["epub"])), "azw3": inspect_azw3_data(bounded_bytes(sources["azw3"]))}
    for item in provenance["files"]:
        need(all(item[key] == checks[item["format"]][key] for key in ("sha256", "size_bytes")), "Actual format bytes differ from fixture provenance")
    return {**provenance, "format_checks": checks, "resource_closure": checks["epub"]}


def authored_canary_epub(path, base_url, source_epub):
    split = urlsplit(base_url)
    need(split.scheme == "http" and split.hostname == "127.0.0.1" and split.port is not None
         and 0 < split.port < 65536 and split.username is None and not split.query and not split.fragment
         and re.fullmatch(r"/[A-Za-z0-9_-]+/", split.path), "Canary base must be exact authored loopback namespace")
    original = bounded_bytes(source_epub)
    original_check = inspect_epub_data(original)
    urls = [base_url + "image.png", base_url + "style.css"]
    with zipfile.ZipFile(io.BytesIO(original)) as archive:
        entries = [(info, archive.read(info.filename)) for info in archive.infolist()]
    with Path(path).open("xb") as stream, zipfile.ZipFile(stream, "w") as archive:
        for info, raw in entries:
            if info.filename == "OEBPS/chapter1.xhtml":
                raw = raw.replace(b"</head>", f'<link rel="stylesheet" type="text/css" href="{urls[1]}"/></head>'.encode("ascii"), 1)
                raw = raw.replace(b"</body>", f'<img src="{urls[0]}" alt="Authored loopback canary"/></body>'.encode("ascii"), 1)
            archive.writestr(info, raw)
    check = inspect_epub_data(bounded_bytes(path), allowed_urls=urls)
    return {"path": str(Path(path)), "source_epub_sha256": original_check["sha256"], "urls": urls,
            "sha256": check["sha256"], "size_bytes": check["size_bytes"], "resource_closure": check,
            "authored_original_derivative": True, "scope": "only-two-declared-loopback-resource-URLs; no internet or script"}


def observe_pdf(path, renderer, render_root):
    path, render_root = Path(path), Path(render_root)
    need(path.is_absolute() and render_root.is_absolute() and not render_root.is_relative_to(ROOT), "PDF/render artifacts must be external absolute paths")
    render_root.mkdir(parents=True, exist_ok=False)
    before = {"sha256": digest(path), "size_bytes": path.stat().st_size}
    reader = PdfReader(path)
    need(not reader.is_encrypted and 1 <= len(reader.pages) <= 16, "Generated fixture PDF outside positive bounded page scope")
    prefix = render_root / "page"
    process = native([str(renderer), "-r", "72", "-cropbox", "-png", str(path), str(prefix)], render_root)
    need(process["exit_code"] == 0, "Actual generated/chapter Poppler rendering failed")
    files = sorted(render_root.glob("page-*.png"), key=lambda item: int(item.stem.rsplit("-", 1)[1]))
    need(len(files) == len(reader.pages), "Actual PDF render count differs from physical pages")
    pages = []
    for index, (page, image) in enumerate(zip(reader.pages, files), 1):
        stream = page.get_contents()
        text = page.extract_text() or ""
        need(len(text.encode("utf-8")) <= 64 * 1024, "Authored PDF text outside fixture bound")
        geometry = {"media_box": list(map(float, page.mediabox)), "crop_box": list(map(float, page.cropbox)), "rotation": page.rotation}
        need(all(math.isfinite(value) for key in ("media_box", "crop_box") for value in geometry[key]), "Invalid generated page geometry")
        pages.append({"page": index, "text": text, "text_sha256": sha256(text.encode("utf-8")).hexdigest(),
                      "content_sha256": sha256(stream.get_data() if stream is not None else b"").hexdigest(),
                      "geometry": geometry, "render": fidelity.png_record(image, index)})
    outline = []
    def walk(nodes, depth=0):
        need(depth <= 8, "Authored PDF outline unexpectedly deep")
        for node in nodes:
            if isinstance(node, list):
                walk(node, depth + 1)
            else:
                need(len(outline) < 16, "Authored PDF outline unexpectedly large")
                page = reader.get_destination_page_number(node)
                need(type(page) is int and 0 <= page < len(pages), "Generated PDF outline has invalid destination")
                outline.append({"title": node.title, "page": page, "depth": depth})
    walk(reader.outline)
    metadata = {str(key): str(value) for key, value in (reader.metadata or {}).items() if isinstance(value, str)}
    need(before == {"sha256": digest(path), "size_bytes": path.stat().st_size}, "PDF changed during physical/render observation")
    return {"path": str(path), **before, "page_count": len(pages), "pages": pages, "outline": outline,
            "metadata": metadata, "render_root": str(render_root), "render_process": process}


def reference_conversions(sources, work, calibre, renderer, *, render_root):
    """Independent native format->PDF references; default is diagnostic only."""
    result = {}
    render_root = Path(render_root)
    render_root.mkdir(parents=True, exist_ok=False)
    for fmt in ("epub", "azw3"):
        source = Path(sources[fmt])
        before = digest(source)
        result[fmt] = {}
        for profile in ("tablet", "default"):
            target = render_root / f"reference-{fmt}-{profile}.pdf"
            options = ["--output-profile", "tablet"] if profile == "tablet" else []
            observation = native([str(calibre), str(source), str(target), *options], work, timeout=180)
            observation["argv"] = observation["command"]  # Exact inherited fixture-record compatibility alias.
            need(observation["exit_code"] == 0 and target.is_file() and target.stat().st_size > 0, "Independent actual format reference conversion failed: " + json.dumps(observation))
            pdf = observe_pdf(target.resolve(), renderer, render_root / f"{fmt}-{profile}-pages")
            markers = re.findall(r"(?<![A-Z0-9-])WBS-PAGE-[0-9]{3}(?![0-9])", "\n".join(page["text"] for page in pdf["pages"]))
            need(markers == fixtures.CHAPTER_MARKERS, "Actual format reference lost or duplicated an authored chapter")
            result[fmt][profile] = {"input_format": fmt, "input_path": str(source), "input_sha256": before,
                                   "profile": profile, "options": options, "process": observation, "observation": pdf}
        need(digest(source) == before, "Reference conversion changed original ebook bytes")
    return result
