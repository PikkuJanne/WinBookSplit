"""Synthetic authored-family identity/closure guards, not Calibre acceptance."""
from copy import copy
import importlib.util
import io
from pathlib import Path
import struct
import sys
import tempfile
import unittest
import warnings
import zipfile

ROOT = Path(__file__).resolve().parents[2]
SPEC = importlib.util.spec_from_file_location("wbs_ebook_assets_units", ROOT / "tests/ebooks/assets.py")
assets = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = assets
SPEC.loader.exec_module(assets)


def rewrite(data, member, transform):
    buffer = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as original, zipfile.ZipFile(buffer, "w") as output:
        for info in original.infolist():
            raw = original.read(info.filename)
            output.writestr(info, transform(raw) if info.filename == member else raw)
    return buffer.getvalue()


def synthetic_mobi8():
    """Parser sample only: it deliberately has no authentic ebook payload."""
    header = bytearray(94)
    header[60:68] = b"BOOKMOBI"
    struct.pack_into(">H", header, 76, 2)
    struct.pack_into(">I", header, 78, 94)
    struct.pack_into(">I", header, 86, 374)
    zero = bytearray(280)
    zero[16:20] = b"MOBI"
    struct.pack_into(">I", zero, 20, 264)
    struct.pack_into(">I", zero, 36, 8)
    return bytes(header + zero + bytearray(32))


class EbookAssetGuards(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix="wbs-ebook-asset-unit-")
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.epub = assets.fixtures.authored_epub(self.root / "original.epub").read_bytes()

    def test_original_epub_has_closed_ordered_navigation(self):
        record = assets.inspect_epub_data(self.epub)
        self.assertEqual(record["spine"], ["c1", "c2", "c3"])
        self.assertEqual([row["title"] for row in record["navigation"]], assets.fixtures.CHAPTER_TITLES)
        self.assertEqual(record["external_references"], [])
        self.assertEqual(record["resource_scope"], "closed-authored-family")

    def test_remote_or_escaping_xml_references_are_rejected(self):
        for value in (b"https://example.invalid/style.css", b"//example.invalid/style.css",
                      b"file:///C:/outside.css", b"../../outside.css", b"/outside.css",
                      b"style.css?fetch=yes", b"style\\.css", b"missing.css"):
            with self.subTest(reference=value), self.assertRaises(ValueError):
                assets.inspect_epub_data(rewrite(self.epub, "OEBPS/chapter1.xhtml",
                                                lambda raw: raw.replace(b'href="style.css"', b'href="' + value + b'"')))

    def test_active_or_entity_xml_is_rejected(self):
        for addition in (b'<script>1</script>', b'<iframe src="chapter1.xhtml"/>',
                         b'<object data="chapter1.xhtml"/>', b'<p onclick="1">active</p>'):
            with self.subTest(addition=addition), self.assertRaises(ValueError):
                assets.inspect_epub_data(rewrite(self.epub, "OEBPS/chapter1.xhtml",
                                                lambda raw: raw.replace(b"</body>", addition + b"</body>")))
        with self.assertRaises(ValueError):
            assets.inspect_epub_data(rewrite(self.epub, "OEBPS/chapter1.xhtml",
                                            lambda raw: raw.replace(b"<html", b'<!DOCTYPE html [<!ENTITY x "outside">]><html', 1)))

    def test_external_css_and_inline_styles_are_rejected(self):
        for css in (b'@import "https://example.invalid/a.css";', b'body {background:url(//example.invalid/a)}',
                    b'body {background:u\\72l(//example.invalid/a)}'):
            with self.subTest(css=css), self.assertRaises(ValueError):
                assets.inspect_epub_data(rewrite(self.epub, "OEBPS/style.css", lambda raw: css))
        for addition in (b'<style>body {background:url(https://example.invalid/a)}</style>',
                         b'<p style="background:url(https://example.invalid/a)">hidden</p>'):
            with self.subTest(inline=addition), self.assertRaises(ValueError):
                assets.inspect_epub_data(rewrite(self.epub, "OEBPS/chapter1.xhtml",
                                                lambda raw: raw.replace(b"</body>", addition + b"</body>")))

    def test_duplicate_or_missing_archive_members_are_rejected(self):
        for duplicate in (True, False):
            buffer = io.BytesIO()
            with zipfile.ZipFile(io.BytesIO(self.epub)) as original, zipfile.ZipFile(buffer, "w") as output:
                for info in original.infolist():
                    if not duplicate and info.filename == "OEBPS/style.css":
                        continue
                    output.writestr(info, original.read(info.filename))
                if duplicate:
                    with warnings.catch_warnings():
                        warnings.simplefilter("ignore", UserWarning)
                        output.writestr("OEBPS/style.css", b"duplicate")
            with self.subTest(duplicate=duplicate), self.assertRaises(ValueError):
                assets.inspect_epub_data(buffer.getvalue())

    def test_mimetype_and_toc_order_are_required(self):
        for member, old, new in (("mimetype", b"application/epub+zip", b"application/pdf"),
                                 ("OEBPS/nav.xhtml", b"Chapter One", b"Wrong"),
                                 ("OEBPS/toc.ncx", b'content src="chapter1.xhtml"', b'content src="chapter2.xhtml"'),
                                 ("OEBPS/content.opf", b'itemref idref="c1"', b'itemref idref="c2"')):
            with self.subTest(member=member), self.assertRaises(ValueError):
                assets.inspect_epub_data(rewrite(self.epub, member, lambda raw: raw.replace(old, new)))
        buffer = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(self.epub)) as original, zipfile.ZipFile(buffer, "w") as output:
            for original_info in original.infolist():
                info = copy(original_info)
                info.compress_type = zipfile.ZIP_DEFLATED
                output.writestr(info, original.read(info.filename))
        with self.assertRaises(ValueError):
            assets.inspect_epub_data(buffer.getvalue())

    def test_canary_allows_exact_two_loopback_urls_only(self):
        source = self.root / "original.epub"
        record = assets.authored_canary_epub(self.root / "canary.epub", "http://127.0.0.1:12345/owned/", source)
        data = (self.root / "canary.epub").read_bytes()
        self.assertEqual(len(record["resource_closure"]["external_references"]), 2)
        self.assertEqual(assets.inspect_epub_data(data, allowed_urls=record["urls"])["sha256"], record["sha256"])
        with self.assertRaises(ValueError):
            assets.inspect_epub_data(data)
        with self.assertRaises(ValueError):
            assets.inspect_epub_data(data, allowed_urls=[url.replace("image.png", "different.png") for url in record["urls"]])

    def test_canary_rejects_nonowned_endpoints(self):
        for base in ("https://127.0.0.1:12345/owned/", "http://localhost:12345/owned/",
                     "http://example.invalid:12345/owned/", "http://user@127.0.0.1:12345/owned/",
                     "http://127.0.0.1:12345/owned/?query=1"):
            with self.subTest(base=base), self.assertRaises(ValueError):
                assets.authored_canary_epub(self.root / "refused.epub", base, self.root / "original.epub")
        self.assertFalse((self.root / "refused.epub").exists())

    def test_synthetic_mobi8_header_parser_controls_are_distinct_from_native_provenance(self):
        record = assets.inspect_azw3_data(synthetic_mobi8())
        self.assertEqual((record["palmdb_identity"], record["mobi_version"], record["encryption_type"]), ("BOOKMOBI", 8, 0))
        self.assertEqual(record["record_count"], 2)

    def test_renamed_epub_is_never_azw3(self):
        with self.assertRaises(ValueError):
            assets.inspect_azw3_data(self.epub)
        forged = bytearray(self.epub)
        forged[60:68] = b"BOOKMOBI"
        with self.assertRaises(ValueError):
            assets.inspect_azw3_data(bytes(forged))

    def test_mobi_header_version_encryption_and_bounds_reject(self):
        defects = ((60, b"FAKEMOBI"), (76, struct.pack(">H", 1)), (76, struct.pack(">H", 4097)),
                   (78, struct.pack(">I", 70)), (86, struct.pack(">I", 94)),
                   (86, struct.pack(">I", 10000)), (94 + 16, b"FAKE"),
                   (94 + 20, struct.pack(">I", 111)), (94 + 20, struct.pack(">I", 501)),
                   (94 + 36, struct.pack(">I", 6)), (94 + 12, struct.pack(">H", 1)))
        for position, replacement in defects:
            data = bytearray(synthetic_mobi8())
            data[position:position + len(replacement)] = replacement
            with self.subTest(position=position, replacement=replacement), self.assertRaises(ValueError):
                assets.inspect_azw3_data(bytes(data))
        for data in (b"", synthetic_mobi8()[:127], b"x" * (assets.MAX_EBOOK_BYTES + 1)):
            with self.subTest(size=len(data)), self.assertRaises(ValueError):
                assets.inspect_azw3_data(data)


if __name__ == "__main__":
    unittest.main(verbosity=2)
